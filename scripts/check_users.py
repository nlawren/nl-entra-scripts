# /// script
# requires-python = ">=3.14"
# dependencies = [
#     "azure-identity",
#     "requests",
#     "rich",
#     "rich-argparse",
#     "xlsxwriter",
# ]
#
# [tool.ty.environment]
# extra-paths = ["."]  # resolve group_members, as `uv run` does
# ///
"""Check that the users listed in a CSV still exist in Entra ID and whether they're enabled.

Example runs:
    uv run check_users.py powerapps_users.csv
    uv run check_users.py powerapps_users.csv --format csv
    uv run check_users.py powerapps_users.csv -o checked.xlsx

Auth uses the Azure CLI's signed-in account, so run `az login` first.

Permissions: User.Read.All and AuditLog.Read.All, with the same sign-in
activity caveats as group_members.py; without them lastSignInDate is blank.

The CSV needs an 'Object id' column, a 'User principle name' column, or both
(any case; 'User principal name' works too). Each user is looked up by object ID
because that never changes, falling back to the UPN if there's no usable ID or
the ID isn't found.

accountStatus is Enabled, Disabled, Not found, or Lookup failed. Lookup failed
means Graph gave no answer, so it says nothing about whether the account exists.
The notes column flags a UPN that has changed, or one that now belongs to a
different account from the object ID in the CSV.
"""

from __future__ import annotations

import csv
import time
from itertools import batched
from pathlib import Path
from typing import Any
from urllib.parse import quote

import requests
from group_members import (
    BATCH_SIZE,
    COLUMNS,
    CONSOLE,
    GUID_RE,
    PROGRESS,
    USER_SELECT,
    add_output_args,
    batch_error,
    batch_get,
    fail,
    graph_session,
    log,
    make_parser,
    running,
    sign_in_date,
    warn,
    write_output,
)
from rich.table import Table

ID_HEADER = "object id"
UPN_HEADERS = ("user principle name", "user principal name")
OUTPUT_COLUMNS = [*COLUMNS, "notes"]
NOT_FOUND = "Not found"
LOOKUP_FAILED = "Lookup failed"
# Accounts needing attention first, so they top the sheet.
STATUS_ORDER = {NOT_FOUND: 0, LOOKUP_FAILED: 1, "Disabled": 2, "Enabled": 3}

type Result = tuple[int, dict[str, Any]]


def read_entries(path: Path) -> list[tuple[str, str]]:
    """Return an (object id, UPN) pair per data row, dropping rows with neither."""
    try:
        with path.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            headers = {h.strip().casefold(): h for h in reader.fieldnames or [] if h}
            id_col = headers.get(ID_HEADER)
            upn_col = next((headers[h] for h in UPN_HEADERS if h in headers), None)
            if not id_col and not upn_col:
                fail(f"{path} has no 'Object id' or 'User principle name' column, found: {list(headers.values())}")
            entries = [
                ((row[id_col] or "").strip() if id_col else "", (row[upn_col] or "").strip() if upn_col else "")
                for row in reader
            ]
    except OSError as e:
        fail(f"Can't read {path}: {e}")
    except UnicodeDecodeError:
        fail(f"{path} isn't UTF-8, save it from Excel as 'CSV UTF-8' and try again")
    kept = [entry for entry in entries if any(entry)]
    if len(kept) != len(entries):
        warn(f"skipped {len(entries) - len(kept)} rows with no object ID or UPN")
    return kept


def upn_path(upn: str) -> str:
    """Return the Graph path that finds a user by UPN.

    A $filter rather than /users/{upn}: guest UPNs contain #, and apostrophes or
    a leading $ need special handling as a path key but not as a filter value.
    """
    odata_filter = quote("userPrincipalName eq '{}'".format(upn.replace("'", "''")), safe="")
    return f"/users?$filter={odata_filter}"


def lookup(
    session: requests.Session,
    paths: dict[str, str],
    select: str,
    description: str,
) -> tuple[dict[str, Result], str]:
    """Run `paths` through $batch; return the results and the $select to use next."""
    results: dict[str, Result] = {}
    if not paths:
        return results, select
    task = PROGRESS.add_task(description, total=len(paths))
    for chunk in batched(paths.items(), BATCH_SIZE, strict=False):
        found, select = batch_get(session, dict(chunk), select)
        results |= found
        PROGRESS.advance(task, len(chunk))
    return results, select


def classify(result: Result | None) -> tuple[dict[str, Any] | None, str]:
    """Return a lookup's (user, error); neither means the account doesn't exist.

    Only a 404 or an empty filter result counts as not found. Anything else, such
    as a 403 or still being throttled, is an error: the account may well exist.
    """
    if result is None:
        return None, "Still throttled by Graph"
    status, body = result
    if status == requests.codes.not_found:
        return None, ""
    if status != requests.codes.ok:
        return None, batch_error(status, body)
    if "value" not in body:  # by object ID
        return body, ""
    return (body["value"][0] if body["value"] else None), ""  # by UPN filter


def fetch_results(session: requests.Session, entries: list[tuple[str, str]]) -> tuple[dict[str, Result], bool]:
    """Look up every entry; return the final results by row index and sign-in availability."""
    by_id: dict[str, str] = {}
    by_upn: dict[str, str] = {}
    for i, (oid, upn) in enumerate(entries):
        if GUID_RE.match(oid):
            by_id[str(i)] = f"/users/{oid}"
        elif upn:
            by_upn[str(i)] = upn_path(upn)

    results, select = lookup(session, by_id, USER_SELECT, "Checking object IDs")
    # Object IDs that have definitely gone get a second chance by UPN.
    for key in by_id:
        upn = entries[int(key)][1]
        if upn and classify(results.get(key)) == (None, ""):
            by_upn[key] = upn_path(upn)
    upn_results, select = lookup(session, by_upn, select, "Checking UPNs")
    results |= upn_results
    return results, select == USER_SELECT


def build_row(oid: str, upn: str, result: Result | None) -> dict[str, Any]:
    """Return the output row for one CSV entry, keeping the CSV's values if there's no account."""
    row: dict[str, Any] = dict.fromkeys(OUTPUT_COLUMNS, "")
    row["type"] = "user"
    row["userPrincipalName"] = upn
    row["id"] = oid
    if not upn and not GUID_RE.match(oid):
        row["accountStatus"] = LOOKUP_FAILED
        row["notes"] = "Object ID isn't a valid GUID and there's no UPN"
        return row

    user, error = classify(result)
    if error:
        row["accountStatus"] = LOOKUP_FAILED
        row["notes"] = error
    elif user is None:
        row["accountStatus"] = NOT_FOUND
    else:
        row["displayName"] = user.get("displayName") or ""
        row["userPrincipalName"] = user.get("userPrincipalName") or ""
        row["mail"] = user.get("mail") or ""
        row["accountType"] = "Guest" if user.get("userType") == "Guest" else "Standard"
        row["accountStatus"] = "Enabled" if user.get("accountEnabled") else "Disabled"
        row["lastSignInDate"] = sign_in_date(user) or ""
        row["id"] = user["id"]
        if oid and oid.casefold() != user["id"].casefold():
            row["notes"] = f"CSV object ID {oid} not found, this UPN belongs to a different account"
        elif upn and upn.casefold() != row["userPrincipalName"].casefold():
            row["notes"] = f"UPN in the CSV was {upn}"
    return row


def summarise(rows: list[dict[str, Any]], *, sign_ins: bool, started: float) -> None:
    """Print an end-of-run summary table of the check."""
    counts = {status: sum(r["accountStatus"] == status for r in rows) for status in STATUS_ORDER}
    found = [r for r in rows if r["accountStatus"] in {"Enabled", "Disabled"}]
    table = Table(title="Summary", title_justify="left", show_header=False)
    table.add_column(style="bold")
    table.add_column(justify="right")
    table.add_row("Users checked", str(len(rows)))
    table.add_row("Enabled", str(counts["Enabled"]), style="green")
    table.add_row("Disabled", str(counts["Disabled"]), style="red")
    table.add_row(NOT_FOUND, str(counts[NOT_FOUND]), style="red")
    if counts[LOOKUP_FAILED]:
        table.add_row(LOOKUP_FAILED, str(counts[LOOKUP_FAILED]), style="bold yellow")
    if flagged := sum(bool(r["notes"]) for r in found):
        table.add_row("Found with a note", str(flagged), style="yellow")
    if sign_ins:
        table.add_row("No sign-in recorded", str(sum(not r["lastSignInDate"] for r in found)), style="yellow")
    else:
        table.add_row("No sign-in recorded", "unavailable", style="dim")
    table.add_row("Time taken", f"{time.monotonic() - started:.1f}s")
    CONSOLE.print(table)


def main() -> None:
    """Parse arguments and write the check results."""
    parser = make_parser(__doc__)
    parser.add_argument("csv", type=Path, help="CSV file listing the users to check")
    add_output_args(parser)
    args = parser.parse_args()

    started = time.monotonic()
    entries = read_entries(args.csv)
    log(f"Started: checking {len(entries)} users from {args.csv}")
    with running():
        session = graph_session()
        results, sign_ins = fetch_results(session, entries)
    rows = [build_row(oid, upn, results.get(str(i))) for i, (oid, upn) in enumerate(entries)]
    rows.sort(key=lambda r: (STATUS_ORDER[r["accountStatus"]], r["userPrincipalName"].casefold()))

    output: Path = args.output or Path(f"{args.csv.stem}_checked.{args.format}")
    write_output(rows, output, args.format, OUTPUT_COLUMNS)
    log(f"Wrote {len(rows)} users to {output}", style="green")
    if failed := sum(r["accountStatus"] == LOOKUP_FAILED for r in rows):
        warn(f"{failed} lookups failed, so those accounts may still exist: see the notes column")
    summarise(rows, sign_ins=sign_ins, started=started)


if __name__ == "__main__":
    main()
