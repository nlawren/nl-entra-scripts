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
"""Export Entra ID guest users from one email domain to Excel (default) or CSV via Microsoft Graph.

Example runs:
    uv run guest_users.py removed.example.com
    uv run guest_users.py removed.example.com --format csv
    uv run guest_users.py removed.example.com -o guests.xlsx

Auth uses the Azure CLI's signed-in account, so run `az login` first.

Permissions: User.Read.All and AuditLog.Read.All, with the same sign-in
activity caveats as group_members.py; without them lastSignInDate is blank.

A guest matches if its mail or its UPN (first.last_domain#EXT#@tenant) is in
the domain. Subdomains don't match: example.com won't match a.example.com.
"""

from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any

import requests
from group_members import (
    COLUMNS,
    GRAPH,
    HTTP_FORBIDDEN,
    PROGRESS,
    USER_SELECT,
    USER_SELECT_BASIC,
    add_output_args,
    graph_error,
    graph_session,
    log,
    make_parser,
    paged,
    running,
    send,
    sign_in_date,
    summarise,
    warn,
    write_output,
)

GUEST_FILTER = "userType eq 'Guest'"


def in_domain(user: dict[str, Any], domain: str) -> bool:
    """Return True if the guest's mail or external UPN is in `domain`."""
    mail = (user.get("mail") or "").casefold()
    upn = (user.get("userPrincipalName") or "").casefold()
    # Guest UPNs encode the home address as local_domain#EXT#@tenant.onmicrosoft.com.
    external = upn.partition("#ext#")[0] if "#ext#" in upn else ""
    return mail.endswith(f"@{domain}") or external.endswith(f"_{domain}")


def count_guests(session: requests.Session) -> int | None:
    """Return the tenant's guest count for the progress bar, or None if unavailable.

    $count is an advanced query, so it's approximate (eventually consistent).
    """
    try:
        resp = send(
            session,
            "GET",
            f"{GRAPH}/users/$count?$filter={GUEST_FILTER}",
            headers={"ConsistencyLevel": "eventual"},
        )
        return int(resp.text)
    except requests.HTTPError, ValueError:
        return None


def list_guests(session: requests.Session, select: str, total: int | None) -> list[dict[str, Any]]:
    """Return every guest user in the tenant with the given $select."""
    url = f"{GRAPH}/users?$filter={GUEST_FILTER}&$select={select}&$top=999"
    task = PROGRESS.add_task("Scanning guests", total=total)
    guests = []
    try:
        for user in paged(session, url):
            guests.append(user)
            PROGRESS.advance(task)
    except BaseException:
        PROGRESS.remove_task(task)
        raise
    PROGRESS.update(task, total=len(guests), completed=len(guests))  # the count was approximate
    return guests


def fetch_guests(session: requests.Session) -> tuple[list[dict[str, Any]], bool]:
    """List all guests, dropping signInActivity if Graph won't return it.

    Graph rejects the whole listing with 403 if signInActivity isn't allowed
    (no P1/P2 licence, permission or role), so retry once without it.

    Returns the guests and whether sign-in activity was available.
    """
    total = count_guests(session)
    try:
        return list_guests(session, USER_SELECT, total), True
    except requests.HTTPError as e:
        if e.response is None or e.response.status_code != HTTP_FORBIDDEN:
            raise
        warn(f"no sign-in activity, lastSignInDate left blank: {graph_error(e)}")
        return list_guests(session, USER_SELECT_BASIC, total), False


def build_rows(guests: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return one row per guest in the same columns as group_members.py."""
    rows = []
    for user in guests:
        row: dict[str, Any] = dict.fromkeys(COLUMNS, "")
        row["type"] = "user"
        row["displayName"] = user.get("displayName") or ""
        row["userPrincipalName"] = user.get("userPrincipalName") or ""
        row["mail"] = user.get("mail") or ""
        row["accountType"] = "Guest"
        row["accountStatus"] = "Enabled" if user.get("accountEnabled") else "Disabled"
        row["lastSignInDate"] = sign_in_date(user) or ""
        row["id"] = user["id"]
        rows.append(row)
    return sorted(rows, key=lambda r: r["displayName"].casefold())


def main() -> None:
    """Parse arguments and write the export."""
    parser = make_parser(__doc__)
    parser.add_argument("domain", help="Email domain, e.g. removed.example.com")
    add_output_args(parser)
    args = parser.parse_args()
    domain = args.domain.removeprefix("@").casefold()

    started = time.monotonic()
    log(f"Started: exporting guests from {domain!r}")
    with running():
        session = graph_session()
        guests, sign_ins = fetch_guests(session)
    matches = [g for g in guests if in_domain(g, domain)]
    log(f"Scanned {len(guests)} guests, {len(matches)} from {domain!r}")
    rows = build_rows(matches)

    stem = re.sub(r"[^\w-]+", "_", domain)
    output: Path = args.output or Path(f"{stem}_guests.{args.format}")
    write_output(rows, output, args.format)
    log(f"Wrote {len(rows)} guests to {output}", style="green")
    summarise(rows, noun="guests", sign_ins=sign_ins, started=started)


if __name__ == "__main__":
    main()
