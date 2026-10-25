# /// script
# requires-python = ">=3.14"
# dependencies = [
#     "azure-identity",
#     "requests",
#     "xlsxwriter",
# ]
# ///
"""Export Entra ID group membership to Excel (default) or CSV via Microsoft Graph.

Usage:
    uv run group_members.py "My Group"
    uv run group_members.py "My Group" --transitive
    uv run group_members.py "My Group" --format csv
    uv run group_members.py 0638c5e9-1424-4da6-9538-a4f8ea8799ac -o members.xlsx

Auth uses the Azure CLI's signed-in account, so run `az login` first.

Permissions: GroupMember.Read.All, User.Read.All and AuditLog.Read.All.
Sign-in activity also needs an Entra ID P1/P2 licence in the tenant and a
directory role such as Reports Reader, Security Reader or Global Reader; without
them the export still runs, but lastSignInDate is left blank.

lastSignInDate is the last *successful* sign-in, which Entra has only recorded
since December 2023, so accounts not used since then show blank.
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import re
import sys
import time
from datetime import date, datetime
from itertools import batched
from pathlib import Path
from typing import TYPE_CHECKING, Any
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
from azure.identity import AzureCliCredential
from xlsxwriter import Workbook

if TYPE_CHECKING:
    from collections.abc import Iterator, Mapping

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPE = "https://graph.microsoft.com/.default"
TIMEOUT = 30
BATCH_SIZE = 20  # Graph $batch maximum
LOCAL_TZ = ZoneInfo("Australia/Sydney")
GUID_RE = re.compile(r"^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$", re.IGNORECASE)
USER_SELECT_BASIC = "id,displayName,userPrincipalName,mail,userType,accountEnabled"
USER_SELECT = f"{USER_SELECT_BASIC},signInActivity"
COLUMNS = [
    "type",
    "displayName",
    "userPrincipalName",
    "mail",
    "accountType",
    "accountStatus",
    "lastSignInDate",
    "id",
]
HTTP_FORBIDDEN = 403
RETRY_STATUSES = frozenset({429, 503, 504})  # throttled or transient
DEFAULT_RETRY_SECONDS = 5
MAX_ATTEMPTS = 6
MAX_COL_PIXELS = 400
PROGRESS_EVERY = 10  # batches between user lookup progress lines


def log(msg: str) -> None:
    """Print a timestamped progress line to stderr, keeping stdout clean."""
    print(f"{time.strftime('%H:%M:%S')} {msg}", file=sys.stderr)  # noqa: T201


def token_identity(token: str) -> str:
    """Return the account and tenant a JWT was issued for, for display only."""
    payload = token.split(".")[1]
    claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
    account = claims.get("upn") or claims.get("unique_name") or claims.get("appid", "unknown")
    return f"{account} (tenant {claims.get('tid', 'unknown')})"


def graph_session() -> requests.Session:
    """Return a requests session carrying a Graph bearer token."""
    token = AzureCliCredential().get_token(SCOPE).token
    log(f"Authenticated as {token_identity(token)}")
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}"})
    return session


def retry_delay(headers: Mapping[str, str] | None) -> int:
    """Return the Retry-After delay in seconds, or a default if absent."""
    return int((headers or {}).get("Retry-After", DEFAULT_RETRY_SECONDS))


def send(session: requests.Session, method: str, url: str, **kwargs: Any) -> requests.Response:  # noqa: ANN401
    """Send a Graph request, retrying throttled and transient failures."""
    for _ in range(MAX_ATTEMPTS - 1):
        resp = session.request(method, url, timeout=TIMEOUT, **kwargs)
        if resp.status_code not in RETRY_STATUSES:
            break
        time.sleep(retry_delay(resp.headers))
    else:
        resp = session.request(method, url, timeout=TIMEOUT, **kwargs)
    resp.raise_for_status()
    return resp


def paged(session: requests.Session, url: str) -> Iterator[dict[str, Any]]:
    """Yield every item from a Graph collection, following @odata.nextLink."""
    next_url: str | None = url
    while next_url:
        body = send(session, "GET", next_url).json()
        yield from body.get("value", [])
        next_url = body.get("@odata.nextLink")


def resolve_group(session: requests.Session, group: str) -> dict[str, Any]:
    """Accept a group object ID or an exact display name and return its id and displayName."""
    if GUID_RE.match(group):
        return send(session, "GET", f"{GRAPH}/groups/{group}?$select=id,displayName").json()
    # Encode the whole filter so &, + and # in a name can't break the query string.
    odata_filter = quote("displayName eq '{}'".format(group.replace("'", "''")), safe="")
    matches = list(paged(session, f"{GRAPH}/groups?$filter={odata_filter}&$select=id,displayName"))
    if len(matches) != 1:
        msg = f"Expected exactly one group named {group!r}, found {len(matches)}"
        raise SystemExit(msg)
    return matches[0]


def fetch_users(session: requests.Session, ids: list[str]) -> dict[str, dict[str, Any]]:
    """Fetch full user records (incl. signInActivity) via $batch, 20 at a time.

    signInActivity can't be selected on the group members endpoint, so each
    user is looked up individually; batching keeps that to one call per 20.

    Graph rejects the whole lookup with 403 if signInActivity isn't allowed
    (no P1/P2 licence, permission or role), so on the first 403 this drops
    signInActivity for the rest of the run rather than losing every user.
    """
    users: dict[str, dict[str, Any]] = {}
    select = USER_SELECT
    log(f"Looking up details for {len(ids)} users")
    for n, chunk in enumerate(batched(ids, BATCH_SIZE, strict=False), start=1):
        select = fetch_chunk(session, chunk, select, users)
        if n % PROGRESS_EVERY == 0:
            log(f"  {min(n * BATCH_SIZE, len(ids))}/{len(ids)} users looked up")
    return users


def fetch_chunk(
    session: requests.Session,
    chunk: tuple[str, ...],
    select: str,
    users: dict[str, dict[str, Any]],
) -> str:
    """Look up one $batch of users into `users`; return the $select to use next."""
    pending = {str(i): uid for i, uid in enumerate(chunk)}
    for _ in range(MAX_ATTEMPTS):
        batch = [{"id": k, "method": "GET", "url": f"/users/{uid}?$select={select}"} for k, uid in pending.items()]
        resp = send(session, "POST", f"{GRAPH}/$batch", json={"requests": batch})
        retry_after = 0
        denied = ""
        for r in resp.json()["responses"]:
            body = r.get("body") or {}
            err = body.get("error", {}).get("message", r["status"])
            if r["status"] in RETRY_STATUSES:
                retry_after = max(retry_after, retry_delay(r.get("headers")))
            elif r["status"] == HTTP_FORBIDDEN and select == USER_SELECT:
                denied = err  # left pending, retried without signInActivity
            elif r["status"] == requests.codes.ok:
                pending.pop(r["id"])
                users[body["id"]] = body
            else:
                pending.pop(r["id"])
                log(f"Warning: user lookup failed: {err}")
        if denied:
            select = USER_SELECT_BASIC
            log(f"Warning: no sign-in activity, lastSignInDate left blank: {denied}")
        if not pending:
            return select
        log(f"Throttled by Graph, waiting {retry_after}s")
        time.sleep(retry_after)
    log(f"Warning: {len(pending)} user lookups still throttled, giving up")
    return select


def sign_in_date(user: dict[str, Any]) -> date | None:
    """Return the last successful sign-in as a local date, or None if never."""
    stamp = (user.get("signInActivity") or {}).get("lastSuccessfulSignInDateTime")
    if not stamp:
        return None
    return datetime.fromisoformat(stamp).astimezone(LOCAL_TZ).date()


def build_rows(
    session: requests.Session,
    group_id: str,
    *,
    transitive: bool,
) -> list[dict[str, Any]]:
    """Return one row per group member; user-only columns blank for others."""
    relation = "transitiveMembers" if transitive else "members"
    url = f"{GRAPH}/groups/{group_id}/{relation}?$select=id,displayName&$top=999"
    members = list(paged(session, url))

    user_ids = [m["id"] for m in members if m.get("@odata.type") == "#microsoft.graph.user"]
    others = len(members) - len(user_ids)
    log(f"Found {len(members)} {'transitive ' if transitive else ''}members: {len(user_ids)} users, {others} other")
    users = fetch_users(session, user_ids)

    rows = []
    for m in members:
        row: dict[str, Any] = dict.fromkeys(COLUMNS, "")
        row["type"] = m.get("@odata.type", "").removeprefix("#microsoft.graph.")
        row["displayName"] = m.get("displayName") or ""
        row["id"] = m["id"]
        if user := users.get(m["id"]):
            row["userPrincipalName"] = user.get("userPrincipalName") or ""
            row["mail"] = user.get("mail") or ""
            row["accountType"] = "Guest" if user.get("userType") == "Guest" else "Standard"
            row["accountStatus"] = "Enabled" if user.get("accountEnabled") else "Disabled"
            row["lastSignInDate"] = sign_in_date(user) or ""
        rows.append(row)

    return sorted(rows, key=lambda r: (r["type"], r["displayName"].casefold()))


def write_csv(rows: list[dict[str, Any]], output: Path) -> None:
    """Write rows as plain CSV (dates as YYYY-MM-DD)."""
    # BOM so Excel reads the file as UTF-8 and doesn't garble non-ASCII names.
    with output.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(rows: list[dict[str, Any]], output: Path) -> None:
    """Write rows as an Excel table: filter buttons, banding, frozen header."""
    with Workbook(output) as wb:
        ws = wb.add_worksheet("Members")
        # Real dates so Excel sorts and filters them by date, not as text.
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
        columns = [{"header": c, "format": date_fmt} if c == "lastSignInDate" else {"header": c} for c in COLUMNS]
        ws.add_table(
            0,
            0,
            max(len(rows), 1),  # header-only table still needs one (blank) row
            len(COLUMNS) - 1,
            {
                "name": "Members",
                "style": "Table Style Medium 2",
                "columns": columns,
                "data": [[row[c] for c in COLUMNS] for row in rows],
            },
        )
        ws.freeze_panes(1, 0)
        ws.autofit(max_width=MAX_COL_PIXELS)


def main() -> None:
    """Parse arguments and write the export."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("group", help="Group display name or object ID")
    parser.add_argument(
        "--transitive",
        action="store_true",
        help="Include members of nested groups",
    )
    parser.add_argument(
        "--format",
        choices=["xlsx", "csv"],
        default="xlsx",
        help="Output format (default: xlsx)",
    )
    parser.add_argument("-o", "--output", type=Path, help="Output file path")
    args = parser.parse_args()

    started = time.monotonic()
    log(f"Started: exporting members of {args.group!r}")
    session = graph_session()
    group = resolve_group(session, args.group)
    log(f"Found group {group['displayName']!r} ({group['id']})")
    rows = build_rows(session, group["id"], transitive=args.transitive)

    stem = re.sub(r"[^\w-]+", "_", args.group)
    output: Path = args.output or Path(f"{stem}_members.{args.format}")
    if args.format == "csv":
        write_csv(rows, output)
    else:
        write_xlsx(rows, output)

    log(f"Wrote {len(rows)} members to {output} in {time.monotonic() - started:.1f}s")


if __name__ == "__main__":
    main()
