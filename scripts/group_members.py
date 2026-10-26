# /// script
# requires-python = ">=3.14"
# dependencies = [
#     "azure-identity",
#     "requests",
#     "rich",
#     "rich-argparse",
#     "xlsxwriter",
# ]
# ///
"""Export Entra ID group membership to Excel (default) or CSV via Microsoft Graph.

Example runs:
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
import logging
import re
import sys
import time
from contextlib import contextmanager
from datetime import date, datetime
from itertools import batched
from pathlib import Path
from typing import TYPE_CHECKING, Any, NoReturn
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AzureCliCredential
from rich.console import Console
from rich.progress import BarColumn, MofNCompleteColumn, Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
from rich.table import Table
from rich_argparse import RawDescriptionRichHelpFormatter
from xlsxwriter import Workbook

if TYPE_CHECKING:
    from collections.abc import Generator, Iterator, Mapping

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

# Everything goes to stderr, keeping stdout clean. Colour and live progress
# switch themselves off when stderr isn't a terminal (piped or redirected).
CONSOLE = Console(stderr=True, log_path=False, log_time_format="%H:%M:%S")
PROGRESS = Progress(
    SpinnerColumn(),
    TextColumn("{task.description}"),
    BarColumn(),
    MofNCompleteColumn(),
    TimeElapsedColumn(),
    console=CONSOLE,
)
# azure-identity logs its own warning on a failed sign-in; graph_session reports it instead.
logging.getLogger("azure.identity").setLevel(logging.ERROR)


def log(msg: str, style: str | None = None) -> None:
    """Print a timestamped progress line; markup off so names can't break it."""
    if CONSOLE.is_terminal:
        CONSOLE.log(msg, style=style, markup=False)
    else:
        # Plain lines when piped: rich would wrap and pad them to 80 columns.
        print(f"{time.strftime('%H:%M:%S')} {msg}", file=sys.stderr)  # noqa: T201


def warn(msg: str) -> None:
    """Log a warning in yellow."""
    log(f"Warning: {msg}", style="yellow")


def fail(msg: str) -> NoReturn:
    """Log an error in red and exit with status 1."""
    log(f"Error: {msg}", style="bold red")
    sys.exit(1)


def graph_error(e: requests.HTTPError) -> str:
    """Return Graph's own message for a failed request."""
    resp = e.response
    if resp is None:
        return str(e)
    try:
        detail = resp.json()["error"]["message"]
    except ValueError, KeyError:
        detail = resp.text
    return f"Graph returned {resp.status_code} {resp.reason}: {detail}"


@contextmanager
def running() -> Generator[None]:
    """Show live progress for the run and report Graph failures cleanly."""
    with PROGRESS:
        try:
            yield
        except requests.HTTPError as e:
            fail(graph_error(e))


def make_parser(doc: str) -> argparse.ArgumentParser:
    """Return a parser whose --help shows the script's docstring, in colour."""
    return argparse.ArgumentParser(description=doc, formatter_class=RawDescriptionRichHelpFormatter)


def add_output_args(parser: argparse.ArgumentParser) -> None:
    """Add the --format and --output options shared by every export."""
    parser.add_argument(
        "--format",
        choices=["xlsx", "csv"],
        default="xlsx",
        help="Output format (default: %(default)s)",
    )
    parser.add_argument("-o", "--output", type=Path, help="Output file path (default: named after the input)")


def token_identity(token: str) -> str:
    """Return the account and tenant a JWT was issued for, for display only."""
    payload = token.split(".")[1]
    claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
    account = claims.get("upn") or claims.get("unique_name") or claims.get("appid", "unknown")
    return f"{account} (tenant {claims.get('tid', 'unknown')})"


def graph_session() -> requests.Session:
    """Return a requests session carrying a Graph bearer token."""
    try:
        token = AzureCliCredential().get_token(SCOPE).token
    except ClientAuthenticationError as e:
        fail(f"Azure CLI sign-in failed, run `az login` first: {str(e).splitlines()[0]}")
    log(f"Authenticated as {token_identity(token)}", style="green")
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}"})
    return session


def retry_delay(headers: Mapping[str, str] | None) -> int:
    """Return the Retry-After delay in seconds, or a default if absent."""
    return int((headers or {}).get("Retry-After", DEFAULT_RETRY_SECONDS))


def wait(seconds: int) -> None:
    """Sleep out a Graph throttle, counting it down on the progress display."""
    log(f"Throttled by Graph, waiting {seconds}s", style="yellow")
    task = PROGRESS.add_task("[yellow]Throttled, waiting (s)", total=seconds)
    for _ in range(seconds):
        time.sleep(1)
        PROGRESS.advance(task)
    PROGRESS.remove_task(task)


def send(session: requests.Session, method: str, url: str, **kwargs: Any) -> requests.Response:  # noqa: ANN401
    """Send a Graph request, retrying throttled and transient failures."""
    for _ in range(MAX_ATTEMPTS - 1):
        resp = session.request(method, url, timeout=TIMEOUT, **kwargs)
        if resp.status_code not in RETRY_STATUSES:
            break
        wait(retry_delay(resp.headers))
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
        fail(f"Expected exactly one group named {group!r}, found {len(matches)}")
    return matches[0]


def fetch_users(session: requests.Session, ids: list[str]) -> tuple[dict[str, dict[str, Any]], bool]:
    """Fetch full user records (incl. signInActivity) via $batch, 20 at a time.

    signInActivity can't be selected on the group members endpoint, so each
    user is looked up individually; batching keeps that to one call per 20.

    Graph rejects the whole lookup with 403 if signInActivity isn't allowed
    (no P1/P2 licence, permission or role), so on the first 403 this drops
    signInActivity for the rest of the run rather than losing every user.

    Returns the users by id and whether sign-in activity was available.
    """
    users: dict[str, dict[str, Any]] = {}
    select = USER_SELECT
    task = PROGRESS.add_task("Looking up users", total=len(ids))
    for chunk in batched(ids, BATCH_SIZE, strict=False):
        select = fetch_chunk(session, chunk, select, users)
        PROGRESS.advance(task, len(chunk))
    return users, select == USER_SELECT


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
                warn(f"user lookup failed: {err}")
        if denied:
            select = USER_SELECT_BASIC
            warn(f"no sign-in activity, lastSignInDate left blank: {denied}")
        if not pending:
            return select
        if retry_after:  # else only the signInActivity denials are pending
            wait(retry_after)
    warn(f"{len(pending)} user lookups still throttled, giving up")
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
) -> tuple[list[dict[str, Any]], bool]:
    """Return one row per group member, and whether sign-in activity was available.

    User-only columns are left blank for members that aren't users.
    """
    relation = "transitiveMembers" if transitive else "members"
    url = f"{GRAPH}/groups/{group_id}/{relation}?$select=id,displayName&$top=999"
    task = PROGRESS.add_task("Listing members", total=None)
    members = []
    for m in paged(session, url):
        members.append(m)
        PROGRESS.advance(task)
    PROGRESS.update(task, total=len(members))

    user_ids = [m["id"] for m in members if m.get("@odata.type") == "#microsoft.graph.user"]
    others = len(members) - len(user_ids)
    log(f"Found {len(members)} {'transitive ' if transitive else ''}members: {len(user_ids)} users, {others} other")
    users, sign_ins = fetch_users(session, user_ids)

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

    return sorted(rows, key=lambda r: (r["type"], r["displayName"].casefold())), sign_ins


def write_output(rows: list[dict[str, Any]], output: Path, fmt: str) -> None:
    """Write rows to `output` in the chosen format."""
    if fmt == "csv":
        write_csv(rows, output)
    else:
        write_xlsx(rows, output)


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


def summarise(rows: list[dict[str, Any]], *, noun: str, sign_ins: bool, started: float) -> None:
    """Print an end-of-run summary table of what was exported."""
    users = [r for r in rows if r["type"] == "user"]
    table = Table(title="Summary", title_justify="left", show_header=False)
    table.add_column(style="bold")
    table.add_column(justify="right")
    table.add_row(noun.capitalize(), str(len(rows)))
    if len(users) != len(rows):
        table.add_row("Non-user members", str(len(rows) - len(users)))
    guests = sum(r["accountType"] == "Guest" for r in users)
    if guests != len(rows):
        table.add_row("Guests", str(guests))
    table.add_row("Enabled", str(sum(r["accountStatus"] == "Enabled" for r in users)), style="green")
    table.add_row("Disabled", str(sum(r["accountStatus"] == "Disabled" for r in users)), style="red")
    if sign_ins:
        table.add_row("No sign-in recorded", str(sum(not r["lastSignInDate"] for r in users)), style="yellow")
    else:
        table.add_row("No sign-in recorded", "unavailable", style="dim")
    table.add_row("Time taken", f"{time.monotonic() - started:.1f}s")
    CONSOLE.print(table)


def main() -> None:
    """Parse arguments and write the export."""
    parser = make_parser(__doc__)
    parser.add_argument("group", help="Group display name or object ID")
    parser.add_argument(
        "--transitive",
        action="store_true",
        help="Include members of nested groups",
    )
    add_output_args(parser)
    args = parser.parse_args()

    started = time.monotonic()
    log(f"Started: exporting members of {args.group!r}")
    with running():
        session = graph_session()
        group = resolve_group(session, args.group)
        log(f"Found group {group['displayName']!r} ({group['id']})")
        rows, sign_ins = build_rows(session, group["id"], transitive=args.transitive)

    stem = re.sub(r"[^\w-]+", "_", args.group)
    output: Path = args.output or Path(f"{stem}_members.{args.format}")
    write_output(rows, output, args.format)
    log(f"Wrote {len(rows)} members to {output}", style="green")
    summarise(rows, noun="members", sign_ins=sign_ins, started=started)


if __name__ == "__main__":
    main()
