# /// script
# requires-python = ">=3.14"
# dependencies = [
#     "azure-identity",
#     "requests",
#     "xlsxwriter",
# ]
#
# [tool.ty.environment]
# extra-paths = ["."]  # resolve group_members, as `uv run` does
# ///
"""Export Entra ID guest users from one email domain to Excel (default) or CSV via Microsoft Graph.

Usage:
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

import argparse
import re
import time
from pathlib import Path
from typing import Any

import requests
from group_members import (
    COLUMNS,
    GRAPH,
    HTTP_FORBIDDEN,
    USER_SELECT,
    USER_SELECT_BASIC,
    graph_session,
    log,
    paged,
    sign_in_date,
    write_csv,
    write_xlsx,
)

PROGRESS_EVERY = 5000  # guests scanned between progress lines


def in_domain(user: dict[str, Any], domain: str) -> bool:
    """Return True if the guest's mail or external UPN is in `domain`."""
    mail = (user.get("mail") or "").casefold()
    upn = (user.get("userPrincipalName") or "").casefold()
    # Guest UPNs encode the home address as local_domain#EXT#@tenant.onmicrosoft.com.
    external = upn.partition("#ext#")[0] if "#ext#" in upn else ""
    return mail.endswith(f"@{domain}") or external.endswith(f"_{domain}")


def list_guests(session: requests.Session, select: str) -> list[dict[str, Any]]:
    """Return every guest user in the tenant with the given $select."""
    url = f"{GRAPH}/users?$filter=userType eq 'Guest'&$select={select}&$top=999"
    guests = []
    for n, user in enumerate(paged(session, url), start=1):
        guests.append(user)
        if n % PROGRESS_EVERY == 0:
            log(f"  {n} guests scanned")
    return guests


def fetch_guests(session: requests.Session) -> list[dict[str, Any]]:
    """List all guests, dropping signInActivity if Graph won't return it.

    Graph rejects the whole listing with 403 if signInActivity isn't allowed
    (no P1/P2 licence, permission or role), so retry once without it.
    """
    try:
        return list_guests(session, USER_SELECT)
    except requests.HTTPError as e:
        if e.response is None or e.response.status_code != HTTP_FORBIDDEN:
            raise
        log(f"Warning: no sign-in activity, lastSignInDate left blank: {e.response.text}")
        return list_guests(session, USER_SELECT_BASIC)


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
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("domain", help="Email domain, e.g. removed.example.com")
    parser.add_argument(
        "--format",
        choices=["xlsx", "csv"],
        default="xlsx",
        help="Output format (default: xlsx)",
    )
    parser.add_argument("-o", "--output", type=Path, help="Output file path")
    args = parser.parse_args()
    domain = args.domain.removeprefix("@").casefold()

    started = time.monotonic()
    log(f"Started: exporting guests from {domain!r}")
    session = graph_session()
    guests = fetch_guests(session)
    matches = [g for g in guests if in_domain(g, domain)]
    log(f"Scanned {len(guests)} guests, {len(matches)} from {domain!r}")
    rows = build_rows(matches)

    stem = re.sub(r"[^\w-]+", "_", domain)
    output: Path = args.output or Path(f"{stem}_guests.{args.format}")
    if args.format == "csv":
        write_csv(rows, output)
    else:
        write_xlsx(rows, output)

    log(f"Wrote {len(rows)} guests to {output} in {time.monotonic() - started:.1f}s")


if __name__ == "__main__":
    main()
