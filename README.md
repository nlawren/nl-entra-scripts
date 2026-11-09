# Entra Scripts

A few maybe useful scripts for querying information from Entra ID (via Microsoft
Graph) and exporting it to Excel or CSV. They're read-only and meant for one-off
requests run from the command line.

| Script | What it exports |
| --- | --- |
| [`group_members.py`](scripts/group_members.py) | Every member of a group, optionally including nested groups |
| [`guest_users.py`](scripts/guest_users.py) | Every guest user from one email domain |
| [`check_users.py`](scripts/check_users.py) | Whether each user listed in a CSV still exists and is enabled |

## Requirements

- [uv](https://docs.astral.sh/uv/). Each script declares its own dependencies
  and Python version (3.14), so `uv run` sets everything up on first use.
- [Azure CLI](https://learn.microsoft.com/cli/azure/), signed in with `az login`.
  The scripts use that account's token; nothing else is stored or configured.
- Graph permissions for that account:

  | Permission | Needed by |
  | --- | --- |
  | `GroupMember.Read.All` | `group_members.py` |
  | `User.Read.All` | all |
  | `AuditLog.Read.All` | all, for last sign-in dates |

Last sign-in dates also need an Entra ID P1/P2 licence in the tenant and a
directory role such as Reports Reader, Security Reader or Global Reader. Without
them the scripts still run, show a warning, and leave `lastSignInDate` blank.

## Usage

Run the scripts from the `scripts` folder. They all have `-h`/`--help`, which
lists every option with examples.

### group_members.py

```sh
uv run group_members.py "My Group"
uv run group_members.py "My Group" --transitive
uv run group_members.py "My Group" --format csv
uv run group_members.py 0638c5e9-1424-4da6-9538-a4f8ea8799ac -o members.xlsx
```

| Argument | Description |
| --- | --- |
| `group` | Group display name (exact match) or object ID |
| `--transitive` | Include members of nested groups |
| `--format {xlsx,csv}` | Output format (default: `xlsx`) |
| `-o`, `--output` | Output file path (default: `My_Group_members.xlsx`, from the group name) |

Members that aren't users, such as nested groups or devices, are listed with
their type and name, and the user-only columns are left blank.

### guest_users.py

```sh
uv run guest_users.py removed.example.com
uv run guest_users.py removed.example.com --format csv
uv run guest_users.py removed.example.com -o guests.xlsx
```

| Argument | Description |
| --- | --- |
| `domain` | Email domain, e.g. `removed.example.com` |
| `--format {xlsx,csv}` | Output format (default: `xlsx`) |
| `-o`, `--output` | Output file path (default: `removed_example_com_guests.xlsx`, from the domain) |

A guest matches if either its mail or its user principal name is in the domain.
Guest user principal names look like
`first.last_removed.example.com#EXT#@tenant.onmicrosoft.com`, so guests with no
mail address are still found. Only the exact domain matches: `example.com` won't
match `a.example.com`.

### check_users.py

```sh
uv run check_users.py powerapps_users.csv
uv run check_users.py powerapps_users.csv --format csv
uv run check_users.py powerapps_users.csv -o checked.xlsx
```

| Argument | Description |
| --- | --- |
| `csv` | CSV file listing the users to check |
| `--format {xlsx,csv}` | Output format (default: `xlsx`) |
| `-o`, `--output` | Output file path (default: `powerapps_users_checked.xlsx`, from the input file) |

The CSV needs an `Object id` column, a `User principle name` column, or both. Any
case works, and so does `User principal name`. Other columns are ignored. Save it
as UTF-8, which Excel calls "CSV UTF-8".

Each user is looked up by object ID, because that never changes, and then by user
principal name if there's no usable ID or the ID isn't found. You get one output
row per CSV row, with the accounts needing attention first and `accountStatus` set
to one of:

| `accountStatus` | Meaning |
| --- | --- |
| `Not found` | Graph confirmed there's no such account |
| `Lookup failed` | Graph gave no answer, so the account may still exist; `notes` says why |
| `Disabled` | The account exists and is disabled |
| `Enabled` | The account exists and is enabled |

An extra `notes` column flags a user principal name that has changed since the CSV
was exported, or one that now belongs to a different account from the CSV's object
ID. In that case the row describes the current account, not the one in the CSV.

## Output

All three scripts write the same columns, and `check_users.py` adds `notes`:

| Column | Description |
| --- | --- |
| `type` | `user`, `group`, `device` and so on |
| `displayName` | Display name |
| `userPrincipalName` | Sign-in name |
| `mail` | Primary email address |
| `accountType` | `Standard` or `Guest` |
| `accountStatus` | `Enabled` or `Disabled`, plus `Not found` and `Lookup failed` from `check_users.py` |
| `lastSignInDate` | Date of the last *successful* sign-in, in Sydney time |
| `id` | Entra object ID |

The Excel output is a filterable table with a frozen header, and dates are real
Excel dates, so they sort and filter properly. The CSV output includes a BOM so
Excel opens it as UTF-8.

Entra has only recorded successful sign-ins since December 2023, so a blank
`lastSignInDate` means the account hasn't signed in since then, or at all.

While running, the scripts show progress bars, coloured warnings and a summary
table at the end. All of this goes to stderr. When the output is piped or
redirected, it falls back to plain timestamped lines.

## Development

The `scripts` folder is also a uv project with ruff and ty as dev dependencies,
and CI runs the same checks on every push and pull request to `main`:

```sh
cd scripts
uv sync
uv run ruff check .
uv run ruff format --diff .
uv run ty check
```

`guest_users.py` and `check_users.py` import their Graph helpers from
`group_members.py`, so the scripts need to stay in the same folder.
