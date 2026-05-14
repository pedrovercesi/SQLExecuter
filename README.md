# SQLExecuter

Interactive runner that applies SQL scripts organized in numbered subfolders against a SQL Server database, via `sqlcmd`. It iterates folders and files in numeric order, supports resume, and writes logs.

## Expected layout

```
SQLExecuter/
├── Apply-Reports.cmd        # double-click wrapper
├── Apply-Reports.ps1        # main runner (PowerShell)
├── deploy-config.json       # connection settings
├── 0.First/                 # numeric-prefixed subfolders
│   └── 01.script.sql
├── 1.Skip/
│   └── ...
├── 2.Reports/
│   └── ...
├── logs/                    # auto-created
└── .deploy-state.json       # auto-created (state / resume)
```

Ordering rules:
- Only subfolders starting with `<number>.` are executed (e.g. `0.First`, `2.Reports`).
- Inside each folder, `.sql` files are ordered by the numeric prefix in their name (e.g. `01.foo.sql`, `02.bar.sql`).

## Configuration — `deploy-config.json`

```json
{
  "server": "INSTANCE_SQL",
  "database": "DatabaseName",
  "authentication": "SQL",
  "username": "userSQL",
  "sqlcmdPath": "sqlcmd",
  "skipFolders": ["1.Skip"]
}
```

Fields:
- `server` *(required)* — SQL Server instance (e.g. `SERVER\INSTANCE`). The literal placeholder `SERVER\INSTANCE` is rejected.
- `database` *(required)* — target database name.
- `authentication` *(required)* — `Windows` or `SQL`.
- `username` — only used when `authentication: "SQL"`. The password is prompted at startup (never stored in a file).
- `sqlcmdPath` *(optional)* — explicit path to `sqlcmd.exe`. If omitted, the script searches `PATH` and the usual SQL Server Client Tools locations.
- `skipFolders` *(optional)* — list of folder names to ignore (case-insensitive).

## Prerequisites

- Windows with PowerShell 5.1+
- `sqlcmd` installed (SQL Server Command Line Utilities / `mssql-tools`)
- Network access and permissions on the target database

## How to run

### Option 1 — double-click (recommended)

Double-click `Apply-Reports.cmd`. The wrapper invokes PowerShell with `-ExecutionPolicy Bypass` and keeps the window open at the end.

### Option 2 — from PowerShell

```powershell
cd C:\Claude\SQLExecuter
.\Apply-Reports.ps1
```

To point at a different config file:

```powershell
.\Apply-Reports.ps1 -ConfigPath .\deploy-config.prod.json
```

## Interactive flow

1. Reads `deploy-config.json` and resolves `sqlcmd`.
2. If `authentication=SQL`, prompts for the password (read as a `SecureString`).
3. Tests the connection with `SELECT @@VERSION`.
4. For each numbered subfolder (excluding `skipFolders`):
   - Lists the scripts and prompts `Y/N/Q` (apply / skip / quit).
   - Runs each script via `sqlcmd -b -V 16` (any error of severity >= 16 fails the script).
   - On failure, offers `R/S/A` — **R**etry, **S**kip script, **A**bort folder.
5. Updates `.deploy-state.json` after each folder (so a later run knows what already ran as OK / PARTIAL / FAILED).
6. Writes a timestamped log to `logs\deploy-yyyyMMdd-HHmmss.log`.
7. Prints a summary at the end (OK / PARTIAL / FAILED / SKIPPED).

## Generated files

- `logs\deploy-*.log` — output from each `sqlcmd` call, with timestamps. Safe to delete.
- `.deploy-state.json` — per-folder state (`OK`, `PARTIAL`, `FAILED`) and which scripts ran. Deleting this file only resets progress tracking; it does not undo database changes.

## Exit codes

- `0` — run completed (even if some folders were skipped by the user).
- `1` — config could not be read, connection failed, or no numbered subfolder was found.
