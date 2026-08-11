# VeeamONE Restoration Status Percentage Report

Generate a **restore status percentage** report (success / warning / failed
rate) from the **Veeam ONE** monitoring database. Produces a CSV plus a
self-contained HTML report (donut chart, daily trend, and a filterable table of
every restore) — no external libraries, no internet access, no changes to Veeam.

The report reads restore activity directly from `[monitor].[BpRestoreItem]` in
the Veeam ONE database (read-only).

> **Not affiliated with Veeam Software.** Community tooling that queries the
> Veeam ONE database read-only. Test before using in production.

## What you get

- **Success rate** across the selected period, with Success / Failed / Skipped counts
- **Donut chart** of the status distribution and a **daily trend** of restores
- **Table** of every restore: type, restored object, destination host, status,
  start/end, elapsed time, size, and status message
- **CSV**, **HTML**, and optional **PDF** output
- **E-mail delivery** — send the report (with attachments) to a distribution list
- **Scheduling** — run it automatically **Daily / Weekly / Monthly / Yearly** and e-mail it

## Requirements

- Windows PowerShell **5.1** (ships with Windows)
- Network access to the SQL Server hosting the Veeam ONE database
- A **read-only** login to the Veeam ONE database. One of:
  - the current Windows account (Integrated auth),
  - a domain account (username + password, net-only impersonation), or
  - a SQL Server login.

  > **Use a read-only account.** The tool only ever runs `SELECT` queries. Grant
  > the login **`db_datareader`** on the Veeam ONE database (no write/DDL rights)
  > so it cannot change anything. Avoid `sa` / admin logins in production.
- For **PDF** output: Microsoft Edge or Google Chrome installed (headless render)
- For **e-mail**: an SMTP relay the host can reach

## Files

| File | Purpose |
|------|---------|
| `Get-VeeamOneRestoreReport.ps1` | The engine — queries the DB, builds CSV/HTML/PDF, and e-mails it |
| `Show-RestoreReportForm.ps1` | Windows GUI front-end (Report / Email / Schedule tabs) |
| `Register-RestoreReportSchedule.ps1` | Registers a Windows Scheduled Task for recurring runs |

## Quick start

### GUI

```powershell
powershell -ExecutionPolicy Bypass -STA -File .\Show-RestoreReportForm.ps1
```

Enter the SQL Server, database (`VeeamONE`), pick an authentication mode and a
date range, then **Run report**.

### Command line

```powershell
# 1. (Recommended once) confirm the state / restore-type codes on your build
.\Get-VeeamOneRestoreReport.ps1 -SqlServer YOURSQL -Database VeeamONE `
    -SqlCredential (Get-Credential sa) -Discover

# 2. Generate the report (last 30 days by default)
.\Get-VeeamOneRestoreReport.ps1 -SqlServer YOURSQL -Database VeeamONE `
    -SqlCredential (Get-Credential sa) -StartDate 2026-07-01

# 3. Add a PDF and e-mail it to a distribution list
.\Get-VeeamOneRestoreReport.ps1 -SqlServer YOURSQL -Database VeeamONE `
    -Period Weekly -Pdf `
    -EmailTo ops@contoso.com,team@contoso.com -EmailFrom veeam-reports@contoso.com `
    -SmtpServer smtp.contoso.com

# Preview the report layout with no database
.\Get-VeeamOneRestoreReport.ps1 -DemoData -Pdf
```

The CSV and HTML are written to `.\output` by default (override with
`-OutputFolder`).

## Key parameters

| Parameter | Description |
|-----------|-------------|
| `-SqlServer` | SQL Server instance (IP or name). Auto-detected from the registry if omitted. |
| `-Database` | Veeam ONE database name (default `VeeamONE`). |
| `-SqlCredential` | `PSCredential` for a SQL login (e.g. `sa`). |
| `-WindowsCredential` | Domain account for Windows auth (net-only impersonation). |
| `-StartDate` / `-EndDate` | Reporting window. Default: last 30 days. |
| `-StartTime` / `-EndTime` | Optional time-of-day window (e.g. nightly `22:00`–`06:00`). |
| `-RestoreType` | Optional `LIKE` filter on the restore type. |
| `-NameLike` | Optional `LIKE` filter on the restored object name. |
| `-Period` | Relative window for scheduled runs: `Daily` / `Weekly` / `Monthly` / `Yearly`. |
| `-Pdf` | Also render a PDF (via headless Edge/Chrome). |
| `-EmailTo` | One or more recipients (enables e-mailing). Comma-separated is fine. |
| `-EmailFrom` | Sender address (defaults to a no-reply on this host). |
| `-SmtpServer` / `-SmtpPort` | SMTP relay host and port (default `25`). |
| `-SmtpUseSsl` | Use SSL/TLS for SMTP. |
| `-SmtpCredential` | `PSCredential` for SMTP auth (omit for anonymous relay). |
| `-Discover` | Inspect the DB and print the real `state` / `item_type` values. |
| `-DemoData` | Build the report from synthetic data (no DB). |

## PDF, e-mail & scheduling

**PDF** — add `-Pdf` (or tick the box on the Report tab). Rendering uses the
headless mode of **Microsoft Edge** or **Google Chrome** already on the machine —
no extra modules. If neither browser is present, the CSV + HTML are still produced.

**E-mail** — supply `-EmailTo` (and `-SmtpServer`). The CSV, HTML and PDF are
attached and a short summary is shown in the message body. In the GUI, use the
**Email** tab.

**Scheduling** — `Register-RestoreReportSchedule.ps1` creates a Windows Scheduled
Task that runs the report on a recurring basis and e-mails it. In the GUI, use the
**Schedule** tab.

```powershell
# Every day at 07:00, PDF, e-mailed to the ops team
.\Register-RestoreReportSchedule.ps1 -Cadence Daily -Time 07:00 `
    -SqlServer YOURSQL -Database VeeamONE -Pdf `
    -EmailTo ops@contoso.com -EmailFrom veeam-reports@contoso.com `
    -SmtpServer smtp.contoso.com `
    -RunAsCredential (Get-Credential CONTOSO\svc_veeam)

# Other cadences
.\Register-RestoreReportSchedule.ps1 -Cadence Weekly  -DayOfWeek Monday -Time 06:30 ...
.\Register-RestoreReportSchedule.ps1 -Cadence Monthly -DayOfMonth 1 ...
.\Register-RestoreReportSchedule.ps1 -Cadence Yearly  -Month January -DayOfMonth 1 ...

# Remove a schedule
.\Register-RestoreReportSchedule.ps1 -Unregister -TaskName "VeeamONE Restore Report"
```

> **Unattended authentication.** A scheduled task runs as the Windows account you
> pass with `-RunAsCredential`, and the report connects to the Veeam ONE database
> as that account (Windows/Integrated auth). Give it **read** access to the Veeam
> ONE database and SMTP relay permission. No SQL or SMTP password is stored in the
> task — only the Windows run-as password, which Windows keeps in its LSA secret
> store. Without `-RunAsCredential`, the task runs only while that user is logged on.

## The data source

Restore activity is read per restored item from:

```
[VeeamONE].[monitor].[BpRestoreItem]
```

| Report column | Table column |
|---------------|--------------|
| Restore type | `item_type` |
| Object | `item_name` |
| Destination host | `destination_host` |
| Status | `state` (2 → Success, 3 → Failed, 6 → Skipped) |
| Start / End | `start_time` / `finish_time` |
| Size | `item_size` |
| Reason | `message` |

## Tuning (only if `-Discover` shows different codes)

The status mapping is the one version-sensitive bit, isolated at the top of
`Get-VeeamOneRestoreReport.ps1`:

```powershell
$script:StateMap = @{ 2 = 'Success'; 3 = 'Failed'; 6 = 'Skipped' }
```

If your build encodes `state` differently, adjust these three lines. When
`state` is unknown, the report falls back to classifying the `message` text, so
restores still get a sensible status either way. Use `$script:ItemTypeMap` to
give the numeric `item_type` values friendly labels.

## Troubleshooting

- **"No restore sessions matched"** — widen the date range, or run `-Discover`
  to confirm the table has rows and the `state` / `item_type` codes match.
- **Login / connection errors** — verify the SQL Server name, that the account
  has read access to the Veeam ONE database, and that SQL/Windows auth matches
  what the server accepts.
- **All restores show one status** — check the `-Discover` output and adjust
  `$script:StateMap`.
