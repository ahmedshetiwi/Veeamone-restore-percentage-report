# VeeamONE Custom Report (Restore Status Percentage Report - Veeam Standalone/Managed RMAN Plugin- SAP HANA Standalone/Managed Plugin - All Backup Jobs Calender - All Backup Job Report - Email&Scheduling)

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
| `Get-VeeamOneRestoreReport.ps1` | Restore report engine — queries the DB, builds CSV/HTML/PDF, e-mails it |
| `Get-VeeamOneBackupJobReport.ps1` | **Backup Jobs** report engine — job inventory (schedule, type, last run) by policy |
| `Get-VeeamOneRmanBackupReport.ps1` | **RMAN Backups** report engine — RMAN plugin rescan success/failure rate |
| `Get-VeeamOneAllBackupsReport.ps1` | **All Backups** report engine — every backup run, rich detail + failure reasons |
| `Get-VeeamOneSapHanaReport.ps1` | **SAP HANA** report engine — backint plugin, managed + standalone/unmanaged |
| `Show-RestoreReportForm.ps1` | Windows GUI front-end (**VeeamOne Custom Reports**) — 8 tabs: **Connection** (shared DB/auth/output for every report), Restoration report, Backup Jobs, All Backups, RMAN Backups, SAP HANA, Email, Schedule. Each report tab has its own Start/End date range. |
| `Register-RestoreReportSchedule.ps1` | Registers a Windows Scheduled Task for recurring runs (any report via `-ReportType`) |

> **Backup reports** are documented in their own section below
> ([Backup reports](#backup-reports-jobs--rman--all-backups)). They share the
> restore report's connection, authentication, PDF, e-mail and scheduling model.

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

### Retry-aware success
If a restore of the **same object** (item name + destination host + type) **failed**
but a later attempt **succeeded within a few hours**, the failed attempt is counted
as **success** — it was resolved on retry, so it should not drag down the success
rate.

- Controlled by **`-RetryWindowHours`** (default **6**; set **0** to disable).
- On the **Restoration report** tab there's a checkbox *"Count a failed restore as
  success if it succeeded on retry within N hours"* with an adjustable hours box.
- Reclassified rows show **Success** with a Reason note *"Auto-resolved: initially
  failed, succeeded on retry at …"*, a **Resolved on retry** tile appears, and a
  line under the tiles states how many were reclassified.

```powershell
.\Get-VeeamOneRestoreReport.ps1 -RetryWindowHours 6    # default (enabled)
.\Get-VeeamOneRestoreReport.ps1 -RetryWindowHours 0    # disable, count every failure
```

### Count by restore job vs restored object
By default every **restored object** is a row (`monitor.BpRestoreItem`). Add
**`-ByRestoreJob`** (or tick *"Count by restore job (session), not per restored
object"* on the Restoration tab) to collapse all objects of one restore session
(`restore_session_uid`) into **one row per restore job**:

- Job **status** = worst object outcome (Failed if any object failed, then
  Running / Warning / Success / Skipped).
- **Object** shows the single object name, or *"N objects"*; **Size** is the sum;
  **Start/End** span the session.
- Tiles and success rate are then per **restore job** instead of per object.

```powershell
.\Get-VeeamOneRestoreReport.ps1 -ByRestoreJob
```

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

---

# Backup reports (Jobs / RMAN / All Backups)

Three backup reports were added alongside the restore report. Each is a
self-contained engine with the **same** connection, authentication (Windows
current-user / domain net-only / SQL login), `-Period`, `-Pdf`, e-mail
(`-EmailTo …`) and scheduling surface as the restore report. In the GUI they are
the **Backup Jobs**, **RMAN Backups** and **All Backups** tabs — they reuse the
Report tab's connection + date range and the Email tab's SMTP settings.

| Report | Engine | Source table | Answers |
|--------|--------|--------------|---------|
| **Backup Jobs** | `Get-VeeamOneBackupJobReport.ps1` | `monitor.BpJob` (+ `BpJobLastFinishedResult`, `BpBackup`) | Every job's configured schedule, type, retention, last run and status, grouped by backup policy |
| **RMAN Backups** | `Get-VeeamOneRmanBackupReport.ps1` | `monitor.BpJobSession` | RMAN plugin (standalone/unmanaged) rescan success / warning / failure rate |
| **All Backups** | `Get-VeeamOneAllBackupsReport.ps1` | `monitor.BpJobSession` | Every backup run — status, type, mode, timing, sizes, processing rate, failure reasons |
| **SAP HANA** | `Get-VeeamOneSapHanaReport.ps1` | `monitor.BpJobSession` | SAP HANA (backint) plugin — success/failure rate across **managed + standalone/unmanaged** |

Each produces the familiar donut + daily-trend + summary tiles + filterable
table, plus CSV / HTML / optional PDF, and optional e-mail.

## Backup Jobs report

```powershell
# All jobs, grouped by policy (RMAN example: filter by name)
.\Get-VeeamOneBackupJobReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)
.\Get-VeeamOneBackupJobReport.ps1 -NameLike '%RMAN%'          # focus RMAN
.\Get-VeeamOneBackupJobReport.ps1 -PolicyLike '%Linux%' -EnabledOnly
.\Get-VeeamOneBackupJobReport.ps1 -DemoData                    # preview, no DB
```

Columns: **Backup policy · Job · Type · Schedule · Retention · Enabled · Last
start · Last end · Next run · Status · Message**. It is an inventory of current
job state, so the date range is not applied to row selection.

**Backup policy column** — `backup_policy_name` from `monitor.BpBackup` is the VBR
*policy* that groups jobs. It is populated for **policy-managed** jobs and **empty
for standalone / unmanaged** jobs (RMAN, plugins). If **no** job in the report has a
policy, the **Backup policy** column, the *Backup policies* tile and the per-policy
chart are **automatically hidden** (the chart switches to *Jobs by type*), so
standalone/unmanaged environments don't get an empty column.

**A5 backup calendar** — add `-Calendar` (or tick *Also generate A5 backup
calendar* on the Backup Jobs tab) to also write a print-friendly **A5** HTML
(and PDF, with `-Pdf`) timetable of the configured schedules, grouped by
schedule type (Daily / Weekly / Monthly / Triggered-rescan) and sorted by run
time, showing each job's time, next run and last start/end.

```powershell
.\Get-VeeamOneBackupJobReport.ps1 -Calendar -Pdf
```

## RMAN Backups report

```powershell
# RMAN plugin rescans, last 30 days (default filter: job name LIKE %RMAN%)
.\Get-VeeamOneRmanBackupReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)
.\Get-VeeamOneRmanBackupReport.ps1 -IncludeAllBackups          # every backup, not just RMAN
.\Get-VeeamOneRmanBackupReport.ps1 -JobNameLike '%ORCL%' -Period Weekly -Pdf
```

`result` codes on the confirmed build: **2 = Success, 3 = Warning, 4 = Failed**
(isolated in `$script:ResultMap`).

The report also includes a **Databases** section: tiles for *Databases backed up*,
*Running normally* (latest run = Success), *Databases w/ issues*, *Database
servers*, and **Total DB data read**, plus a per-database table with columns
**Job name · Object (database) · RMAN scan · Runs · Last status · Last run ·
Total read · Success rate**. Object (database), host and the RMAN scan name
(`<host>-scan`, matching `monitor.BpEpPluginCluster.access_name`) are derived
from the RMAN job name (`BKP-<host>-<db>…RMAN`) — for RMAN standalone/unmanaged,
one job = one database. **Total read** is the sum of data read across that
database's rescans (the readable DB size).

**Object drill-down** — the **Object (database)** cell lists the real Oracle
databases (and PDBs, shown as `DB > PDB`) that each RMAN job protects, read from
the Veeam ONE plugin inventory:

```
monitor.BpEpPluginCluster (scan, access_name = "<host>-scan")
  --BpEpPluginClusterToDbEntityLink-->  BpEpPluginDbEntity (type 0 = Oracle home)
  --parent_id-->  BpEpPluginDbEntity (type 1 = database)
  --parent_id-->  BpEpPluginDbEntity (type 2 = PDB)
```

Each RMAN job is matched to its scan by the `<host>-scan` naming; if the plugin
tables are unavailable the cell falls back to the database parsed from the job name.

> **Performance / hangs.** All report queries run under `READ UNCOMMITTED`, so a
> `SELECT` never waits on the Veeam ONE collector's writer locks (this is what
> could make a report appear to "keep loading"). The object drill-down query also
> has a bounded 90-second timeout. If the plugin tables are very large or you want
> to skip the drill-down entirely, add **`-SkipObjectDrilldown`** — the report
> still lists every job, scan and size, with the object column falling back to the
> job-name database.

### Sizes: DB read vs stored
| Column | Table column | Meaning |
|--------|--------------|---------|
| **DB read** / Total DB read | `processed_used_size` | Source data read from the database (the readable DB size — where large 100s-of-GB/TB scans show). |
| **Stored (repo)** / Total stored | `transferred_size` | Data written to the repository after compression/dedup. |

`backedup_size` is **not** used — it is 0 on many RMAN archive scans and hid the
large scans.

> **Per-database names inside an instance** (e.g. SQL `msdb` / `master` / user DBs,
> Oracle PDBs) are **not** stored in the job/session tables. To locate the VeeamONE
> table that does hold per-object/database rows on your build, run:
> ```powershell
> .\Get-VeeamOneRmanBackupReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa) -DiscoverObjects
> ```
> It lists candidate tables + columns; share the result (or export that table) and a
> per-database drill-down can be wired into the report.

## All Backups report

```powershell
.\Get-VeeamOneAllBackupsReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)
.\Get-VeeamOneAllBackupsReport.ps1 -StatusFilter Failed         # only failures
.\Get-VeeamOneAllBackupsReport.ps1 -NameLike '%SQL%' -Period Monthly -Pdf `
    -EmailTo ops@contoso.com -SmtpServer smtp.contoso.com
```

Columns: **Job · Type · Mode · Status · Start · End · Elapsed · Data read ·
Transferred · Rate · Retried · Reason**, plus a status-by-job-type breakdown and
a top-failing-jobs panel.

## SAP HANA report

```powershell
# SAP HANA (backint) plugin - both managed and unmanaged, last 30 days
.\Get-VeeamOneSapHanaReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)
.\Get-VeeamOneSapHanaReport.ps1 -Deployment Unmanaged          # standalone/unmanaged only
.\Get-VeeamOneSapHanaReport.ps1 -Deployment Managed -Period Weekly -Pdf
```

Columns: **Job · Deployment · Mode · Status · Start · End · Elapsed · Data read ·
Transferred · Reason**, with Managed / Unmanaged tiles. **Deployment** is derived
from the job name — VBR-managed application policies carry a `DBBKP-<host>-SAP\…`
prefix; bare `<host> SAP backint backup (<repo>)` jobs are standalone/unmanaged.
The default filter is `-JobNameLike '%backint%'`.

## Job-type names

The numeric Veeam `job_type` is overloaded, so all three reports classify each
job by its **name first** (most reliable), then fall back to a numeric map
(`$script:JobTypeNumMap`). Recognised categories include:

`Oracle RMAN Plugin (standalone/unmanaged)`, `Oracle Plugin Backup`,
`Oracle Plugin - redo/archived logs`, `SAP HANA Plugin (backint)`,
`Microsoft SQL Server (plugin/agent)`, `Microsoft SQL Log Backup (plugin)`,
`Veeam Agent Backup`, `VM Backup (VMware/virtual)`, `Replication`,
`Backup Copy`, `Backup to Tape`, `NAS / File Share Backup`, `Nutanix AHV Backup`.

To adjust a label, edit the name-pattern rules in `Get-JobTypeName` or the
`$script:JobTypeNumMap` table near the top of each engine.

### Agent backups: SQL vs RMAN-trigger

In the **Backup Jobs** and **All Backups** reports, plain `Veeam Agent Backup`
rows are further split into:

- **`Veeam Agent - RMAN host (script)`** — the agent's host also has an RMAN
  plugin job, i.e. an Oracle DB server where RMAN runs (typically kicked off by a
  script on the host).
- **`Veeam Agent - SQL server`** — the host name indicates SQL.
- **`Veeam Agent Backup`** — everything else.

> This is an **inference**. VeeamONE does not store a job's pre/post script or
> (for these agent jobs) an application-aware flag — in the sample data every
> agent job had `app_aware_processing_enabled = 0`. The split is derived from the
> job/host name and from whether the same host also has an RMAN plugin job. SQL
> databases in this environment are protected by the **SQL plugin** (they appear
> as *Microsoft SQL …*), not by application-aware agent jobs.

## Scheduling any report (same scheduler + SMTP)

`Register-RestoreReportSchedule.ps1` schedules **any** report via `-ReportType`
(`Restore` | `BackupJobs` | `RmanBackups` | `AllBackups` | `SapHana`, default `Restore`),
using one generic `-Filter` that maps to the right filter parameter per report.
The GUI **Schedule** tab has a *Report to schedule* picker that drives this.

```powershell
# Daily RMAN backup success report at 07:00, PDF, e-mailed
.\Register-RestoreReportSchedule.ps1 -ReportType RmanBackups -Cadence Daily -Time 07:00 `
    -SqlServer YOURSQL -Database VeeamONE -Pdf `
    -EmailTo ops@contoso.com -SmtpServer smtp.contoso.com `
    -RunAsCredential (Get-Credential CONTOSO\svc_veeam)

# Weekly backup-jobs inventory, filtered to RMAN
.\Register-RestoreReportSchedule.ps1 -ReportType BackupJobs -Cadence Weekly -DayOfWeek Monday `
    -Filter '%RMAN%' -SqlServer YOURSQL -EmailTo ops@contoso.com -SmtpServer smtp.contoso.com

# Remove a report's schedule
.\Register-RestoreReportSchedule.ps1 -Unregister -ReportType AllBackups
```

Each report type registers under its own task name (e.g. *VeeamONE RMAN Backup
Report*), so all four can be scheduled side by side.

