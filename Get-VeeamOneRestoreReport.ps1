<#
.SYNOPSIS
    Collects restore-session details from the Veeam ONE monitoring database
    (Microsoft SQL Server) and exports a CSV plus a self-contained HTML report
    with a success / failure rate chart.

.DESCRIPTION
    Runs on (or next to) the Veeam ONE server. Reads the Veeam ONE database
    connection from the registry automatically (override with -SqlServer /
    -Database), connects with Windows Integrated authentication by default (or a
    domain account via net-only impersonation, or a SQL login via -SqlCredential),
    and queries the restore sessions Veeam ONE has collected from the monitored
    Veeam Backup & Replication servers.

    Two modes:
      1. -Discover : inspect the DB - list restore/session-looking tables with
                     their columns, and the distinct session types with counts and
                     sample names, so you can confirm which table/type values map
                     to restore operations on your Veeam ONE build.
      2. Report    : produce the CSV + HTML report (success / failure / warning
                     rate, daily trend, and a filterable table of restore
                     sessions).

    Output per restore session: Restore type, Object (restored item), Initiated
    by, Status, Start date/time, End date/time, Elapsed time, Reason - plus a
    success/failure donut and a daily trend in the HTML report.

    NOTE: The Veeam ONE database schema differs between versions. The connection,
    filtering, CSV and HTML generation are version-independent. The only
    version-sensitive part is the table/column mapping, isolated in the
    $script:Schema hashtable below - the ONE thing you may need to adjust after
    running -Discover on your build. Override it at run time without editing the
    file using -SessionTable / -TypeColumn if your table/column names differ.

.PARAMETER SqlServer
    SQL Server instance hosting the Veeam ONE database, e.g.
    "VEEAMONE\VEEAMSQL2016". Auto-detected from the registry if omitted.

.PARAMETER Database
    Veeam ONE database name (default auto-detected, usually "VeeamOne").

.PARAMETER SqlCredential
    Optional PSCredential for SQL authentication. Omit to use the current
    Windows account (Integrated Security).

.PARAMETER WindowsCredential
    Optional domain account for Windows auth via net-only impersonation
    (same mechanism as "runas /netonly"). Use to authenticate to a remote SQL
    server as a domain account without running the whole session as that user.

.PARAMETER StartDate
    Only include restore sessions that STARTED on/after this date. Default: 30 days ago.

.PARAMETER EndDate
    Only include restore sessions that STARTED before this date (exclusive end of
    day if no time-of-day supplied). Default: now.

.PARAMETER StartTime
    Optional time-of-day lower bound (e.g. "22:00"). Combine with -EndTime for a
    window. If -EndTime < -StartTime the window is treated as crossing midnight.

.PARAMETER EndTime
    Optional time-of-day upper bound (e.g. "06:00").

.PARAMETER RestoreType
    Optional filter on the session/restore type text (SQL LIKE, one or more
    patterns, e.g. "%File%","%Instant%"). Omit to include every restore type.

.PARAMETER NameLike
    Optional extra SQL LIKE filter on the session/object name (e.g. "%SQLPROD%").

.PARAMETER SessionTable
    Override the restore-session table (from -Discover), e.g. "[dbo].[Sessions]".
    Overrides $script:Schema.Table without editing the file.

.PARAMETER TypeColumn
    Override the column that holds the restore/session type (from -Discover).
    Overrides $script:Schema.TypeCol without editing the file.

.PARAMETER OutputFolder
    Where to write the CSV + HTML. Default: .\output next to this script.

.PARAMETER Discover
    Run schema/type discovery and exit (no report produced).

.PARAMETER DemoData
    Generate the CSV + HTML from synthetic data (no DB) to preview the report format.

.EXAMPLE
    # Step 1 - inspect the Veeam ONE DB: which table + type values = restores
    .\Get-VeeamOneRestoreReport.ps1 -Discover

.EXAMPLE
    # Step 2 - last 30 days restore report (all restore types)
    .\Get-VeeamOneRestoreReport.ps1

.EXAMPLE
    # Only file-level & instant recoveries, custom range
    .\Get-VeeamOneRestoreReport.ps1 -RestoreType '%File%','%Instant%' `
        -StartDate 2026-07-01 -EndDate 2026-08-01

.EXAMPLE
    # Preview the report layout with no database
    .\Get-VeeamOneRestoreReport.ps1 -DemoData
#>
[CmdletBinding(DefaultParameterSetName = 'Report')]
param(
    [string]   $SqlServer,
    [string]   $Database,
    [pscredential] $SqlCredential,       # SQL Server login (SQL authentication)
    [pscredential] $WindowsCredential,   # domain account for Windows auth (net-only impersonation)

    [datetime] $StartDate = (Get-Date).AddDays(-30).Date,
    [datetime] $EndDate   = (Get-Date),

    [string]   $StartTime,
    [string]   $EndTime,

    [Parameter(ParameterSetName = 'Report')]
    [string[]] $RestoreType,

    [string]   $NameLike,

    [string]   $SessionTable,
    [string]   $TypeColumn,

    [string]   $OutputFolder = (Join-Path $PSScriptRoot 'output'),

    # Also render a PDF copy of the report (via headless Edge/Chrome).
    [switch]   $Pdf,

    # Relative reporting window - convenient for scheduled runs. Overrides the
    # default -StartDate (Daily=last 24h, Weekly=7d, Monthly=1 month, Yearly=1yr).
    [ValidateSet('Daily','Weekly','Monthly','Yearly')]
    [string]   $Period,

    # --- Email delivery (optional) ---
    [string[]]     $EmailTo,          # one or more recipients; enables emailing
    [string]       $EmailFrom,        # sender address (defaults to a no-reply on this host)
    [string]       $SmtpServer,       # SMTP relay host
    [int]          $SmtpPort = 25,
    [switch]       $SmtpUseSsl,
    [pscredential] $SmtpCredential,   # omit for anonymous relay
    [string]       $EmailSubject,     # defaults to a summary line

    [Parameter(ParameterSetName = 'Discover')]
    [switch]   $Discover,

    # Generate the CSV + HTML from synthetic data (no DB) - to preview the report format.
    [switch]   $DemoData
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:ImpToken = $null   # net-only impersonation token when a domain account is used

# A relative -Period sets the reporting window automatically (unless an explicit
# -StartDate was supplied). Ideal for recurring scheduled runs.
if ($Period -and -not $PSBoundParameters.ContainsKey('StartDate')) {
    $now = Get-Date
    switch ($Period) {
        'Daily'   { $StartDate = $now.AddDays(-1)   }
        'Weekly'  { $StartDate = $now.AddDays(-7)   }
        'Monthly' { $StartDate = $now.AddMonths(-1) }
        'Yearly'  { $StartDate = $now.AddYears(-1)  }
    }
    if (-not $PSBoundParameters.ContainsKey('EndDate')) { $EndDate = $now }
}

# Accept recipients as an array or a comma/semicolon-separated string (a
# scheduled task passes them as one argument).
if ($EmailTo) {
    $EmailTo = @($EmailTo | ForEach-Object { $_ -split '[;,]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# ---------------------------------------------------------------------------
# 0. Version-sensitive schema mapping (the ONLY thing you may need to adjust
#    after -Discover). Override Table/TypeCol at run time with
#    -SessionTable / -TypeColumn without editing this file.
#
#    Verified against a Veeam ONE POC build (2026-08-08): restore activity is
#    stored per restored item in [monitor].[BpRestoreItem] in the *VeeamONE*
#    database. The previous mapping ([dbo].[Sessions]) does not exist on Veeam
#    ONE and is why the old report returned no data.
# ---------------------------------------------------------------------------
$script:Schema = @{
    Table        = '[monitor].[BpRestoreItem]'  # restored-item table in the Veeam ONE monitor DB
    IdCol        = 'uid'
    NameCol      = 'item_name'          # restored object (e.g. C:\NASTEST, VM name, DB name)
    TypeCol      = 'item_type'          # restore operation type (numeric enum or text)
    DestCol      = 'destination'        # restore destination path
    DestHostCol  = 'destination_host'   # host the item was restored to
    SizeCol      = 'item_size'          # restored size in bytes
    StartCol     = 'start_time'
    EndCol       = 'finish_time'
    StateCol     = 'state'              # outcome code (see $script:StateMap below)
    ReasonCol    = 'message'            # status / error message
    SessionCol   = 'restore_session_uid'
}
if ($SessionTable) { $script:Schema.Table   = $SessionTable }
if ($TypeColumn)   { $script:Schema.TypeCol = $TypeColumn }

# Outcome mapping for [state] in [monitor].[BpRestoreItem]. Confirmed on the
# Veeam ONE build: 2 = Success, 3 = Failed, 6 = Skipped. If -Discover shows a
# different convention on your build, adjust here. When [state] is null/unknown
# the report falls back to classifying the [message] text, so restores still get
# a sensible status either way.
$script:StateMap = @{
    2 = 'Success'
    3 = 'Failed'
    6 = 'Skipped'
}

# Optional friendly labels for the numeric [item_type]. Unknown values are shown
# as-is. Populate this after -Discover reveals the item_type values on your build.
$script:ItemTypeMap = @{
}

# ---------------------------------------------------------------------------
# 1. Resolve the Veeam ONE DB connection (registry auto-detect)
# ---------------------------------------------------------------------------
function Resolve-VeeamOneSqlConnection {
    param([string]$SqlServer, [string]$Database)

    if (-not $SqlServer -or -not $Database) {
        $regPaths = @(
            'HKLM:\SOFTWARE\Veeam\Veeam ONE Monitor',
            'HKLM:\SOFTWARE\Veeam\Veeam ONE Reporting',
            'HKLM:\SOFTWARE\Veeam\Veeam ONE Settings',
            'HKLM:\SOFTWARE\Wow6432Node\Veeam\Veeam ONE Monitor',
            'HKLM:\SOFTWARE\Wow6432Node\Veeam\Veeam ONE Reporting'
        )
        function Get-Prop { param($obj, [string[]]$names)
            foreach ($n in $names) {
                if ($obj.PSObject.Properties[$n] -and $obj.$n) { return $obj.$n }
            }
            return $null
        }
        foreach ($rp in $regPaths) {
            if (-not (Test-Path $rp)) { continue }
            $k = Get-ItemProperty -Path $rp -ErrorAction SilentlyContinue
            if (-not $k) { continue }
            if (-not $SqlServer) {
                $srv  = Get-Prop $k @('SqlServerName','DatabaseServer','SqlServer')
                $inst = Get-Prop $k @('SqlInstanceName','DatabaseInstance','SqlInstance')
                if ($srv) {
                    if ($inst -and $inst -ne 'MSSQLSERVER') { $SqlServer = "$srv\$inst" } else { $SqlServer = $srv }
                }
            }
            if (-not $Database) {
                $Database = Get-Prop $k @('SqlDatabaseName','DatabaseName','SqlDatabase')
            }
            if ($SqlServer -and $Database) { break }
        }
    }

    if (-not $SqlServer) { $SqlServer = '.' }
    if (-not $Database)  { $Database  = 'VeeamONE' }
    [pscustomobject]@{ Server = $SqlServer; Database = $Database }
}

function Get-SqlConnectionString {
    param([string]$Server, [string]$Database, [pscredential]$Cred)
    $base = "Server=$Server;Database=$Database;Application Name=VeeamOneRestoreReport;Connect Timeout=30;"
    if ($Cred) {
        $u = $Cred.UserName
        $p = $Cred.GetNetworkCredential().Password
        return "$base User ID=$u;Password=$p;"
    }
    return "$base Integrated Security=SSPI;"
}

function Enter-DomainImpersonation {
    # LogonUser with LOGON32_LOGON_NEW_CREDENTIALS (net-only) - the token authenticates
    # to the remote SQL server as the domain account, while local disk I/O stays as the
    # current user. Same mechanism as "runas /netonly".
    param([pscredential]$Credential)
    if (-not $Credential) { return $null }
    if (-not ([System.Management.Automation.PSTypeName]'VeeamOneRestore.NativeLogon').Type) {
        Add-Type -UsingNamespace 'Microsoft.Win32.SafeHandles' -Namespace 'VeeamOneRestore' -Name 'NativeLogon' -MemberDefinition @'
[DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern bool LogonUser(string user, string domain, string password,
    int logonType, int logonProvider, out SafeAccessTokenHandle token);
'@
    }
    $user = $Credential.UserName; $domain = $null
    if     ($user -like '*\*') { $p = $user.Split('\', 2); $domain = $p[0]; $user = $p[1] }
    elseif ($user -like '*@*') { $domain = $null }   # UPN: pass whole thing as user
    $pw    = $Credential.GetNetworkCredential().Password
    $token = $null
    $ok = [VeeamOneRestore.NativeLogon]::LogonUser($user, $domain, $pw, 9, 3, [ref]$token)  # 9=NEW_CREDENTIALS, 3=WINNT50
    if (-not $ok) {
        $code = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "LogonUser failed for '$($Credential.UserName)' (Win32 error $code). Check the domain\username and password."
    }
    return $token
}

function Invoke-VeeamSql {
    param([string]$ConnectionString, [string]$Query, [hashtable]$Params)
    $ctx = $null
    if ($script:ImpToken) {
        $wi  = New-Object System.Security.Principal.WindowsIdentity($script:ImpToken.DangerousGetHandle())
        $ctx = $wi.Impersonate()
    }
    $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = 300
        if ($Params) {
            foreach ($kv in $Params.GetEnumerator()) {
                $val = $kv.Value; if ($null -eq $val) { $val = [DBNull]::Value }
                [void]$cmd.Parameters.AddWithValue("@$($kv.Key)", $val)
            }
        }
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $table   = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
        return ,$table
    }
    finally { $conn.Close(); $conn.Dispose(); if ($ctx) { $ctx.Undo() } }
}

# ---------------------------------------------------------------------------
# 2. Build the session SELECT from the schema mapping (version-sensitive block)
# ---------------------------------------------------------------------------
function Get-RestoreSessionQuery {
    param([hashtable]$S, [string]$TypeFilter, [string]$NameFilter)
    $sql = @"
SELECT
    s.[$($S.IdCol)]       AS ItemId,
    s.[$($S.NameCol)]     AS SessionName,
    s.[$($S.TypeCol)]     AS RestoreType,
    s.[$($S.DestCol)]     AS Destination,
    s.[$($S.DestHostCol)] AS DestinationHost,
    s.[$($S.SizeCol)]     AS SizeBytes,
    s.[$($S.StartCol)]    AS StartDate,
    s.[$($S.EndCol)]      AS EndDate,
    s.[$($S.StateCol)]    AS ResultCode,
    s.[$($S.ReasonCol)]   AS Reason
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start
  AND s.[$($S.StartCol)] <  @end
  $TypeFilter
  $NameFilter
ORDER BY s.[$($S.StartCol)] DESC
"@
    return $sql
}

# ---------------------------------------------------------------------------
# 3. Helpers
# ---------------------------------------------------------------------------
function ConvertTo-StatusText {
    # Map the restore item's [state] to Success / Warning / Failed. Uses the
    # (overridable) $script:StateMap for numeric codes, recognises text states,
    # and finally falls back to classifying the [message] text so a restore
    # still gets a sensible status even if the state convention differs.
    param($ResultCode, $Message)

    if ($null -ne $ResultCode -and $ResultCode -isnot [DBNull]) {
        $n = 0
        if ([int]::TryParse([string]$ResultCode, [ref]$n)) {
            if ($script:StateMap.ContainsKey($n)) { return $script:StateMap[$n] }
        }
        else {
            switch -Regex ([string]$ResultCode) {
                '^(?i)success'      { return 'Success' }
                '^(?i)warn'         { return 'Warning' }
                '^(?i)(fail|error)' { return 'Failed'  }
            }
        }
    }

    if ($Message -and $Message -isnot [DBNull] -and [string]$Message -ne '') {
        switch -Regex ([string]$Message) {
            '(?i)(error|fail|unable|cannot|denied)' { return 'Failed'  }
            '(?i)skip'                              { return 'Skipped' }
            '(?i)warn'                              { return 'Warning' }
            '(?i)(success|completed|finished)'      { return 'Success' }
        }
    }
    return 'Unknown'
}

function ConvertTo-RestoreType {
    param($TypeValue)
    if ($null -eq $TypeValue -or $TypeValue -is [DBNull]) { return '' }
    $n = 0
    if ([int]::TryParse([string]$TypeValue, [ref]$n) -and $script:ItemTypeMap.ContainsKey($n)) {
        return $script:ItemTypeMap[$n]
    }
    return [string]$TypeValue
}

function Format-Size {
    param($Bytes)
    if ($null -eq $Bytes -or $Bytes -is [DBNull]) { return '' }
    $b = 0.0
    if (-not [double]::TryParse([string]$Bytes, [ref]$b) -or $b -le 0) { return '' }
    $u = 'B','KB','MB','GB','TB','PB'; $i = 0
    while ($b -ge 1024 -and $i -lt $u.Count - 1) { $b /= 1024; $i++ }
    '{0:N2} {1}' -f $b, $u[$i]
}

function Format-Elapsed {
    param([datetime]$Start, $End)
    if (-not $End -or $End -eq [datetime]::MinValue) { return '' }
    $ts = (New-TimeSpan -Start $Start -End $End)
    if ($ts.Ticks -lt 0) { return '' }
    '{0:00}:{1:00}:{2:00}' -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds
}

function Test-InTimeWindow {
    param([datetime]$When, [string]$From, [string]$To)
    if (-not $From -and -not $To) { return $true }
    $t   = $When.TimeOfDay
    $f   = if ($From) { [TimeSpan]::Parse($From) } else { [TimeSpan]::Zero }
    $til = if ($To)   { [TimeSpan]::Parse($To)   } else { [TimeSpan]::FromDays(1) }
    if ($f -le $til) { return ($t -ge $f -and $t -le $til) }        # same-day window
    return ($t -ge $f -or $t -le $til)                              # crosses midnight
}

# Encode text for safe HTML output
function Get-Html { param([string]$s) if ($null -eq $s) { return '' } [System.Web.HttpUtility]::HtmlEncode($s) }

# Render the HTML report to PDF using headless Edge/Chrome (no extra modules).
function ConvertTo-Pdf {
    param([string]$HtmlPath, [string]$PdfPath)
    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )
    $exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $exe) {
        throw "No Microsoft Edge or Google Chrome found to render the PDF. Install Edge/Chrome, or use the CSV/HTML output."
    }
    $uri = ([System.Uri]$HtmlPath).AbsoluteUri
    if (Test-Path $PdfPath) { Remove-Item $PdfPath -Force }
    # Newer Chromium uses --headless=new; older builds use --headless. Try both.
    foreach ($headless in '--headless=new', '--headless') {
        $argLine = '{0} --disable-gpu --no-pdf-header-footer --print-to-pdf="{1}" "{2}"' -f $headless, $PdfPath, $uri
        Start-Process -FilePath $exe -ArgumentList $argLine -Wait -WindowStyle Hidden
        if (Test-Path $PdfPath) { return $PdfPath }
    }
    throw "PDF render command completed but produced no file (browser: $exe)."
}

# Compact, Outlook-friendly HTML summary for the email body (full report attached).
function New-EmailBodyHtml {
    param([hashtable]$Stats, [string]$Range, [string]$Db, [hashtable]$Colors)
    $cell = "padding:6px 14px;border:1px solid #e3e8ee;font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    $num  = { param($v,$c) "<td style='$cell text-align:center;font-weight:700;color:$c;'>$v</td>" }
    @"
<div style='font-family:Segoe UI,Arial,sans-serif;color:#1a2733;'>
  <h2 style='margin:0 0 4px;'>VeeamONE - Restoration Status Percentage Report</h2>
  <div style='color:#6b7885;font-size:13px;margin-bottom:14px;'>$Range &middot; DB: $Db</div>
  <table style='border-collapse:collapse;margin-bottom:12px;'>
    <tr>
      <td style='$cell color:#6b7885;'>Total</td>
      <td style='$cell color:#6b7885;'>Success</td>
      <td style='$cell color:#6b7885;'>Failed</td>
      <td style='$cell color:#6b7885;'>Skipped</td>
      <td style='$cell color:#6b7885;'>Success rate</td>
    </tr>
    <tr>
      $(& $num $Stats.Total '#1a2733')
      $(& $num $Stats.Success $Colors.Success)
      $(& $num $Stats.Failed $Colors.Failed)
      $(& $num $Stats.Skipped $Colors.Skipped)
      $(& $num ("{0}%" -f $Stats.Rate) '#1a2733')
    </tr>
  </table>
  <div style='color:#6b7885;font-size:12px;'>The full report is attached (HTML/PDF/CSV). Success rate excludes skipped restores.</div>
</div>
"@
}

function Send-ReportEmail {
    param(
        [string[]]$To, [string]$From, [string]$SmtpServer, [int]$SmtpPort,
        [switch]$UseSsl, [pscredential]$Cred, [string]$Subject,
        [string]$BodyHtml, [string[]]$Attachments
    )
    if (-not $SmtpServer) { throw "-EmailTo was supplied but -SmtpServer is missing." }
    if (-not $From) { $From = "VeeamONE-Restore-Report@$($env:COMPUTERNAME)" }
    $params = @{
        To         = $To
        From       = $From
        Subject    = $Subject
        Body       = $BodyHtml
        BodyAsHtml = $true
        SmtpServer = $SmtpServer
        Port       = $SmtpPort
        ErrorAction = 'Stop'
    }
    if ($UseSsl) { $params.UseSsl = $true }
    if ($Cred)   { $params.Credential = $Cred }
    $existing = @($Attachments | Where-Object { $_ -and (Test-Path $_) })
    if ($existing.Count -gt 0) { $params.Attachments = $existing }
    Send-MailMessage @params
}

# ---------------------------------------------------------------------------
# 4. Connect
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

if ($DemoData) {
    $connInfo = [pscustomobject]@{ Server = '(demo)'; Database = 'VeeamOne (demo data)' }
    Write-Host "DEMO MODE - generating synthetic report (no database queried)." -ForegroundColor Magenta
}
else {
    $connInfo = Resolve-VeeamOneSqlConnection -SqlServer $SqlServer -Database $Database
    $connStr  = Get-SqlConnectionString -Server $connInfo.Server -Database $connInfo.Database -Cred $SqlCredential
    if ($WindowsCredential) {
        $script:ImpToken = Enter-DomainImpersonation -Credential $WindowsCredential
        Write-Host "Auth       : Windows (domain account $($WindowsCredential.UserName), net-only)" -ForegroundColor Cyan
    }
    elseif ($SqlCredential) { Write-Host "Auth       : SQL login ($($SqlCredential.UserName))" -ForegroundColor Cyan }
    else                    { Write-Host "Auth       : Windows (current user)" -ForegroundColor Cyan }
    Write-Host "Veeam ONE DB: $($connInfo.Server) / $($connInfo.Database)" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# 5. Discovery mode
# ---------------------------------------------------------------------------
if ($Discover) {
    $S = $script:Schema

    Write-Host "`n=== Candidate restore tables (name + column list) ===" -ForegroundColor Yellow
    Write-Host "The report reads $($S.Table). If your build differs, pick the right" -ForegroundColor Gray
    Write-Host "table here and set `$script:Schema.Table (or pass -SessionTable).`n" -ForegroundColor Gray
    $cols = Invoke-VeeamSql -ConnectionString $connStr -Query @'
SELECT  t.TABLE_SCHEMA + '.' + t.TABLE_NAME              AS TableName,
        STUFF((SELECT ', ' + c.COLUMN_NAME
               FROM INFORMATION_SCHEMA.COLUMNS c
               WHERE c.TABLE_SCHEMA = t.TABLE_SCHEMA AND c.TABLE_NAME = t.TABLE_NAME
               ORDER BY c.ORDINAL_POSITION
               FOR XML PATH('')), 1, 2, '')              AS Columns
FROM INFORMATION_SCHEMA.TABLES t
WHERE t.TABLE_NAME LIKE '%Restore%' OR t.TABLE_NAME LIKE '%Recovery%'
ORDER BY TableName
'@
    $cols | Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host

    # Distinct [state] values + a sample message each -> confirm the Success/
    # Warning/Failed mapping in $script:StateMap.
    Write-Host "=== Distinct '$($S.StateCol)' values in $($S.Table) (last 90 days) ===" -ForegroundColor Yellow
    Write-Host "Match each state code to its sample message, then adjust `$script:StateMap if needed.`n" -ForegroundColor Gray
    try {
        $stateSql = @"
SELECT s.[$($S.StateCol)]      AS StateCode,
       COUNT(*)                AS Items,
       SUM(CASE WHEN s.[$($S.EndCol)] IS NULL THEN 1 ELSE 0 END) AS StillRunning,
       MAX(s.[$($S.ReasonCol)]) AS SampleMessage
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start
GROUP BY s.[$($S.StateCol)]
ORDER BY Items DESC
"@
        $states = Invoke-VeeamSql -ConnectionString $connStr -Params @{ start = (Get-Date).AddDays(-90) } -Query $stateSql
        $states | Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    }
    catch {
        Write-Warning "Could not enumerate states from $($S.Table): $($_.Exception.Message)"
    }

    # Distinct [item_type] values + a sample item name each -> populate
    # $script:ItemTypeMap for friendly labels, and filter with -RestoreType.
    Write-Host "=== Distinct '$($S.TypeCol)' values in $($S.Table) (last 90 days) ===" -ForegroundColor Yellow
    Write-Host "Map each type to a sample restored object, then filter with -RestoreType.`n" -ForegroundColor Gray
    try {
        $typeSql = @"
SELECT s.[$($S.TypeCol)]      AS RestoreType,
       COUNT(*)               AS Items,
       MIN(s.[$($S.StartCol)]) AS FirstSeen,
       MAX(s.[$($S.StartCol)]) AS LastSeen,
       MAX(s.[$($S.NameCol)])  AS SampleObject
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start
GROUP BY s.[$($S.TypeCol)]
ORDER BY Items DESC
"@
        $types = Invoke-VeeamSql -ConnectionString $connStr -Params @{ start = (Get-Date).AddDays(-90) } -Query $typeSql
        $types | Format-Table -AutoSize | Out-String | Write-Host
    }
    catch {
        Write-Warning "Could not enumerate types from $($S.Table): $($_.Exception.Message)"
        Write-Warning "Confirm the table from the list above and set `$script:Schema.Table (or use -SessionTable)."
    }
    Write-Host "Discovery complete." -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# 6. Build the record set - synthetic (demo) or queried from the Veeam ONE DB
# ---------------------------------------------------------------------------
$records = New-Object System.Collections.Generic.List[object]

if ($DemoData) {
    $types = 'Full VM Restore','File-Level Restore','Instant VM Recovery','Application Item Restore','Disk Restore','Guest File Restore'
    $objs  = 'SQLPROD01','WEB-APP-02','FILESRV01','DC01','ERP-DB','EXCH-MBX01','APP-NODE-03','C:\NASTEST'
    $hosts = 'VoneSrvP01','ESX-PROD-01','HV-NODE-02','FILESRV01'
    $rng   = New-Object System.Random 20260806
    for ($d = 0; $d -lt 14; $d++) {
        $day = (Get-Date).Date.AddDays(-13 + $d)
        foreach ($i in 0..($rng.Next(1,5))) {
            $stt = $day.AddHours($rng.Next(8,20)).AddMinutes($rng.Next(0,59))
            $en  = $stt.AddSeconds($rng.Next(60, 4200))
            $roll = $rng.NextDouble()
            $status = if ($roll -gt 0.90) { 'Failed' } elseif ($roll -gt 0.80) { 'Skipped' } else { 'Success' }
            $obj = $objs[$rng.Next(0,$objs.Count)]
            $sz  = if ($status -eq 'Success') { [long]($rng.Next(50, 20000)) * 1MB } else { 0 }
            $records.Add([pscustomobject]@{
                RestoreType     = $types[$rng.Next(0,$types.Count)]
                ObjectName      = $obj
                DestinationHost = $hosts[$rng.Next(0,$hosts.Count)]
                Status          = $status
                StartDate       = $stt.ToString('yyyy-MM-dd')
                StartTime       = $stt.ToString('HH:mm:ss')
                EndDate         = $en.ToString('yyyy-MM-dd')
                EndTime         = $en.ToString('HH:mm:ss')
                ElapsedTime     = Format-Elapsed -Start $stt -End $en
                Size            = Format-Size $sz
                Reason          = if ($status -eq 'Failed') { 'Restore failed: sample error (demo)' } elseif ($status -eq 'Skipped') { 'Item skipped (demo)' } else { '' }
            }) | Out-Null
        }
    }
    $StartDate = (Get-Date).Date.AddDays(-13); $EndDate = (Get-Date)
}
else {

    $typeFilter = ''
    $nameFilter = if ($NameLike) { "AND s.[$($script:Schema.NameCol)] LIKE @namelike" } else { '' }
    $sqlParams  = @{}

    if ($RestoreType -and $RestoreType.Count -gt 0) {
        $ors = for ($idx = 0; $idx -lt $RestoreType.Count; $idx++) {
            $sqlParams["rt$idx"] = $RestoreType[$idx]
            "s.[$($script:Schema.TypeCol)] LIKE @rt$idx"
        }
        $typeFilter = "AND (" + ($ors -join ' OR ') + ")"
    }
    if ($NameLike) { $sqlParams['namelike'] = $NameLike }

    $sessionSql = Get-RestoreSessionQuery -S $script:Schema -TypeFilter $typeFilter -NameFilter $nameFilter

    # If EndDate carries a time-of-day (picked in the GUI) honor it exactly;
    # a date-only value covers the whole day.
    $endExclusive = if ($EndDate.TimeOfDay -eq [TimeSpan]::Zero) { $EndDate.Date.AddDays(1) } else { $EndDate }
    $sqlParams['start'] = $StartDate
    $sqlParams['end']   = $endExclusive

    Write-Host "Querying restore sessions $($StartDate.ToString('yyyy-MM-dd')) .. $($EndDate.ToString('yyyy-MM-dd'))$(if ($RestoreType) { " (type ~ $($RestoreType -join ', '))" })..." -ForegroundColor Cyan
    try {
        $sessions = Invoke-VeeamSql -ConnectionString $connStr -Query $sessionSql -Params $sqlParams
    }
    catch {
        Write-Warning "Restore-session query failed against $($script:Schema.Table)."
        Write-Warning "Message: $($_.Exception.Message)"
        Write-Warning "Run -Discover to confirm the correct table/columns, then adjust `$script:Schema (or pass -SessionTable / -TypeColumn)."
        return
    }

    # ---------------------------------------------------------------------------
    # 7. Shape the rows (one row per restore session)
    # ---------------------------------------------------------------------------
    foreach ($s in $sessions.Rows) {
        $start = [datetime]$s['StartDate']
        if (-not (Test-InTimeWindow -When $start -From $StartTime -To $EndTime)) { continue }

        $end = if ($s['EndDate'] -is [datetime]) { [datetime]$s['EndDate'] } else { $null }

        # A restore still in progress has no finish_time; label it Running
        # rather than mapping its state code to an outcome.
        $status = if (-not $end) { 'Running' } else { ConvertTo-StatusText $s['ResultCode'] $s['Reason'] }

        $records.Add([pscustomobject]@{
            RestoreType     = ConvertTo-RestoreType $s['RestoreType']
            ObjectName      = [string]$s['SessionName']
            DestinationHost = if ($s['DestinationHost'] -isnot [DBNull]) { [string]$s['DestinationHost'] } else { '' }
            Status          = $status
            StartDate       = $start.ToString('yyyy-MM-dd')
            StartTime       = $start.ToString('HH:mm:ss')
            EndDate         = if ($end) { $end.ToString('yyyy-MM-dd') } else { '' }
            EndTime         = if ($end) { $end.ToString('HH:mm:ss') } else { '' }
            ElapsedTime     = Format-Elapsed -Start $start -End $end
            Size            = Format-Size $s['SizeBytes']
            Reason          = if ($s['Reason'] -isnot [DBNull]) { [string]$s['Reason'] } else { '' }
        }) | Out-Null
    }

}  # end else (Veeam ONE DB path)

Write-Host "Collected $($records.Count) restore session(s)." -ForegroundColor Green
if ($records.Count -eq 0) {
    Write-Warning "No restore sessions matched. Check the date range, -RestoreType filter, and (via -Discover) the table/columns."
    return
}

# ---------------------------------------------------------------------------
# 8. Export CSV
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath = Join-Path $OutputFolder "VeeamONE-Restore-Report-$stamp.csv"
$records |
    Select-Object RestoreType, ObjectName, DestinationHost, Status, StartDate, StartTime, EndDate, EndTime, ElapsedTime, Size, Reason |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV  : $csvPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 9. Build the HTML report (self-contained, hand-rendered SVG - no CDN/JS libs)
# ---------------------------------------------------------------------------
$byStatus = $records | Group-Object Status | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Count = $_.Count } }
$total    = $records.Count
$succ     = ($byStatus | Where-Object Name -eq 'Success' | ForEach-Object Count) ; if (-not $succ) { $succ = 0 }
$warn     = ($byStatus | Where-Object Name -eq 'Warning' | ForEach-Object Count) ; if (-not $warn) { $warn = 0 }
$fail     = ($byStatus | Where-Object Name -eq 'Failed'  | ForEach-Object Count) ; if (-not $fail) { $fail = 0 }
$skip     = ($byStatus | Where-Object Name -eq 'Skipped' | ForEach-Object Count) ; if (-not $skip) { $skip = 0 }
$other    = $total - $succ - $warn - $fail - $skip
# Success rate is measured over restores that actually ran (skipped items are
# neither a success nor a failure, so they are excluded from the denominator).
$rated    = $total - $skip
$rate     = if ($rated) { [math]::Round(($succ / $rated) * 100, 1) } else { 0 }

$statusColors = @{ Success = '#00b336'; Warning = '#ffb300'; Failed = '#e5202e'; Skipped = '#8a63d2'; Running = '#2b8fd6'; Unknown = '#8a8a8a' }

# --- SVG donut for status distribution ---
function New-DonutSvg {
    param([hashtable]$Data, [hashtable]$Colors, [int]$Size = 220)
    $sum = ($Data.Values | Measure-Object -Sum).Sum
    if (-not $sum) { return '' }
    $cx = $Size / 2; $cy = $Size / 2; $r = ($Size / 2) - 12; $inner = $r * 0.6
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg viewBox='0 0 $Size $Size' width='$Size' height='$Size' role='img'>")
    $angle = -90.0
    foreach ($k in ($Data.Keys | Sort-Object)) {
        $v = $Data[$k]; if ($v -le 0) { continue }
        $sweep = 360.0 * $v / $sum
        $a0 = $angle * [math]::PI / 180; $a1 = ($angle + $sweep) * [math]::PI / 180
        $x0 = $cx + $r * [math]::Cos($a0); $y0 = $cy + $r * [math]::Sin($a0)
        $x1 = $cx + $r * [math]::Cos($a1); $y1 = $cy + $r * [math]::Sin($a1)
        $large = if ($sweep -gt 180) { 1 } else { 0 }
        $col = if ($Colors.ContainsKey($k)) { $Colors[$k] } else { '#8a8a8a' }
        $path = "M $cx $cy L {0:N3} {1:N3} A $r $r 0 $large 1 {2:N3} {3:N3} Z" -f $x0,$y0,$x1,$y1
        [void]$sb.Append("<path d='$path' fill='$col'/>")
        $angle += $sweep
    }
    [void]$sb.Append("<circle cx='$cx' cy='$cy' r='$inner' fill='var(--card)'/>")
    [void]$sb.Append("<text x='$cx' y='$($cy-4)' text-anchor='middle' font-size='30' font-weight='700' fill='var(--fg)'>$rate%</text>")
    [void]$sb.Append("<text x='$cx' y='$($cy+18)' text-anchor='middle' font-size='12' fill='var(--muted)'>success</text>")
    [void]$sb.Append('</svg>')
    $sb.ToString()
}
$donut = New-DonutSvg -Data @{ Success = $succ; Warning = $warn; Failed = $fail; Skipped = $skip; Unknown = $other } -Colors $statusColors

# --- SVG daily trend (stacked bars) ---
function New-TrendSvg {
    param($Records, [hashtable]$Colors)
    $days = @($Records | Group-Object StartDate | Sort-Object Name)
    if ($days.Count -eq 0) { return '' }
    $w = [math]::Max(360, $days.Count * 46 + 40); $h = 180; $padB = 26; $padT = 10; $plot = $h - $padB - $padT
    $max = ($days | ForEach-Object { $_.Group.Count } | Measure-Object -Maximum).Maximum
    if (-not $max) { $max = 1 }
    $bw = 28; $gap = ($w - 40) / [math]::Max(1,$days.Count)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg viewBox='0 0 $w $h' width='100%' height='$h' preserveAspectRatio='xMinYMin meet'>")
    $i = 0
    foreach ($d in $days) {
        $x = 30 + $i * $gap + ($gap - $bw)/2
        $y = $padT + $plot
        foreach ($st in 'Failed','Warning','Skipped','Success') {
            $c = @($d.Group | Where-Object Status -eq $st).Count
            if ($c -le 0) { continue }
            $bh = $plot * $c / $max
            $y -= $bh
            $col = $Colors[$st]
            [void]$sb.Append(("<rect x='{0:N1}' y='{1:N1}' width='$bw' height='{2:N1}' fill='$col'/>" -f $x,$y,$bh))
        }
        $lbl = ($d.Name).Substring(5)  # MM-dd
        [void]$sb.Append("<text x='{0:N1}' y='$($h-8)' text-anchor='middle' font-size='10' fill='var(--muted)'>$lbl</text>" -f ($x + $bw/2))
        $i++
    }
    [void]$sb.Append('</svg>')
    $sb.ToString()
}
$trend = New-TrendSvg -Records $records -Colors $statusColors

# --- table rows ---
$rowsHtml = [System.Text.StringBuilder]::new()
foreach ($r in $records) {
    $sc = if ($statusColors.ContainsKey($r.Status)) { $statusColors[$r.Status] } else { '#8a8a8a' }
    [void]$rowsHtml.Append("<tr>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.RestoreType)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.ObjectName)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.DestinationHost)</td>")
    [void]$rowsHtml.Append("<td><span class='badge' style='background:$sc'>$(Get-Html $r.Status)</span></td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.StartDate) $(Get-Html $r.StartTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.EndDate) $(Get-Html $r.EndTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.ElapsedTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Size)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Reason)</td>")
    [void]$rowsHtml.Append("</tr>")
}

$legend = @"
<div class='legend'>
  <span><i style='background:$($statusColors.Success)'></i>Success ($succ)</span>
  $(if ($warn -gt 0) { "<span><i style='background:$($statusColors.Warning)'></i>Warning ($warn)</span>" })
  <span><i style='background:$($statusColors.Failed)'></i>Failed ($fail)</span>
  $(if ($skip -gt 0) { "<span><i style='background:$($statusColors.Skipped)'></i>Skipped ($skip)</span>" })
  $(if ($other -gt 0) { "<span><i style='background:$($statusColors.Unknown)'></i>Other ($other)</span>" })
</div>
"@

$rangeText = "$($StartDate.ToString('yyyy-MM-dd'))  to  $($EndDate.ToString('yyyy-MM-dd'))"
if ($StartTime -or $EndTime) { $rangeText += "   (window $([string]$StartTime)-$([string]$EndTime))" }

$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Veeam ONE Restore Report</title>
<style>
  :root { --bg:#f4f6f8; --card:#ffffff; --fg:#1a2733; --muted:#6b7885; --line:#e3e8ee; --accent:#00b336; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#0f1620; --card:#17212e; --fg:#e8eef4; --muted:#8ea0b0; --line:#243546; }
  }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI',Roboto,Arial,sans-serif; background:var(--bg); color:var(--fg); }
  .wrap { max-width:1200px; margin:0 auto; padding:24px; }
  header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
  h1 { font-size:20px; margin:0; }
  .sub { color:var(--muted); font-size:13px; }
  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:14px; margin:20px 0; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:16px; }
  .card .n { font-size:26px; font-weight:700; }
  .card .l { color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }
  .charts { display:grid; grid-template-columns:240px 1fr; gap:16px; margin-bottom:22px; }
  @media (max-width:720px){ .charts{ grid-template-columns:1fr; } }
  .panel { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:16px; }
  .panel h2 { font-size:14px; margin:0 0 10px; color:var(--muted); font-weight:600; }
  .donutwrap { display:flex; flex-direction:column; align-items:center; gap:10px; }
  .legend { display:flex; flex-wrap:wrap; gap:12px; font-size:12px; color:var(--muted); }
  .legend i { display:inline-block; width:10px; height:10px; border-radius:2px; margin-right:5px; vertical-align:middle; }
  .tablewrap { background:var(--card); border:1px solid var(--line); border-radius:10px; overflow-x:auto; }
  table { border-collapse:collapse; width:100%; font-size:13px; }
  th,td { padding:9px 12px; text-align:left; border-bottom:1px solid var(--line); white-space:nowrap; }
  th { position:sticky; top:0; background:var(--card); color:var(--muted); font-weight:600; }
  tr:hover td { background:rgba(0,179,54,.05); }
  .badge { color:#fff; padding:2px 9px; border-radius:10px; font-size:11px; font-weight:600; }
  footer { color:var(--muted); font-size:11px; margin-top:18px; text-align:center; }
</style>
</head>
<body>
<div class='wrap'>
  <header>
    <div>
      <h1>Veeam ONE - Restore Sessions Report</h1>
      <div class='sub'>Restore success / failure rate &middot; $rangeText &middot; DB: $(Get-Html $connInfo.Database)</div>
    </div>
    <div class='sub'>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
  </header>

  <div class='cards'>
    <div class='card'><div class='n'>$total</div><div class='l'>Total restores</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Success)'>$succ</div><div class='l'>Success</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Failed)'>$fail</div><div class='l'>Failed</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Skipped)'>$skip</div><div class='l'>Skipped</div></div>
    <div class='card'><div class='n'>$rate%</div><div class='l'>Success rate</div></div>
  </div>

  <div class='charts'>
    <div class='panel'>
      <h2>Success / Failure</h2>
      <div class='donutwrap'>$donut $legend</div>
    </div>
    <div class='panel'>
      <h2>Daily trend</h2>
      $trend
    </div>
  </div>

  <div class='tablewrap'>
    <table>
      <thead><tr>
        <th>Restore type</th><th>Object</th><th>Destination host</th><th>Status</th>
        <th>Start</th><th>End</th><th>Elapsed</th><th>Size</th><th>Reason</th>
      </tr></thead>
      <tbody>
        $($rowsHtml.ToString())
      </tbody>
    </table>
  </div>

  <footer>Veeam ONE Restore Report &middot; source data from the Veeam ONE database (read-only)</footer>
</div>
</body>
</html>
"@

$htmlPath = Join-Path $OutputFolder "VeeamONE-Restore-Report-$stamp.html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "HTML : $htmlPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 10. Optional PDF (headless Edge/Chrome)
# ---------------------------------------------------------------------------
$pdfPath = $null
if ($Pdf) {
    $pdfPath = Join-Path $OutputFolder "VeeamONE-Restore-Report-$stamp.pdf"
    try {
        ConvertTo-Pdf -HtmlPath $htmlPath -PdfPath $pdfPath | Out-Null
        Write-Host "PDF  : $pdfPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "PDF generation failed: $($_.Exception.Message)"
        $pdfPath = $null
    }
}

# ---------------------------------------------------------------------------
# 11. Optional email delivery
# ---------------------------------------------------------------------------
if ($EmailTo -and $EmailTo.Count -gt 0) {
    $subject = if ($EmailSubject) { $EmailSubject } else {
        "VeeamONE Restore Report - $rate% success ($total restores) - $($StartDate.ToString('yyyy-MM-dd'))..$($EndDate.ToString('yyyy-MM-dd'))"
    }
    $bodyHtml = New-EmailBodyHtml `
        -Stats @{ Total = $total; Success = $succ; Failed = $fail; Skipped = $skip; Rate = $rate } `
        -Range $rangeText -Db ([string]$connInfo.Database) -Colors $statusColors
    $attachments = @($csvPath, $htmlPath, $pdfPath) | Where-Object { $_ }
    try {
        Send-ReportEmail -To $EmailTo -From $EmailFrom -SmtpServer $SmtpServer -SmtpPort $SmtpPort `
            -UseSsl:$SmtpUseSsl -Cred $SmtpCredential -Subject $subject -BodyHtml $bodyHtml -Attachments $attachments
        Write-Host "Email: sent to $($EmailTo -join ', ') via $SmtpServer`:$SmtpPort" -ForegroundColor Green
    }
    catch {
        Write-Warning "Email delivery failed: $($_.Exception.Message)"
    }
}

Write-Host "`nSummary: $total restores | $succ success | $fail failed | $skip skipped | $rate% success rate" -ForegroundColor Cyan
[pscustomobject]@{ Csv = $csvPath; Html = $htmlPath; Pdf = $pdfPath; Total = $total; Success = $succ; Warning = $warn; Failed = $fail; Skipped = $skip; SuccessRate = $rate }
