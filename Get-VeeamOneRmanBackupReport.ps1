<#
.SYNOPSIS
    VeeamONE RMAN (standalone / unmanaged) backup success-rate report.

    Collects backup job SESSIONS (the RMAN plugin "rescans" and other backup
    runs) from the Veeam ONE monitoring database and exports a CSV plus a
    self-contained HTML report with a success / warning / failure rate chart -
    the same look and feel as the VeeamONE Restore report.

.DESCRIPTION
    Runs on (or next to) the Veeam ONE server. Reads the Veeam ONE database
    connection from the registry automatically (override with -SqlServer /
    -Database), connects with Windows Integrated authentication by default (or a
    domain account via net-only impersonation, or a SQL login via -SqlCredential),
    and queries the backup job sessions Veeam ONE collected from the monitored
    Veeam Backup & Replication servers.

    Each RMAN-plugin backup is triggered by an RMAN script on the database host,
    which makes the standalone/unmanaged job run ("rescan"). Every such run is a
    row in [monitor].[BpJobSession]. By default the report keeps only jobs whose
    name matches -JobNameLike ("%RMAN%"); pass -IncludeAllBackups to report on
    every backup job session alongside the RMAN ones.

    Output per session: Job name, Mode (Full/Incremental), Status, Start / End,
    Elapsed, Data read, Transferred, and the failure/warning message - plus a
    success/failure donut and a daily trend in the HTML report.

.PARAMETER SqlServer
    SQL Server instance hosting the Veeam ONE database, e.g.
    "VEEAMONE\VEEAMSQL2016". Auto-detected from the registry if omitted.

.PARAMETER Database
    Veeam ONE database name (default auto-detected, usually "VeeamONE").

.PARAMETER SqlCredential
    Optional PSCredential for SQL authentication. Omit to use the current
    Windows account (Integrated Security).

.PARAMETER WindowsCredential
    Optional domain account for Windows auth via net-only impersonation.

.PARAMETER StartDate / EndDate
    Reporting window on the session START time. Default: last 30 days.

.PARAMETER StartTime / EndTime
    Optional time-of-day window (e.g. nightly "22:00"-"06:00").

.PARAMETER JobNameLike
    SQL LIKE filter on the job name. Default "%RMAN%" (RMAN plugin jobs). Ignored
    when -IncludeAllBackups is set.

.PARAMETER IncludeAllBackups
    Report on every backup job session, not just the RMAN ones.

.PARAMETER OutputFolder
    Where to write the CSV + HTML (+ PDF). Default: .\output next to this script.

.PARAMETER Pdf
    Also render a PDF via headless Edge/Chrome.

.PARAMETER Period
    Relative window for scheduled runs: Daily / Weekly / Monthly / Yearly.

.PARAMETER EmailTo / EmailFrom / SmtpServer / SmtpPort / SmtpUseSsl / SmtpCredential / EmailSubject
    Optional e-mail delivery of the report (attachments: CSV + HTML + PDF).

.PARAMETER Discover
    Inspect the DB - list the session result codes and job types present - and exit.

.PARAMETER DemoData
    Build the report from synthetic data (no DB) to preview the format.

.EXAMPLE
    # Last 30 days of RMAN plugin backups
    .\Get-VeeamOneRmanBackupReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)

.EXAMPLE
    # Every backup (not just RMAN), specific month, e-mailed as PDF
    .\Get-VeeamOneRmanBackupReport.ps1 -IncludeAllBackups -StartDate 2026-07-01 -EndDate 2026-08-01 `
        -Pdf -EmailTo ops@contoso.com -SmtpServer smtp.contoso.com

.EXAMPLE
    .\Get-VeeamOneRmanBackupReport.ps1 -DemoData
#>
[CmdletBinding(DefaultParameterSetName = 'Report')]
param(
    [string]   $SqlServer,
    [string]   $Database,
    [pscredential] $SqlCredential,
    [pscredential] $WindowsCredential,

    [datetime] $StartDate = (Get-Date).AddDays(-30).Date,
    [datetime] $EndDate   = (Get-Date),

    [string]   $StartTime,
    [string]   $EndTime,

    [Parameter(ParameterSetName = 'Report')]
    [string]   $JobNameLike = '%RMAN%',
    [Parameter(ParameterSetName = 'Report')]
    [switch]   $IncludeAllBackups,

    # Skip the plugin-inventory drill-down (objects/databases under each job).
    # Use if the plugin tables are large/slow; the report still lists jobs, scans
    # (derived) and sizes, with the object column falling back to the job-name DB.
    [Parameter(ParameterSetName = 'Report')]
    [switch]   $SkipObjectDrilldown,

    [string]   $OutputFolder = (Join-Path $PSScriptRoot 'output'),

    [switch]   $Pdf,

    [ValidateSet('Daily','Weekly','Monthly','Yearly')]
    [string]   $Period,

    [string[]]     $EmailTo,
    [string]       $EmailFrom,
    [string]       $SmtpServer,
    [int]          $SmtpPort = 25,
    [switch]       $SmtpUseSsl,
    [pscredential] $SmtpCredential,
    [string]       $EmailSubject,

    [Parameter(ParameterSetName = 'Discover')]
    [switch]   $Discover,

    # Locate the VeeamONE table that stores per-DATABASE / per-object detail
    # (the individual databases inside each server/instance). Prints candidate
    # tables + columns + samples, then exits. The 6 job/session tables do NOT
    # carry database names - this finds where they live on your build.
    [Parameter(ParameterSetName = 'DiscoverObjects')]
    [switch]   $DiscoverObjects,

    [switch]   $DemoData
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:ImpToken = $null

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
if ($EmailTo) {
    $EmailTo = @($EmailTo | ForEach-Object { $_ -split '[;,]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# ---------------------------------------------------------------------------
# 0. Version-sensitive mapping (adjust after -Discover if your build differs).
#    Backup job sessions live in [monitor].[BpJobSession]; the [result] code is
#    2 = Success, 3 = Warning, 4 = Failed on the confirmed Veeam ONE build.
# ---------------------------------------------------------------------------
$script:Schema = @{
    Table       = '[monitor].[BpJobSession]'
    IdCol       = 'uid'
    JobNameCol  = 'job_name'
    JobTypeCol  = 'job_type'
    IsFullCol   = 'is_full'
    ResultCol   = 'result'
    StateCol    = 'state'
    StartCol    = 'start_time'
    DurationCol = 'duration'            # seconds
    # processed_used_size = the DB data actually read/processed by the scan (the
    # "readable DB size"). Use it instead of backedup_size, which is 0 for many
    # RMAN archive scans and hid the large (100s of GB / TB) scans.
    ReadCol     = 'processed_used_size'
    XferCol     = 'transferred_size'    # bytes written to target
    ReasonCol   = 'failure_message'
}
$script:ResultMap = @{
    2 = 'Success'
    3 = 'Warning'
    4 = 'Failed'
}

# Friendly Veeam job-type names. Numeric codes are derived from the Veeam ONE
# data (job_type + source_type + platform + name pattern). The classifier below
# reads the job NAME first (most reliable), then falls back to this numeric map.
$script:JobTypeNumMap = @{
    0='File / Object Storage Backup'; 1='VM Backup (VMware/virtual)'; 2='Replication'; 5='Backup to Tape'
    6='Backup to Tape'; 9='Replication (DR)'; 10='Microsoft SQL Log Backup (plugin)'
    13='Oracle RMAN Plugin - archived logs (unmanaged)'; 15='Oracle RMAN Plugin - full/incremental (unmanaged)'
    16='Backup Copy'; 17='Backup Copy'; 18='Veeam Agent Backup'; 21='NAS / File Share Backup'
    22='Nutanix AHV Backup'; 23='NAS Backup Copy'; 24='Backup Copy'; 25='Nutanix AHV Backup'
    31='Application DB Plugin Backup'; 32='SAP HANA Plugin (backint)'; 33='Oracle Plugin Backup'
    34='Oracle Plugin - redo/archived logs'; 35='Oracle Backup Copy'; 36='Oracle Backup Copy'; 37='Oracle Backup Copy'
}
function Get-JobTypeName {
    param([string]$JobName, $TypeCode)
    $n = "$JobName"
    $isCopy = ($n -match '^\s*CPY-') -or ($n -match '(?i)\bcopy\b')
    $label = $null
    if     ($n -match '(?i)RMAN')                                   { $mode = if ($n -match '(?i)arch') { ' - archived logs' } elseif ($n -match '(?i)full') { ' - full' } else { '' }; $label = "Oracle RMAN Plugin$mode (standalone/unmanaged)" }
    elseif ($n -match '(?i)backint|SAP\s*HANA')                     { $label = 'SAP HANA Plugin (backint)' }
    elseif ($n -match '(?i)Redo Log')                               { $label = 'Oracle Plugin - redo/archived logs' }
    elseif ($n -match '(?i)Oracle')                                 { $label = 'Oracle Plugin Backup' }
    elseif ($n -match '(?i)SQL Server Transaction Log')             { $label = 'Microsoft SQL Log Backup (plugin)' }
    elseif ($n -match '(?i)\bSQL\b' -or $n -match '(?i)(^|[-_])PG-'){ $label = 'Microsoft SQL Server (plugin/agent)' }
    elseif ($n -match '(?i)^\s*REP-|replicat')                      { $label = 'Replication' }
    elseif ($n -match '(?i)on Tape|[-_]TAPE')                       { $label = 'Backup to Tape' }
    elseif ($n -match '(?i)CIFS|NFS|\bNAS\b|Share|Filer')           { $label = 'NAS / File Share Backup' }
    elseif ($n -match '(?i)Nutanix')                                { $label = 'Nutanix AHV Backup' }
    if (-not $label) {
        $c = 0
        if ([int]::TryParse("$TypeCode", [ref]$c) -and $script:JobTypeNumMap.ContainsKey($c)) { $label = $script:JobTypeNumMap[$c] }
        elseif ("$TypeCode" -ne '' -and $TypeCode -isnot [DBNull]) { $label = 'Other backup job' }
        else { $label = '' }
    }
    if ($isCopy -and $label -and $label -notmatch '(?i)copy|tape') { $label = "$label (Copy)" }
    return $label
}

# ---------------------------------------------------------------------------
# 1. Connection helpers (identical model to the Restore report)
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
            foreach ($n in $names) { if ($obj.PSObject.Properties[$n] -and $obj.$n) { return $obj.$n } }
            return $null
        }
        foreach ($rp in $regPaths) {
            if (-not (Test-Path $rp)) { continue }
            $k = Get-ItemProperty -Path $rp -ErrorAction SilentlyContinue
            if (-not $k) { continue }
            if (-not $SqlServer) {
                $srv  = Get-Prop $k @('SqlServerName','DatabaseServer','SqlServer')
                $inst = Get-Prop $k @('SqlInstanceName','DatabaseInstance','SqlInstance')
                if ($srv) { if ($inst -and $inst -ne 'MSSQLSERVER') { $SqlServer = "$srv\$inst" } else { $SqlServer = $srv } }
            }
            if (-not $Database) { $Database = Get-Prop $k @('SqlDatabaseName','DatabaseName','SqlDatabase') }
            if ($SqlServer -and $Database) { break }
        }
    }
    if (-not $SqlServer) { $SqlServer = '.' }
    if (-not $Database)  { $Database  = 'VeeamONE' }
    [pscustomobject]@{ Server = $SqlServer; Database = $Database }
}

function Get-SqlConnectionString {
    param([string]$Server, [string]$Database, [pscredential]$Cred)
    $base = "Server=$Server;Database=$Database;Application Name=VeeamOneRmanBackupReport;Connect Timeout=30;"
    if ($Cred) {
        $u = $Cred.UserName; $p = $Cred.GetNetworkCredential().Password
        return "$base User ID=$u;Password=$p;"
    }
    return "$base Integrated Security=SSPI;"
}

function Enter-DomainImpersonation {
    param([pscredential]$Credential)
    if (-not $Credential) { return $null }
    if (-not ([System.Management.Automation.PSTypeName]'VeeamOneRman.NativeLogon').Type) {
        Add-Type -UsingNamespace 'Microsoft.Win32.SafeHandles' -Namespace 'VeeamOneRman' -Name 'NativeLogon' -MemberDefinition @'
[DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern bool LogonUser(string user, string domain, string password,
    int logonType, int logonProvider, out SafeAccessTokenHandle token);
'@
    }
    $user = $Credential.UserName; $domain = $null
    if     ($user -like '*\*') { $p = $user.Split('\', 2); $domain = $p[0]; $user = $p[1] }
    elseif ($user -like '*@*') { $domain = $null }
    $pw    = $Credential.GetNetworkCredential().Password
    $token = $null
    $ok = [VeeamOneRman.NativeLogon]::LogonUser($user, $domain, $pw, 9, 3, [ref]$token)
    if (-not $ok) {
        $code = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "LogonUser failed for '$($Credential.UserName)' (Win32 error $code). Check the domain\username and password."
    }
    return $token
}

function Invoke-VeeamSql {
    param([string]$ConnectionString, [string]$Query, [hashtable]$Params, [int]$TimeoutSec = 300)
    $ctx = $null
    if ($script:ImpToken) {
        $wi  = New-Object System.Security.Principal.WindowsIdentity($script:ImpToken.DangerousGetHandle())
        $ctx = $wi.Impersonate()
    }
    $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        # Read-only reporting: READ UNCOMMITTED so a SELECT never waits on the
        # Veeam ONE collector's writer locks (that is what made the report "hang").
        $cmd.CommandText    = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`r`n" + $Query
        $cmd.CommandTimeout = $TimeoutSec
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
# 2. Helpers
# ---------------------------------------------------------------------------
function ConvertTo-StatusText {
    param($ResultCode, $Message)
    if ($null -ne $ResultCode -and $ResultCode -isnot [DBNull]) {
        $n = 0
        if ([int]::TryParse([string]$ResultCode, [ref]$n) -and $script:ResultMap.ContainsKey($n)) {
            return $script:ResultMap[$n]
        }
        switch -Regex ([string]$ResultCode) {
            '^(?i)success'      { return 'Success' }
            '^(?i)warn'         { return 'Warning' }
            '^(?i)(fail|error)' { return 'Failed'  }
        }
    }
    if ($Message -and $Message -isnot [DBNull] -and [string]$Message -ne '') {
        switch -Regex ([string]$Message) {
            '(?i)(error|fail|unable|cannot)' { return 'Failed'  }
            '(?i)warn'                       { return 'Warning' }
            '(?i)(success|completed)'        { return 'Success' }
        }
    }
    return 'Unknown'
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

function Format-DurationSec {
    param($Seconds)
    $s = 0.0
    if (-not [double]::TryParse([string]$Seconds, [ref]$s) -or $s -lt 0) { return '' }
    $ts = [TimeSpan]::FromSeconds([int]$s)
    '{0:00}:{1:00}:{2:00}' -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds
}

function Test-InTimeWindow {
    param([datetime]$When, [string]$From, [string]$To)
    if (-not $From -and -not $To) { return $true }
    $t   = $When.TimeOfDay
    $f   = if ($From) { [TimeSpan]::Parse($From) } else { [TimeSpan]::Zero }
    $til = if ($To)   { [TimeSpan]::Parse($To)   } else { [TimeSpan]::FromDays(1) }
    if ($f -le $til) { return ($t -ge $f -and $t -le $til) }
    return ($t -ge $f -or $t -le $til)
}

function Get-Html { param([string]$s) if ($null -eq $s) { return '' } [System.Web.HttpUtility]::HtmlEncode($s) }

# Parse the protected host + database out of an RMAN job name, e.g.
#   BKP-PRDODB09-DMSPP-Full_RMAN            -> host PRDODB09, db DMSPP
#   BKP-PRDODB27_KAIACON_ARCHIVE_RMAN - ... -> host PRDODB27, db KAIACON
function Get-RmanDbInfo {
    param([string]$JobName)
    $h = ''; $db = ''
    $core = ($JobName -split '\s-\s')[0]         # drop " - fqdn" suffix if present
    if ($core -match '(?i)^\s*BKP-([A-Za-z0-9]+)[-_]([A-Za-z0-9]+)') { $h = $Matches[1]; $db = $Matches[2] }
    elseif ($core -match '(?i)^\s*BKP-([A-Za-z0-9]+)') { $h = $Matches[1] }
    if ($db -match '(?i)^(Full|Archive|Arch|Inc|Incr|Log)$') { $db = '' }   # that token was the mode, not a db
    # RMAN unmanaged registers the host as a "<host>-scan" cluster in Veeam ONE
    # (monitor.BpEpPluginCluster.access_name). Derive it from the host token.
    $scan = if ($h) { "$($h.ToLower())-scan" } else { '' }
    [pscustomobject]@{ DbHost = $h; Database = $db; ScanName = $scan }
}

function ConvertTo-Pdf {
    param([string]$HtmlPath, [string]$PdfPath)
    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )
    $exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $exe) { throw "No Microsoft Edge or Google Chrome found to render the PDF. Install Edge/Chrome, or use the CSV/HTML output." }
    $uri = ([System.Uri]$HtmlPath).AbsoluteUri
    if (Test-Path $PdfPath) { Remove-Item $PdfPath -Force }
    foreach ($headless in '--headless=new', '--headless') {
        $argLine = '{0} --disable-gpu --no-pdf-header-footer --print-to-pdf="{1}" "{2}"' -f $headless, $PdfPath, $uri
        Start-Process -FilePath $exe -ArgumentList $argLine -Wait -WindowStyle Hidden
        if (Test-Path $PdfPath) { return $PdfPath }
    }
    throw "PDF render command completed but produced no file (browser: $exe)."
}

function New-EmailBodyHtml {
    param([hashtable]$Stats, [string]$Range, [string]$Db, [string]$Scope, [hashtable]$Colors)
    $cell = "padding:6px 14px;border:1px solid #e3e8ee;font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    $num  = { param($v,$c) "<td style='$cell text-align:center;font-weight:700;color:$c;'>$v</td>" }
    @"
<div style='font-family:Segoe UI,Arial,sans-serif;color:#1a2733;'>
  <h2 style='margin:0 0 4px;'>VeeamONE - RMAN / Backup Success Rate Report</h2>
  <div style='color:#6b7885;font-size:13px;margin-bottom:14px;'>$Scope &middot; $Range &middot; DB: $Db</div>
  <table style='border-collapse:collapse;margin-bottom:12px;'>
    <tr>
      <td style='$cell color:#6b7885;'>Total</td>
      <td style='$cell color:#6b7885;'>Success</td>
      <td style='$cell color:#6b7885;'>Warning</td>
      <td style='$cell color:#6b7885;'>Failed</td>
      <td style='$cell color:#6b7885;'>Success rate</td>
    </tr>
    <tr>
      $(& $num $Stats.Total '#1a2733')
      $(& $num $Stats.Success $Colors.Success)
      $(& $num $Stats.Warning $Colors.Warning)
      $(& $num $Stats.Failed $Colors.Failed)
      $(& $num ("{0}%" -f $Stats.Rate) '#1a2733')
    </tr>
  </table>
  <div style='color:#6b7885;font-size:12px;'>The full report is attached (HTML/PDF/CSV).</div>
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
    if (-not $From) { $From = "VeeamONE-RMAN-Report@$($env:COMPUTERNAME)" }
    $params = @{
        To = $To; From = $From; Subject = $Subject; Body = $BodyHtml; BodyAsHtml = $true
        SmtpServer = $SmtpServer; Port = $SmtpPort; ErrorAction = 'Stop'
    }
    if ($UseSsl) { $params.UseSsl = $true }
    if ($Cred)   { $params.Credential = $Cred }
    $existing = @($Attachments | Where-Object { $_ -and (Test-Path $_) })
    if ($existing.Count -gt 0) { $params.Attachments = $existing }
    Send-MailMessage @params
}

# ---------------------------------------------------------------------------
# 3. Connect
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

if ($DemoData) {
    $connInfo = [pscustomobject]@{ Server = '(demo)'; Database = 'VeeamONE (demo data)' }
    Write-Host "DEMO MODE - generating synthetic RMAN backup report (no database queried)." -ForegroundColor Magenta
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
# 4. Discovery mode
# ---------------------------------------------------------------------------
if ($DiscoverObjects) {
    Write-Host "`n=== Candidate tables that may hold per-DATABASE / per-object detail ===" -ForegroundColor Yellow
    Write-Host "The 6 job/session tables do not carry database names. Look here for the" -ForegroundColor Gray
    Write-Host "table whose rows are individual databases/objects (with a session/backup link).`n" -ForegroundColor Gray
    try {
        Invoke-VeeamSql -ConnectionString $connStr -Query @'
SELECT  t.TABLE_SCHEMA + '.' + t.TABLE_NAME AS TableName,
        STUFF((SELECT ', ' + c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c
               WHERE c.TABLE_SCHEMA=t.TABLE_SCHEMA AND c.TABLE_NAME=t.TABLE_NAME
               ORDER BY c.ORDINAL_POSITION FOR XML PATH('')),1,2,'') AS Columns
FROM INFORMATION_SCHEMA.TABLES t
WHERE t.TABLE_NAME LIKE '%Object%' OR t.TABLE_NAME LIKE '%Task%'   OR t.TABLE_NAME LIKE '%File%'
   OR t.TABLE_NAME LIKE '%Database%' OR t.TABLE_NAME LIKE '%Item%' OR t.TABLE_NAME LIKE '%Point%'
   OR t.TABLE_NAME LIKE '%Oracle%'   OR t.TABLE_NAME LIKE '%Sql%'  OR t.TABLE_NAME LIKE '%Vm%'
ORDER BY TableName
'@ | Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Table scan failed: $($_.Exception.Message)" }

    Write-Host "=== Columns whose NAME suggests a database/object/instance ===" -ForegroundColor Yellow
    try {
        Invoke-VeeamSql -ConnectionString $connStr -Query @'
SELECT TABLE_SCHEMA + '.' + TABLE_NAME AS TableName, COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE COLUMN_NAME LIKE '%database%' OR COLUMN_NAME LIKE '%object_name%'
   OR COLUMN_NAME LIKE '%instance%' OR COLUMN_NAME LIKE '%db_name%'
   OR COLUMN_NAME LIKE '%item_name%' OR COLUMN_NAME LIKE '%obj_name%'
ORDER BY TableName, COLUMN_NAME
'@ | Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Column scan failed: $($_.Exception.Message)" }

    Write-Host "Object discovery complete. Send me the promising table (or export it) and I'll add a per-database drill-down." -ForegroundColor Green
    return
}

if ($Discover) {
    $S = $script:Schema
    Write-Host "`n=== Distinct '$($S.ResultCol)' values in $($S.Table) (last 90 days) ===" -ForegroundColor Yellow
    Write-Host "Confirm the result -> Success/Warning/Failed mapping in `$script:ResultMap.`n" -ForegroundColor Gray
    try {
        $sql = @"
SELECT s.[$($S.ResultCol)] AS ResultCode, COUNT(*) AS Sessions,
       SUM(CASE WHEN s.[$($S.ReasonCol)] IS NULL OR s.[$($S.ReasonCol)]='' THEN 0 ELSE 1 END) AS WithMessage,
       MAX(s.[$($S.ReasonCol)]) AS SampleMessage
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start
GROUP BY s.[$($S.ResultCol)] ORDER BY Sessions DESC
"@
        Invoke-VeeamSql -ConnectionString $connStr -Params @{ start = (Get-Date).AddDays(-90) } -Query $sql |
            Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Could not enumerate results: $($_.Exception.Message)" }

    Write-Host "=== RMAN job sessions (job_name LIKE '%RMAN%') by result (last 90 days) ===" -ForegroundColor Yellow
    try {
        $sql = @"
SELECT s.[$($S.ResultCol)] AS ResultCode, COUNT(*) AS Sessions, MAX(s.[$($S.JobNameCol)]) AS SampleJob
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start AND s.[$($S.JobNameCol)] LIKE '%RMAN%'
GROUP BY s.[$($S.ResultCol)] ORDER BY Sessions DESC
"@
        Invoke-VeeamSql -ConnectionString $connStr -Params @{ start = (Get-Date).AddDays(-90) } -Query $sql |
            Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Could not enumerate RMAN sessions: $($_.Exception.Message)" }
    Write-Host "Discovery complete." -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# 5. Build the record set
# ---------------------------------------------------------------------------
$records = New-Object System.Collections.Generic.List[object]
$scopeText = if ($IncludeAllBackups) { 'All backup jobs' } else { "Jobs matching '$JobNameLike'" }

if ($DemoData) {
    $jobs = 'BKP-PRDODB09-DMSPP-Full_RMAN','BKP-PRDODB27_KAIACON_ARCHIVE_RMAN','BKP-STGORADB01-STG19C03_Archive_RMAN',
            'BKP-PRDODB05_PRDCDB01_Full_RMAN','BKP-UATNGUDB1_NGU_Archive_RMAN'
    $rng = New-Object System.Random 20260812
    # Generate across the SELECTED [StartDate, EndDate] so the demo honors the picked range.
    $spanDays = [math]::Min(60, [math]::Max(1, [int][math]::Ceiling(($EndDate - $StartDate).TotalDays)))
    for ($d = 0; $d -lt $spanDays; $d++) {
        $day = $StartDate.Date.AddDays($d)
        foreach ($i in 0..($rng.Next(2,6))) {
            $stt = $day.AddHours($rng.Next(0,23)).AddMinutes($rng.Next(0,59))
            if ($stt -lt $StartDate -or $stt -ge $EndDate) { continue }
            $dur = $rng.Next(30, 5400)
            $roll = $rng.NextDouble()
            $status = if ($roll -gt 0.90) { 'Failed' } elseif ($roll -gt 0.80) { 'Warning' } else { 'Success' }
            $isFull = $rng.Next(0,2)
            $read = [long]($rng.Next(500, 60000)) * 1MB
            $xfer = if ($status -eq 'Failed') { 0 } else { [long]($read * ($rng.NextDouble()*0.3)) }
            $en = $stt.AddSeconds($dur)
            $jn = $jobs[$rng.Next(0,$jobs.Count)]
            $records.Add([pscustomobject]@{
                JobName     = $jn
                JobTypeName = Get-JobTypeName $jn 15
                Mode        = if ($isFull) { 'Full' } else { 'Incremental' }
                Status      = $status
                StartDate   = $stt.ToString('yyyy-MM-dd')
                StartTime   = $stt.ToString('HH:mm:ss')
                EndDate     = $en.ToString('yyyy-MM-dd')
                EndTime     = $en.ToString('HH:mm:ss')
                ElapsedTime = Format-DurationSec $dur
                DataRead    = Format-Size $read
                ReadBytesNum = [long]$read
                Transferred = Format-Size $xfer
                XferBytesNum = [long]$xfer
                Reason      = if ($status -eq 'Failed') { 'Task failed. Error: Failed to connect to the host (demo)' } elseif ($status -eq 'Warning') { 'Processing host (demo)' } else { '' }
            }) | Out-Null
        }
    }
    # $StartDate / $EndDate kept as picked so the demo reflects the selected range.
}
else {
    $S = $script:Schema
    $nameFilter = if (-not $IncludeAllBackups) { "AND s.[$($S.JobNameCol)] LIKE @jobname" } else { '' }
    $sql = @"
SELECT
    s.[$($S.IdCol)]       AS SessionId,
    s.[$($S.JobNameCol)]  AS JobName,
    s.[$($S.JobTypeCol)]  AS JobType,
    s.[$($S.IsFullCol)]   AS IsFull,
    s.[$($S.ResultCol)]   AS ResultCode,
    s.[$($S.StartCol)]    AS StartDate,
    s.[$($S.DurationCol)] AS DurationSec,
    s.[$($S.ReadCol)]     AS ReadBytes,
    s.[$($S.XferCol)]     AS XferBytes,
    s.[$($S.ReasonCol)]   AS Reason
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start
  AND s.[$($S.StartCol)] <  @end
  $nameFilter
ORDER BY s.[$($S.StartCol)] DESC
"@
    $endExclusive = if ($EndDate.TimeOfDay -eq [TimeSpan]::Zero) { $EndDate.Date.AddDays(1) } else { $EndDate }
    $sqlParams = @{ start = $StartDate; end = $endExclusive }
    if (-not $IncludeAllBackups) { $sqlParams['jobname'] = $JobNameLike }

    Write-Host "Querying backup sessions $($StartDate.ToString('yyyy-MM-dd')) .. $($EndDate.ToString('yyyy-MM-dd')) [$scopeText]..." -ForegroundColor Cyan
    try {
        $sessions = Invoke-VeeamSql -ConnectionString $connStr -Query $sql -Params $sqlParams
    }
    catch {
        Write-Warning "Backup-session query failed against $($S.Table): $($_.Exception.Message)"
        Write-Warning "Run -Discover to confirm the result codes on your build."
        return
    }

    foreach ($s in $sessions.Rows) {
        $start = [datetime]$s['StartDate']
        if (-not (Test-InTimeWindow -When $start -From $StartTime -To $EndTime)) { continue }
        $durSec = 0.0; [double]::TryParse([string]$s['DurationSec'], [ref]$durSec) | Out-Null
        $end = $start.AddSeconds([int]$durSec)
        $isFull = ($s['IsFull'] -isnot [DBNull]) -and ([string]$s['IsFull'] -eq '1' -or [string]$s['IsFull'] -eq 'True')
        $rb = 0.0; [double]::TryParse([string]$s['ReadBytes'], [ref]$rb) | Out-Null
        $xb = 0.0; [double]::TryParse([string]$s['XferBytes'], [ref]$xb) | Out-Null
        $records.Add([pscustomobject]@{
            JobName     = [string]$s['JobName']
            JobTypeName = Get-JobTypeName ([string]$s['JobName']) $s['JobType']
            Mode        = if ($isFull) { 'Full' } else { 'Incremental' }
            Status      = ConvertTo-StatusText $s['ResultCode'] $s['Reason']
            StartDate   = $start.ToString('yyyy-MM-dd')
            StartTime   = $start.ToString('HH:mm:ss')
            EndDate     = $end.ToString('yyyy-MM-dd')
            EndTime     = $end.ToString('HH:mm:ss')
            ElapsedTime = Format-DurationSec $durSec
            DataRead    = Format-Size $s['ReadBytes']
            ReadBytesNum = [long]$rb
            Transferred = Format-Size $s['XferBytes']
            XferBytesNum = [long]$xb
            Reason      = if ($s['Reason'] -isnot [DBNull]) { [string]$s['Reason'] } else { '' }
        }) | Out-Null
    }
}

Write-Host "Collected $($records.Count) backup session(s)." -ForegroundColor Green
if ($records.Count -eq 0) {
    Write-Warning "No backup sessions matched. Widen the date range, adjust -JobNameLike, or use -IncludeAllBackups / -Discover."
    return
}

# ---------------------------------------------------------------------------
# 5b. RMAN scan -> databases inventory (drill-down from the plugin tables)
#     monitor.BpEpPluginCluster (scan)  --ClusterToDbEntityLink-->  home entity
#     --> child BpEpPluginDbEntity rows (type 1 = database, type 2 = PDB).
#     Keyed by the scan access_name (e.g. "prdodb09-scan") so each RMAN job can
#     list the objects (databases/PDBs) it protects.
# ---------------------------------------------------------------------------
$scanDbMap = @{}   # scanNameShort -> List[string] of "Database" or "Database > PDB"
if ($DemoData) {
    $scanDbMap['prdodb09-scan']   = [System.Collections.Generic.List[string]]@('DMSPP')
    $scanDbMap['prdodb27-scan']   = [System.Collections.Generic.List[string]]@('KAIACON','ARCHCON')
    $scanDbMap['stgoradb01-scan'] = [System.Collections.Generic.List[string]]@('STG19C03')
    $scanDbMap['prdodb05-scan']   = [System.Collections.Generic.List[string]]@('PRDCDB01','PRDCDB01 > PDB1','PRDCDB01 > PDB2')
    $scanDbMap['uatngudb1-scan']  = [System.Collections.Generic.List[string]]@('NGU')
}
elseif ($SkipObjectDrilldown) {
    Write-Host "Object drill-down skipped (-SkipObjectDrilldown); objects fall back to the job-name database." -ForegroundColor DarkYellow
}
else {
    try {
        Write-Host "Reading plugin DB inventory (scan -> databases)..." -ForegroundColor Cyan
        $mapSql = @'
SELECT c.[access_name]     AS ScanName,
       db.[display_name]   AS DbName,
       pdb.[display_name]  AS PdbName
FROM [monitor].[BpEpPluginCluster] c
JOIN [monitor].[BpEpPluginClusterToDbEntityLink] l ON l.[cluster_id] = c.[uid]
JOIN [monitor].[BpEpPluginDbEntity] db  ON db.[parent_id] = l.[uid] AND db.[type] = 1 AND db.[delete_time] IS NULL
LEFT JOIN [monitor].[BpEpPluginDbEntity] pdb ON pdb.[parent_id] = db.[uid] AND pdb.[type] = 2 AND pdb.[delete_time] IS NULL
WHERE c.[delete_time] IS NULL
ORDER BY c.[access_name], db.[display_name], pdb.[display_name]
'@
        # Bounded timeout so a slow/locked plugin table can never hang the report.
        $mapRows = Invoke-VeeamSql -ConnectionString $connStr -Query $mapSql -TimeoutSec 90
        foreach ($row in $mapRows.Rows) {
            $sn = ([string]$row['ScanName']).ToLower()
            if (-not $sn) { continue }
            $snShort = ($sn -split '\.')[0]     # drop domain: preodb17-scan.dom -> preodb17-scan
            $db = [string]$row['DbName']
            $pdb = if ($row['PdbName'] -isnot [DBNull]) { [string]$row['PdbName'] } else { '' }
            $obj = if ($pdb) { "$db > $pdb" } else { $db }
            if (-not $scanDbMap.ContainsKey($snShort)) { $scanDbMap[$snShort] = New-Object System.Collections.Generic.List[string] }
            if ($obj -and -not $scanDbMap[$snShort].Contains($obj)) { $scanDbMap[$snShort].Add($obj) }
        }
        Write-Host "Plugin DB inventory: $($scanDbMap.Keys.Count) scan(s) mapped to databases." -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Could not read plugin DB inventory (BpEpPlugin* tables): $($_.Exception.Message)."
        Write-Warning "Objects fall back to the database parsed from the job name."
    }
}

# ---------------------------------------------------------------------------
# 6. Export CSV
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath = Join-Path $OutputFolder "VeeamONE-RMAN-Backup-Report-$stamp.csv"
$records |
    Select-Object JobName, JobTypeName, Mode, Status, StartDate, StartTime, EndDate, EndTime, ElapsedTime, DataRead, Transferred, Reason |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV  : $csvPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 7. Build the HTML report (self-contained SVG - no CDN/JS libs)
# ---------------------------------------------------------------------------
$byStatus = $records | Group-Object Status | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Count = $_.Count } }
$total = $records.Count
$succ  = ($byStatus | Where-Object Name -eq 'Success' | ForEach-Object Count) ; if (-not $succ) { $succ = 0 }
$warn  = ($byStatus | Where-Object Name -eq 'Warning' | ForEach-Object Count) ; if (-not $warn) { $warn = 0 }
$fail  = ($byStatus | Where-Object Name -eq 'Failed'  | ForEach-Object Count) ; if (-not $fail) { $fail = 0 }
$other = $total - $succ - $warn - $fail
$rate  = if ($total) { [math]::Round(($succ / $total) * 100, 1) } else { 0 }

$statusColors = @{ Success = '#00b336'; Warning = '#ffb300'; Failed = '#e5202e'; Unknown = '#8a8a8a' }

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
$donut = New-DonutSvg -Data @{ Success = $succ; Warning = $warn; Failed = $fail; Unknown = $other } -Colors $statusColors

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
        foreach ($st in 'Failed','Warning','Success') {
            $c = @($d.Group | Where-Object Status -eq $st).Count
            if ($c -le 0) { continue }
            $bh = $plot * $c / $max
            $y -= $bh
            $col = $Colors[$st]
            [void]$sb.Append(("<rect x='{0:N1}' y='{1:N1}' width='$bw' height='{2:N1}' fill='$col'/>" -f $x,$y,$bh))
        }
        $lbl = ($d.Name).Substring(5)
        [void]$sb.Append("<text x='{0:N1}' y='$($h-8)' text-anchor='middle' font-size='10' fill='var(--muted)'>$lbl</text>" -f ($x + $bw/2))
        $i++
    }
    [void]$sb.Append('</svg>')
    $sb.ToString()
}
$trend = New-TrendSvg -Records $records -Colors $statusColors

$rowsHtml = [System.Text.StringBuilder]::new()
foreach ($r in $records) {
    $sc = if ($statusColors.ContainsKey($r.Status)) { $statusColors[$r.Status] } else { '#8a8a8a' }
    [void]$rowsHtml.Append("<tr>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.JobName)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.JobTypeName)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Mode)</td>")
    [void]$rowsHtml.Append("<td><span class='badge' style='background:$sc'>$(Get-Html $r.Status)</span></td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.StartDate) $(Get-Html $r.StartTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.EndDate) $(Get-Html $r.EndTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.ElapsedTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.DataRead)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Transferred)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Reason)</td>")
    [void]$rowsHtml.Append("</tr>")
}

$legend = @"
<div class='legend'>
  <span><i style='background:$($statusColors.Success)'></i>Success ($succ)</span>
  <span><i style='background:$($statusColors.Warning)'></i>Warning ($warn)</span>
  <span><i style='background:$($statusColors.Failed)'></i>Failed ($fail)</span>
  $(if ($other -gt 0) { "<span><i style='background:$($statusColors.Unknown)'></i>Other ($other)</span>" })
</div>
"@

# --- per-database summary (each RMAN plugin job = one protected database) ---
# Columns: Job name | Object (database) | RMAN scan | Runs | Last status |
#          Last run | Total read (readable DB size) | Success rate
$dbGroups = @($records | Group-Object JobName)
$dbCount  = $dbGroups.Count
$hostCount = @($records | ForEach-Object { (Get-RmanDbInfo $_.JobName).DbHost } | Where-Object { $_ } | Sort-Object -Unique).Count
$dbNormal = 0
$totalReadBytes = 0.0
$totalStoredBytes = 0.0
$dbRowsHtml = [System.Text.StringBuilder]::new()
foreach ($g in ($dbGroups | Sort-Object Name)) {
    $latest = $g.Group | Sort-Object { "$($_.StartDate) $($_.StartTime)" } -Descending | Select-Object -First 1
    $info   = Get-RmanDbInfo $g.Name
    $runs   = $g.Count
    $ds     = @($g.Group | Where-Object Status -eq 'Success').Count
    $drate  = if ($runs) { [math]::Round($ds / $runs * 100, 1) } else { 0 }
    # Total readable DB size = sum of data read across this database's rescans;
    # Total stored = sum of data written to the repository (post-compression).
    $grpBytes = ($g.Group | Measure-Object -Property ReadBytesNum -Sum).Sum
    if (-not $grpBytes) { $grpBytes = 0 }
    $grpStored = ($g.Group | Measure-Object -Property XferBytesNum -Sum).Sum
    if (-not $grpStored) { $grpStored = 0 }
    $totalReadBytes += $grpBytes
    $totalStoredBytes += $grpStored
    if ($latest.Status -eq 'Success') { $dbNormal++ }
    $sc = if ($statusColors.ContainsKey($latest.Status)) { $statusColors[$latest.Status] } else { '#8a8a8a' }
    # Objects (databases) under this job: from the plugin inventory keyed by the
    # scan name; fall back to the DB parsed from the job name. @() keeps it an
    # array even when the map holds a single database (StrictMode-safe).
    $objs = if ($info.ScanName -and $scanDbMap.ContainsKey($info.ScanName)) { $scanDbMap[$info.ScanName] } else { $null }
    $objs = @($objs)   # force array (the if-assignment unrolls a single-item list to a scalar)
    $objectCell = if ($objs.Count -gt 0) {
        (($objs | ForEach-Object { Get-Html $_ }) -join '<br>')
    } elseif ($info.Database) { Get-Html $info.Database } else { Get-Html $g.Name }
    [void]$dbRowsHtml.Append("<tr>")
    [void]$dbRowsHtml.Append("<td>$(Get-Html $g.Name)</td>")
    [void]$dbRowsHtml.Append("<td>$objectCell</td>")
    [void]$dbRowsHtml.Append("<td>$(Get-Html $info.ScanName)</td>")
    [void]$dbRowsHtml.Append("<td>$runs</td>")
    [void]$dbRowsHtml.Append("<td><span class='badge' style='background:$sc'>$(Get-Html $latest.Status)</span></td>")
    [void]$dbRowsHtml.Append("<td>$(Get-Html $latest.StartDate) $(Get-Html $latest.StartTime)</td>")
    [void]$dbRowsHtml.Append("<td>$(Format-Size $grpBytes)</td>")
    [void]$dbRowsHtml.Append("<td>$(Format-Size $grpStored)</td>")
    [void]$dbRowsHtml.Append("<td>$drate%</td>")
    [void]$dbRowsHtml.Append("</tr>")
}
$dbIssues = $dbCount - $dbNormal
$totalReadText = Format-Size $totalReadBytes
$totalStoredText = Format-Size $totalStoredBytes

$rangeText = "$($StartDate.ToString('yyyy-MM-dd'))  to  $($EndDate.ToString('yyyy-MM-dd'))"
if ($StartTime -or $EndTime) { $rangeText += "   (window $([string]$StartTime)-$([string]$EndTime))" }

$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>VeeamONE RMAN Backup Report</title>
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
      <h1>VeeamONE - RMAN / Backup Success Rate Report</h1>
      <div class='sub'>$scopeText &middot; $rangeText &middot; DB: $(Get-Html $connInfo.Database)</div>
    </div>
    <div class='sub'>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
  </header>

  <div class='cards'>
    <div class='card'><div class='n'>$total</div><div class='l'>Total runs</div></div>
    <div class='card'><div class='n'>$dbCount</div><div class='l'>Databases backed up</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Success)'>$dbNormal</div><div class='l'>Running normally</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Failed)'>$dbIssues</div><div class='l'>Databases w/ issues</div></div>
    <div class='card'><div class='n'>$hostCount</div><div class='l'>Database servers</div></div>
    <div class='card'><div class='n'>$totalReadText</div><div class='l'>Total DB read</div></div>
    <div class='card'><div class='n'>$totalStoredText</div><div class='l'>Total stored (repo)</div></div>
    <div class='card'><div class='n'>$rate%</div><div class='l'>Run success rate</div></div>
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

  <div class='panel' style='margin-bottom:22px;'>
    <h2>Databases protected ($dbCount) &middot; running normally $dbNormal &middot; needs attention $dbIssues</h2>
    <div class='tablewrap' style='border:0;'>
      <table>
        <thead><tr>
          <th>Job name</th><th>Object (database)</th><th>RMAN scan</th><th>Runs</th><th>Last status</th><th>Last run</th><th>Total DB read</th><th>Total stored</th><th>Success rate</th>
        </tr></thead>
        <tbody>
          $($dbRowsHtml.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <div class='panel' style='margin-bottom:0;'>
    <h2>RMAN scans - individual runs ($total) &middot; total DB read $totalReadText</h2>
    <div class='tablewrap' style='border:0;'>
      <table>
        <thead><tr>
          <th>Job name (scan)</th><th>Type</th><th>Mode</th><th>Status</th>
          <th>Start</th><th>End</th><th>Elapsed</th><th>DB read</th><th>Stored (repo)</th><th>Reason</th>
        </tr></thead>
        <tbody>
          $($rowsHtml.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <footer>VeeamONE RMAN / Backup Report &middot; source data from the Veeam ONE database (read-only)</footer>
</div>
</body>
</html>
"@

$htmlPath = Join-Path $OutputFolder "VeeamONE-RMAN-Backup-Report-$stamp.html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "HTML : $htmlPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 8. Optional PDF
# ---------------------------------------------------------------------------
$pdfPath = $null
if ($Pdf) {
    $pdfPath = Join-Path $OutputFolder "VeeamONE-RMAN-Backup-Report-$stamp.pdf"
    try { ConvertTo-Pdf -HtmlPath $htmlPath -PdfPath $pdfPath | Out-Null; Write-Host "PDF  : $pdfPath" -ForegroundColor Green }
    catch { Write-Warning "PDF generation failed: $($_.Exception.Message)"; $pdfPath = $null }
}

# ---------------------------------------------------------------------------
# 9. Optional email delivery
# ---------------------------------------------------------------------------
if ($EmailTo -and $EmailTo.Count -gt 0) {
    $subject = if ($EmailSubject) { $EmailSubject } else {
        "VeeamONE RMAN/Backup Report - $rate% success ($total runs) - $($StartDate.ToString('yyyy-MM-dd'))..$($EndDate.ToString('yyyy-MM-dd'))"
    }
    $bodyHtml = New-EmailBodyHtml `
        -Stats @{ Total = $total; Success = $succ; Warning = $warn; Failed = $fail; Rate = $rate } `
        -Range $rangeText -Db ([string]$connInfo.Database) -Scope $scopeText -Colors $statusColors
    $attachments = @($csvPath, $htmlPath, $pdfPath) | Where-Object { $_ }
    try {
        Send-ReportEmail -To $EmailTo -From $EmailFrom -SmtpServer $SmtpServer -SmtpPort $SmtpPort `
            -UseSsl:$SmtpUseSsl -Cred $SmtpCredential -Subject $subject -BodyHtml $bodyHtml -Attachments $attachments
        Write-Host "Email: sent to $($EmailTo -join ', ') via $SmtpServer`:$SmtpPort" -ForegroundColor Green
    }
    catch { Write-Warning "Email delivery failed: $($_.Exception.Message)" }
}

Write-Host "`nSummary: $total runs | $succ success | $warn warning | $fail failed | $rate% success rate" -ForegroundColor Cyan
[pscustomobject]@{ Csv = $csvPath; Html = $htmlPath; Pdf = $pdfPath; Total = $total; Success = $succ; Warning = $warn; Failed = $fail; SuccessRate = $rate }
