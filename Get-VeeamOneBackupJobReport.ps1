<#
.SYNOPSIS
    VeeamONE Backup Jobs report - an inventory of every backup job with its
    schedule, last start / end, next run, last result (status) and the backup
    policy it belongs to.

.DESCRIPTION
    Runs on (or next to) the Veeam ONE server. Reads the Veeam ONE database
    connection from the registry automatically (override with -SqlServer /
    -Database) and queries the backup job configuration Veeam ONE collected from
    the monitored Veeam Backup & Replication servers.

    Data sources (Veeam ONE monitor schema):
      - [monitor].[BpJob]                    - job config + schedule + last result
      - [monitor].[BpJobLastFinishedResult]  - last start / end + failure message
      - [monitor].[BpBackup]                 - backup policy name (via policy tag)

    Output per job: Backup policy, Job name, Type, Schedule, Enabled, Last start,
    Last end, Next run, Status (last result) and message - plus a donut of the
    last-result status, a "jobs per backup policy" chart, and summary tiles in
    the HTML report.

.PARAMETER SqlServer / Database / SqlCredential / WindowsCredential
    Connection to the Veeam ONE database (same model as the other reports).

.PARAMETER NameLike
    Optional SQL LIKE filter on the job name (e.g. "%RMAN%", "%SQL%").

.PARAMETER PolicyLike
    Optional SQL LIKE filter on the backup policy name.

.PARAMETER EnabledOnly
    Only include jobs whose schedule is enabled.

.PARAMETER OutputFolder / Pdf / Period
    Output location, optional PDF render, and relative window (for scheduling).
    Note: this is an inventory of current job state; -StartDate/-EndDate/-Period
    are accepted for a uniform scheduling surface and only label the report.

.PARAMETER EmailTo / EmailFrom / SmtpServer / SmtpPort / SmtpUseSsl / SmtpCredential / EmailSubject
    Optional e-mail delivery (attachments: CSV + HTML + PDF).

.PARAMETER Discover
    List the distinct job type codes present (with sample names) and exit.

.PARAMETER DemoData
    Build the report from synthetic data (no DB) to preview the format.

.EXAMPLE
    .\Get-VeeamOneBackupJobReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)

.EXAMPLE
    # Only RMAN jobs, e-mailed as a PDF
    .\Get-VeeamOneBackupJobReport.ps1 -NameLike '%RMAN%' -Pdf `
        -EmailTo ops@contoso.com -SmtpServer smtp.contoso.com

.EXAMPLE
    .\Get-VeeamOneBackupJobReport.ps1 -DemoData
#>
[CmdletBinding(DefaultParameterSetName = 'Report')]
param(
    [string]   $SqlServer,
    [string]   $Database,
    [pscredential] $SqlCredential,
    [pscredential] $WindowsCredential,

    [datetime] $StartDate = (Get-Date).AddDays(-30).Date,
    [datetime] $EndDate   = (Get-Date),

    [Parameter(ParameterSetName = 'Report')]
    [string]   $NameLike,
    [Parameter(ParameterSetName = 'Report')]
    [string]   $PolicyLike,
    [Parameter(ParameterSetName = 'Report')]
    [switch]   $EnabledOnly,

    [string]   $OutputFolder = (Join-Path $PSScriptRoot 'output'),

    [switch]   $Pdf,

    # Also emit an A5 "backup calendar" (print-friendly timetable of configured
    # schedules: schedule type, time to run, next run, last start/end).
    [switch]   $Calendar,

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

    [switch]   $DemoData
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:ImpToken = $null

if ($EmailTo) {
    $EmailTo = @($EmailTo | ForEach-Object { $_ -split '[;,]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# ---------------------------------------------------------------------------
# 0. Version-sensitive mapping (adjust after -Discover if your build differs).
#    last_finished_result: 2 = Success, 3 = Warning, 4 = Failed (confirmed build).
# ---------------------------------------------------------------------------
$script:ResultMap = @{
    2 = 'Success'
    3 = 'Warning'
    4 = 'Failed'
}
# Friendly Veeam job-type names. Numeric codes derived from the Veeam ONE data
# (type + source_type + platform + name pattern). Get-JobTypeName reads the job
# NAME first (most reliable), then falls back to this numeric map.
$script:JobTypeNumMap = @{
    0='File / Object Storage Backup'; 1='VM Backup (VMware/virtual)'; 2='Replication'; 5='Backup to Tape'
    6='Backup to Tape'; 9='Replication (DR)'; 10='Microsoft SQL Log Backup (plugin)'
    13='Oracle RMAN Plugin - archived logs (unmanaged)'; 15='Oracle RMAN Plugin - full/incremental (unmanaged)'
    16='Backup Copy'; 17='Backup Copy'; 18='Veeam Agent Backup'; 21='NAS / File Share Backup'
    22='Nutanix AHV Backup'; 23='NAS Backup Copy'; 24='Backup Copy'; 25='Nutanix AHV Backup'
    31='Application DB Plugin Backup'; 32='SAP HANA Plugin (backint)'; 33='Oracle Plugin Backup'
    34='Oracle Plugin - redo/archived logs'; 35='Oracle Backup Copy'; 36='Oracle Backup Copy'; 37='Oracle Backup Copy'
}

# ---------------------------------------------------------------------------
# 1. Connection helpers (identical model to the other reports)
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
    $base = "Server=$Server;Database=$Database;Application Name=VeeamOneBackupJobReport;Connect Timeout=30;"
    if ($Cred) {
        $u = $Cred.UserName; $p = $Cred.GetNetworkCredential().Password
        return "$base User ID=$u;Password=$p;"
    }
    return "$base Integrated Security=SSPI;"
}

function Enter-DomainImpersonation {
    param([pscredential]$Credential)
    if (-not $Credential) { return $null }
    if (-not ([System.Management.Automation.PSTypeName]'VeeamOneJob.NativeLogon').Type) {
        Add-Type -UsingNamespace 'Microsoft.Win32.SafeHandles' -Namespace 'VeeamOneJob' -Name 'NativeLogon' -MemberDefinition @'
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
    $ok = [VeeamOneJob.NativeLogon]::LogonUser($user, $domain, $pw, 9, 3, [ref]$token)
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
        # Read-only reporting: READ UNCOMMITTED so a SELECT never waits on the
        # Veeam ONE collector's writer locks (prevents the report from hanging).
        $cmd.CommandText    = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`r`n" + $Query
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
# 2. Helpers
# ---------------------------------------------------------------------------
function ConvertTo-JobStatus {
    param($ResultCode)
    if ($null -eq $ResultCode -or $ResultCode -is [DBNull] -or [string]$ResultCode -eq '') { return 'Never run' }
    $n = 0
    if ([int]::TryParse([string]$ResultCode, [ref]$n) -and $script:ResultMap.ContainsKey($n)) { return $script:ResultMap[$n] }
    switch -Regex ([string]$ResultCode) {
        '^(?i)success'      { return 'Success' }
        '^(?i)warn'         { return 'Warning' }
        '^(?i)(fail|error)' { return 'Failed'  }
    }
    return 'Unknown'
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

function Format-Retention {
    param($RetainDays, $RetentionPolicy)
    $d = 0; $p = 0
    if ($RetainDays -isnot [DBNull] -and [int]::TryParse([string]$RetainDays, [ref]$d) -and $d -gt 0) { return "$d days" }
    if ($RetentionPolicy -isnot [DBNull] -and [int]::TryParse([string]$RetentionPolicy, [ref]$p) -and $p -gt 0) { return "$p restore points" }
    return ''
}

function Format-Dt {
    param($Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return '' }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm') }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [ref]$dt)) { return $dt.ToString('yyyy-MM-dd HH:mm') }
    return ''
}

function Format-Schedule {
    param($ScheduleType, $SchedTime, $Enabled)
    $st = if ($ScheduleType -is [DBNull] -or [string]$ScheduleType -eq '') { $null } else { [string]$ScheduleType }
    if (-not $st) { return 'Triggered / rescan' }   # RMAN + agent-managed jobs have no VBR schedule
    # SchedTime is a datetime whose TIME-OF-DAY is the scheduled run time.
    $tod = ''
    if ($SchedTime -isnot [DBNull]) {
        $dt = [datetime]::MinValue
        if ($SchedTime -is [datetime]) { $tod = ([datetime]$SchedTime).ToString('HH:mm') }
        elseif ([datetime]::TryParse([string]$SchedTime, [ref]$dt)) { $tod = $dt.ToString('HH:mm') }
    }
    if ($tod) { "$st @ $tod" } else { $st }
}

function Get-TimeOfDay {
    param($SchedTime)
    if ($SchedTime -is [DBNull]) { return '' }
    if ($SchedTime -is [datetime]) { return ([datetime]$SchedTime).ToString('HH:mm') }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse([string]$SchedTime, [ref]$dt)) { return $dt.ToString('HH:mm') }
    return ''
}

function Get-Html { param([string]$s) if ($null -eq $s) { return '' } [System.Web.HttpUtility]::HtmlEncode($s) }

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
    param([hashtable]$Stats, [string]$Db, [hashtable]$Colors)
    $cell = "padding:6px 14px;border:1px solid #e3e8ee;font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    $num  = { param($v,$c) "<td style='$cell text-align:center;font-weight:700;color:$c;'>$v</td>" }
    @"
<div style='font-family:Segoe UI,Arial,sans-serif;color:#1a2733;'>
  <h2 style='margin:0 0 4px;'>VeeamONE - Backup Jobs Report</h2>
  <div style='color:#6b7885;font-size:13px;margin-bottom:14px;'>Job inventory &middot; DB: $Db</div>
  <table style='border-collapse:collapse;margin-bottom:12px;'>
    <tr>
      <td style='$cell color:#6b7885;'>Jobs</td>
      <td style='$cell color:#6b7885;'>Success</td>
      <td style='$cell color:#6b7885;'>Warning</td>
      <td style='$cell color:#6b7885;'>Failed</td>
      <td style='$cell color:#6b7885;'>Disabled</td>
    </tr>
    <tr>
      $(& $num $Stats.Total '#1a2733')
      $(& $num $Stats.Success $Colors.Success)
      $(& $num $Stats.Warning $Colors.Warning)
      $(& $num $Stats.Failed $Colors.Failed)
      $(& $num $Stats.Disabled '#6b7885')
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
    if (-not $From) { $From = "VeeamONE-BackupJob-Report@$($env:COMPUTERNAME)" }
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
    Write-Host "DEMO MODE - generating synthetic backup jobs report (no database queried)." -ForegroundColor Magenta
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
if ($Discover) {
    Write-Host "`n=== Distinct BpJob.[type] values (with sample job names) ===" -ForegroundColor Yellow
    Write-Host "Add friendly labels to `$script:JobTypeMap for the ones you care about.`n" -ForegroundColor Gray
    try {
        Invoke-VeeamSql -ConnectionString $connStr -Query @'
SELECT j.[type] AS JobType, COUNT(*) AS Jobs, MAX(j.[name]) AS SampleName
FROM [monitor].[BpJob] j
GROUP BY j.[type] ORDER BY Jobs DESC
'@ | Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Could not enumerate job types: $($_.Exception.Message)" }

    Write-Host "=== Distinct last_finished_result values ===" -ForegroundColor Yellow
    try {
        Invoke-VeeamSql -ConnectionString $connStr -Query @'
SELECT j.[last_finished_result] AS ResultCode, COUNT(*) AS Jobs
FROM [monitor].[BpJob] j
GROUP BY j.[last_finished_result] ORDER BY Jobs DESC
'@ | Format-Table -AutoSize | Out-String | Write-Host
    } catch { Write-Warning "Could not enumerate results: $($_.Exception.Message)" }
    Write-Host "Discovery complete." -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# 5. Build the record set
# ---------------------------------------------------------------------------
$records = New-Object System.Collections.Generic.List[object]

if ($DemoData) {
    $policies = @(
        @{ p='BKP-PDC1_PRD_Linux_CLS01';   n=6 },
        @{ p='BKP-PDC1_PRD_Windows_CLS01'; n=5 },
        @{ p='BKP-NUTANIX-PDC1-TST_CLS01'; n=4 },
        @{ p='RMAN Standalone (unmanaged)'; n=8 }
    )
    $rng = New-Object System.Random 20260812
    foreach ($pol in $policies) {
        for ($i = 1; $i -le $pol.n; $i++) {
            $isRman = $pol.p -like '*RMAN*'
            $name = if ($isRman) { "BKP-PRDODB{0:00}-DB{1}_RMAN" -f $rng.Next(1,99), $i } else { "$($pol.p) - HOST{0:00}" -f $i }
            $roll = $rng.NextDouble()
            $status = if ($roll -gt 0.88) { 'Failed' } elseif ($roll -gt 0.76) { 'Warning' } elseif ($roll -gt 0.70) { 'Never run' } else { 'Success' }
            $enabled = $rng.NextDouble() -gt 0.15
            $lastStart = (Get-Date).AddHours(-$rng.Next(1, 240))
            $tcode = if ($isRman) { 15 } else { 18 }
            $rhr = $rng.Next(0,23)
            $records.Add([pscustomobject]@{
                Policy    = $pol.p
                JobName   = $name
                JobType   = Get-JobTypeName $name $tcode
                Schedule  = if ($isRman) { 'Triggered / rescan' } else { 'Daily @ {0:00}:00' -f $rhr }
                SchedType = if ($isRman) { 'Triggered / rescan' } else { 'Daily' }
                RunTime   = if ($isRman) { '' } else { '{0:00}:00' -f $rhr }
                Retention = if ($isRman) { '14 days' } else { '{0} restore points' -f $rng.Next(7,31) }
                Enabled   = if ($enabled) { 'Enabled' } else { 'Disabled' }
                LastStart = if ($status -eq 'Never run') { '' } else { $lastStart.ToString('yyyy-MM-dd HH:mm') }
                LastEnd   = if ($status -eq 'Never run') { '' } else { $lastStart.AddMinutes($rng.Next(2,90)).ToString('yyyy-MM-dd HH:mm') }
                NextRun   = if ($enabled -and -not $isRman) { (Get-Date).AddHours($rng.Next(1,24)).ToString('yyyy-MM-dd HH:mm') } else { '' }
                Status    = $status
                Message   = if ($status -eq 'Failed') { 'Task failed. Error: sample (demo)' } elseif ($status -eq 'Warning') { 'Completed with warnings (demo)' } else { '' }
            }) | Out-Null
        }
    }
}
else {
    $filters = New-Object System.Collections.Generic.List[string]
    $sqlParams = @{}
    if ($NameLike)    { $filters.Add("j.[name] LIKE @namelike");   $sqlParams['namelike']   = $NameLike }
    if ($PolicyLike)  { $filters.Add("bp.backup_policy_name LIKE @pollike"); $sqlParams['pollike'] = $PolicyLike }
    if ($EnabledOnly) { $filters.Add("j.[schedule_enabled] = 1") }
    $where = if ($filters.Count -gt 0) { 'WHERE ' + ($filters -join ' AND ') } else { '' }

    $sql = @"
SELECT
    j.[name]                 AS JobName,
    j.[type]                 AS JobTypeCode,
    j.[schedule_type]        AS ScheduleType,
    j.[schedule_enabled]     AS ScheduleEnabled,
    j.[time]                 AS SchedTime,
    j.[next_run_time]        AS NextRun,
    j.[last_end_time]        AS LastEnd,
    j.[last_finished_result] AS ResultCode,
    j.[description]          AS Description,
    j.[retention_policy]     AS RetentionPolicy,
    j.[retain_days]          AS RetainDays,
    lfr.[last_start_time]    AS LastStart,
    lfr.[last_end_time]      AS LfrEnd,
    lfr.[failure_message]    AS Message,
    bp.[backup_policy_name]  AS PolicyName
FROM [monitor].[BpJob] j
LEFT JOIN [monitor].[BpJobLastFinishedResult] lfr ON lfr.[uid] = j.[uid]
OUTER APPLY (
    SELECT TOP 1 b.[backup_policy_name]
    FROM [monitor].[BpBackup] b
    WHERE b.[backup_policy_tag] = j.[backup_policy_tag]
      AND b.[backup_policy_name] IS NOT NULL AND b.[backup_policy_name] <> ''
) bp
$where
ORDER BY bp.[backup_policy_name], j.[name]
"@

    Write-Host "Querying backup jobs$(if ($NameLike) { " name~'$NameLike'" })$(if ($PolicyLike) { " policy~'$PolicyLike'" })$(if ($EnabledOnly) { ' (enabled only)' })..." -ForegroundColor Cyan
    try {
        $jobs = Invoke-VeeamSql -ConnectionString $connStr -Query $sql -Params $sqlParams
    }
    catch {
        Write-Warning "Backup-job query failed: $($_.Exception.Message)"
        Write-Warning "Run -Discover to confirm the codes on your build."
        return
    }

    foreach ($j in $jobs.Rows) {
        $policy = if ($j['PolicyName'] -isnot [DBNull] -and [string]$j['PolicyName'] -ne '') { [string]$j['PolicyName'] } else { '(no policy)' }
        # Prefer the last-finished-result table's start; fall back to job's last_end_time.
        $lastStart = Format-Dt $j['LastStart']
        $lastEnd   = if ($j['LfrEnd'] -isnot [DBNull]) { Format-Dt $j['LfrEnd'] } else { Format-Dt $j['LastEnd'] }
        $records.Add([pscustomobject]@{
            Policy    = $policy
            JobName   = [string]$j['JobName']
            JobType   = Get-JobTypeName ([string]$j['JobName']) $j['JobTypeCode']
            Schedule  = Format-Schedule $j['ScheduleType'] $j['SchedTime'] $j['ScheduleEnabled']
            SchedType = if ($j['ScheduleType'] -isnot [DBNull] -and [string]$j['ScheduleType'] -ne '') { [string]$j['ScheduleType'] } else { 'Triggered / rescan' }
            RunTime   = Get-TimeOfDay $j['SchedTime']
            Retention = Format-Retention $j['RetainDays'] $j['RetentionPolicy']
            Enabled   = if ([string]$j['ScheduleEnabled'] -eq '1' -or [string]$j['ScheduleEnabled'] -eq 'True') { 'Enabled' } else { 'Disabled' }
            LastStart = $lastStart
            LastEnd   = $lastEnd
            NextRun   = Format-Dt $j['NextRun']
            Status    = ConvertTo-JobStatus $j['ResultCode']
            Message   = if ($j['Message'] -isnot [DBNull]) { [string]$j['Message'] } else { '' }
        }) | Out-Null
    }
}

Write-Host "Collected $($records.Count) backup job(s)." -ForegroundColor Green
if ($records.Count -eq 0) {
    Write-Warning "No backup jobs matched. Adjust -NameLike / -PolicyLike, drop -EnabledOnly, or run -Discover."
    return
}

# Refine plain agent jobs: SQL server vs Oracle/RMAN host (script-driven).
# VeeamONE does not store the job's script, so an agent job whose host also has
# an RMAN plugin job is inferred as the script-driven RMAN case; a SQL-named host
# is inferred as a SQL server.
$rmanHostSet = @{}
foreach ($r in $records) {
    if ($r.JobName -match '(?i)RMAN' -and $r.JobName -match '(?i)BKP-([A-Za-z0-9]+)') { $rmanHostSet[$Matches[1].ToUpper()] = $true }
}
foreach ($r in $records) {
    if ($r.JobType -eq 'Veeam Agent Backup') {
        $tok = (($r.JobName -replace '(?i)\s*Backup\s*$','') -split '[\s\\]')[0]
        $tok = ($tok -replace '\..*$','').ToUpper()
        if     ($tok -and $rmanHostSet.ContainsKey($tok)) { $r.JobType = 'Veeam Agent - RMAN host (script)' }
        elseif ($r.JobName -match '(?i)SQL')              { $r.JobType = 'Veeam Agent - SQL server' }
    }
}

# "Backup policy" only applies to policy-managed jobs. Standalone / unmanaged jobs
# (RMAN, plugin) have none, so hide the policy column/tile/chart when NO job has one.
$realPolicies = @($records | Where-Object { $_.Policy -and $_.Policy -ne '(no policy)' } | Select-Object -ExpandProperty Policy -Unique)
$hasPolicy   = $realPolicies.Count -gt 0
$policyCount = $realPolicies.Count
$policyText  = if ($hasPolicy) { "across $policyCount policies" } else { 'standalone / unmanaged (no backup policy)' }

# ---------------------------------------------------------------------------
# 6. Export CSV
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath = Join-Path $OutputFolder "VeeamONE-BackupJobs-Report-$stamp.csv"
$csvCols = if ($hasPolicy) { 'Policy','JobName','JobType','Schedule','Retention','Enabled','LastStart','LastEnd','NextRun','Status','Message' }
           else            { 'JobName','JobType','Schedule','Retention','Enabled','LastStart','LastEnd','NextRun','Status','Message' }
$records | Select-Object $csvCols | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV  : $csvPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 7. Build the HTML report
# ---------------------------------------------------------------------------
$total    = $records.Count
$succ     = @($records | Where-Object Status -eq 'Success').Count
$warn     = @($records | Where-Object Status -eq 'Warning').Count
$fail     = @($records | Where-Object Status -eq 'Failed').Count
$never    = @($records | Where-Object Status -eq 'Never run').Count
$otherN   = $total - $succ - $warn - $fail - $never
$disabled = @($records | Where-Object Enabled -eq 'Disabled').Count
# $hasPolicy / $policyCount computed above (before CSV export).

$statusColors = @{ Success = '#00b336'; Warning = '#ffb300'; Failed = '#e5202e'; 'Never run' = '#8a63d2'; Unknown = '#8a8a8a' }

function New-DonutSvg {
    param([hashtable]$Data, [hashtable]$Colors, [string]$CenterTop, [string]$CenterBot, [int]$Size = 220)
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
    [void]$sb.Append("<text x='$cx' y='$($cy-2)' text-anchor='middle' font-size='28' font-weight='700' fill='var(--fg)'>$CenterTop</text>")
    [void]$sb.Append("<text x='$cx' y='$($cy+18)' text-anchor='middle' font-size='12' fill='var(--muted)'>$CenterBot</text>")
    [void]$sb.Append('</svg>')
    $sb.ToString()
}
$donut = New-DonutSvg -Data @{ Success = $succ; Warning = $warn; Failed = $fail; 'Never run' = $never; Unknown = $otherN } `
    -Colors $statusColors -CenterTop "$total" -CenterBot 'jobs'

# --- horizontal "jobs per policy" bar (top 12), stacked by last-result status ---
function New-PolicyBarSvg {
    param($Records, [hashtable]$Colors, [int]$Top = 12, [string]$Field = 'Policy')
    $groups = $Records | Group-Object $Field | Sort-Object Count -Descending | Select-Object -First $Top
    if (@($groups).Count -eq 0) { return '' }
    $rowH = 26; $gap = 8; $labelW = 240; $barMax = 460
    $h = (@($groups).Count) * ($rowH + $gap) + 10
    $w = $labelW + $barMax + 60
    $max = ($groups | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    if (-not $max) { $max = 1 }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg viewBox='0 0 $w $h' width='100%' height='$h' preserveAspectRatio='xMinYMin meet'>")
    $y = 6
    foreach ($g in $groups) {
        $name = $g.Name; if ($name.Length -gt 40) { $name = $name.Substring(0,38) + '..' }
        [void]$sb.Append("<text x='0' y='$($y+17)' font-size='12' fill='var(--fg)'>$([System.Web.HttpUtility]::HtmlEncode($name))</text>")
        $x = $labelW
        foreach ($st in 'Success','Warning','Failed','Never run','Unknown') {
            $c = @($g.Group | Where-Object Status -eq $st).Count
            if ($c -le 0) { continue }
            $bw = $barMax * $c / $max
            $col = if ($Colors.ContainsKey($st)) { $Colors[$st] } else { '#8a8a8a' }
            [void]$sb.Append(("<rect x='{0:N1}' y='$y' width='{1:N1}' height='$rowH' fill='$col'/>" -f $x,$bw))
            $x += $bw
        }
        [void]$sb.Append("<text x='$($x+6)' y='$($y+17)' font-size='12' fill='var(--muted)'>$($g.Count)</text>")
        $y += $rowH + $gap
    }
    [void]$sb.Append('</svg>')
    $sb.ToString()
}
# Group by policy when policies exist, otherwise by job type (more useful for
# standalone / unmanaged environments where there is no backup policy).
$barField = if ($hasPolicy) { 'Policy' } else { 'JobType' }
$barTitle = if ($hasPolicy) { 'Jobs per backup policy (top 12)' } else { 'Jobs by type (top 12)' }
$policyBar = New-PolicyBarSvg -Records $records -Colors $statusColors -Field $barField

$legend = @"
<div class='legend'>
  <span><i style='background:$($statusColors.Success)'></i>Success ($succ)</span>
  <span><i style='background:$($statusColors.Warning)'></i>Warning ($warn)</span>
  <span><i style='background:$($statusColors.Failed)'></i>Failed ($fail)</span>
  $(if ($never -gt 0) { "<span><i style='background:$($statusColors.'Never run')'></i>Never run ($never)</span>" })
  $(if ($otherN -gt 0) { "<span><i style='background:$($statusColors.Unknown)'></i>Other ($otherN)</span>" })
</div>
"@

$rowsHtml = [System.Text.StringBuilder]::new()
foreach ($r in $records) {
    $sc = if ($statusColors.ContainsKey($r.Status)) { $statusColors[$r.Status] } else { '#8a8a8a' }
    $en = if ($r.Enabled -eq 'Disabled') { "<span style='color:var(--muted)'>Disabled</span>" } else { 'Enabled' }
    [void]$rowsHtml.Append("<tr>")
    if ($hasPolicy) { [void]$rowsHtml.Append("<td>$(Get-Html $r.Policy)</td>") }
    [void]$rowsHtml.Append("<td>$(Get-Html $r.JobName)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.JobType)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Schedule)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Retention)</td>")
    [void]$rowsHtml.Append("<td>$en</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.LastStart)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.LastEnd)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.NextRun)</td>")
    [void]$rowsHtml.Append("<td><span class='badge' style='background:$sc'>$(Get-Html $r.Status)</span></td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Message)</td>")
    [void]$rowsHtml.Append("</tr>")
}

$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>VeeamONE Backup Jobs Report</title>
<style>
  :root { --bg:#f4f6f8; --card:#ffffff; --fg:#1a2733; --muted:#6b7885; --line:#e3e8ee; --accent:#00b336; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#0f1620; --card:#17212e; --fg:#e8eef4; --muted:#8ea0b0; --line:#243546; }
  }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI',Roboto,Arial,sans-serif; background:var(--bg); color:var(--fg); }
  .wrap { max-width:1280px; margin:0 auto; padding:24px; }
  header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
  h1 { font-size:20px; margin:0; }
  .sub { color:var(--muted); font-size:13px; }
  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:14px; margin:20px 0; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:16px; }
  .card .n { font-size:26px; font-weight:700; }
  .card .l { color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }
  .charts { display:grid; grid-template-columns:260px 1fr; gap:16px; margin-bottom:22px; }
  @media (max-width:820px){ .charts{ grid-template-columns:1fr; } }
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
      <h1>VeeamONE - Backup Jobs Report</h1>
      <div class='sub'>Job inventory &middot; schedule, last run &amp; status$(if ($hasPolicy) { ' by backup policy' }) &middot; DB: $(Get-Html $connInfo.Database)</div>
    </div>
    <div class='sub'>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
  </header>

  <div class='cards'>
    <div class='card'><div class='n'>$total</div><div class='l'>Backup jobs</div></div>
    $(if ($hasPolicy) { "<div class='card'><div class='n'>$policyCount</div><div class='l'>Backup policies</div></div>" })
    <div class='card'><div class='n' style='color:$($statusColors.Success)'>$succ</div><div class='l'>Last run: success</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Warning)'>$warn</div><div class='l'>Last run: warning</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Failed)'>$fail</div><div class='l'>Last run: failed</div></div>
    <div class='card'><div class='n'>$disabled</div><div class='l'>Disabled</div></div>
  </div>

  <div class='charts'>
    <div class='panel'>
      <h2>Last-run status</h2>
      <div class='donutwrap'>$donut $legend</div>
    </div>
    <div class='panel'>
      <h2>$barTitle</h2>
      $policyBar
    </div>
  </div>

  <div class='tablewrap'>
    <table>
      <thead><tr>
        $(if ($hasPolicy) { '<th>Backup policy</th>' })<th>Job name</th><th>Type</th><th>Schedule</th><th>Retention</th><th>Enabled</th>
        <th>Last start</th><th>Last end</th><th>Next run</th><th>Status</th><th>Message</th>
      </tr></thead>
      <tbody>
        $($rowsHtml.ToString())
      </tbody>
    </table>
  </div>

  <footer>VeeamONE Backup Jobs Report &middot; source data from the Veeam ONE database (read-only)</footer>
</div>
</body>
</html>
"@

$htmlPath = Join-Path $OutputFolder "VeeamONE-BackupJobs-Report-$stamp.html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "HTML : $htmlPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 8. Optional PDF
# ---------------------------------------------------------------------------
$pdfPath = $null
if ($Pdf) {
    $pdfPath = Join-Path $OutputFolder "VeeamONE-BackupJobs-Report-$stamp.pdf"
    try { ConvertTo-Pdf -HtmlPath $htmlPath -PdfPath $pdfPath | Out-Null; Write-Host "PDF  : $pdfPath" -ForegroundColor Green }
    catch { Write-Warning "PDF generation failed: $($_.Exception.Message)"; $pdfPath = $null }
}

# ---------------------------------------------------------------------------
# 9. Optional email delivery
# ---------------------------------------------------------------------------
if ($EmailTo -and $EmailTo.Count -gt 0) {
    $subject = if ($EmailSubject) { $EmailSubject } else {
        "VeeamONE Backup Jobs Report - $total jobs $policyText ($fail failed / $warn warning)"
    }
    $bodyHtml = New-EmailBodyHtml `
        -Stats @{ Total = $total; Success = $succ; Warning = $warn; Failed = $fail; Disabled = $disabled } `
        -Db ([string]$connInfo.Database) -Colors $statusColors
    $attachments = @($csvPath, $htmlPath, $pdfPath) | Where-Object { $_ }
    try {
        Send-ReportEmail -To $EmailTo -From $EmailFrom -SmtpServer $SmtpServer -SmtpPort $SmtpPort `
            -UseSsl:$SmtpUseSsl -Cred $SmtpCredential -Subject $subject -BodyHtml $bodyHtml -Attachments $attachments
        Write-Host "Email: sent to $($EmailTo -join ', ') via $SmtpServer`:$SmtpPort" -ForegroundColor Green
    }
    catch { Write-Warning "Email delivery failed: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------------------
# 10. Optional A5 backup calendar (print-friendly schedule timetable)
# ---------------------------------------------------------------------------
$calPath = $null
$calPdfPath = $null
if ($Calendar) {
    $order = @{ 'Daily'=1; 'Periodically'=2; 'Continuously'=3; 'Weekly'=4; 'Monthly'=5; 'Triggered / rescan'=9 }
    $calGroups = $records | Group-Object SchedType |
        Sort-Object @{ e = { if ($order.ContainsKey($_.Name)) { $order[$_.Name] } else { 7 } } }, Name
    $calSections = [System.Text.StringBuilder]::new()
    foreach ($cg in $calGroups) {
        $sorted = $cg.Group | Sort-Object @{ e = { if ($_.RunTime) { $_.RunTime } else { '~~:~~' } } }, JobName
        $rowsC = [System.Text.StringBuilder]::new()
        foreach ($r in $sorted) {
            $dis = if ($r.Enabled -eq 'Disabled') { " <span class='dis'>(disabled)</span>" } else { '' }
            [void]$rowsC.Append("<tr>")
            [void]$rowsC.Append("<td>$(Get-Html $r.JobName)$dis</td>")
            [void]$rowsC.Append("<td>$(Get-Html $r.JobType)</td>")
            [void]$rowsC.Append("<td>$(Get-Html $r.NextRun)</td>")
            [void]$rowsC.Append("<td>$(Get-Html $r.LastStart)</td>")
            [void]$rowsC.Append("<td>$(Get-Html $r.LastEnd)</td>")
            [void]$rowsC.Append("</tr>")
        }
        [void]$calSections.Append("<h2>$(Get-Html $cg.Name) &middot; $($cg.Count) job(s)</h2>")
        [void]$calSections.Append("<table><thead><tr><th>Job</th><th>Type</th><th>Next run</th><th>Last start</th><th>Last end</th></tr></thead><tbody>$($rowsC.ToString())</tbody></table>")
    }

    $calHtml = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<title>VeeamONE Backup Calendar</title>
<style>
  @page { size: A5 portrait; margin: 8mm; }
  :root { --fg:#1a2733; --muted:#6b7885; --line:#d8dee6; --accent:#00b336; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI',Arial,sans-serif; color:var(--fg); font-size:10px; }
  h1 { font-size:15px; margin:0 0 2px; }
  .sub { color:var(--muted); font-size:9px; margin-bottom:8px; }
  h2 { font-size:11px; margin:10px 0 3px; padding:3px 6px; background:var(--accent); color:#fff; border-radius:4px; }
  table { border-collapse:collapse; width:100%; margin-bottom:4px; }
  th,td { padding:3px 5px; text-align:left; border-bottom:1px solid var(--line); vertical-align:top; }
  th { color:var(--muted); font-weight:600; font-size:9px; text-transform:uppercase; letter-spacing:.03em; }
  td { font-size:9.5px; }
  .tm { white-space:nowrap; font-variant-numeric:tabular-nums; font-weight:600; width:44px; }
  .dis { color:#e5202e; }
  footer { color:var(--muted); font-size:8px; margin-top:6px; }
</style>
</head>
<body>
  <h1>Backup Calendar</h1>
  <div class='sub'>Configured schedule &amp; time to run &middot; $total jobs $policyText &middot; DB: $(Get-Html $connInfo.Database) &middot; generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
  $($calSections.ToString())
  <footer>VeeamONE Backup Calendar (A5) &middot; times are the job's configured run time; "Triggered / rescan" jobs (e.g. RMAN unmanaged) run on demand.</footer>
</body>
</html>
"@
    $calPath = Join-Path $OutputFolder "VeeamONE-BackupJobs-Calendar-$stamp.html"
    $calHtml | Out-File -FilePath $calPath -Encoding UTF8
    Write-Host "Cal  : $calPath" -ForegroundColor Green
    if ($Pdf) {
        $calPdfPath = Join-Path $OutputFolder "VeeamONE-BackupJobs-Calendar-$stamp.pdf"
        try { ConvertTo-Pdf -HtmlPath $calPath -PdfPath $calPdfPath | Out-Null; Write-Host "CalPDF: $calPdfPath" -ForegroundColor Green }
        catch { Write-Warning "Calendar PDF failed: $($_.Exception.Message)"; $calPdfPath = $null }
    }
}

Write-Host "`nSummary: $total jobs | $policyText | $succ success | $warn warning | $fail failed | $disabled disabled" -ForegroundColor Cyan
[pscustomobject]@{ Csv = $csvPath; Html = $htmlPath; Pdf = $pdfPath; Calendar = $calPath; CalendarPdf = $calPdfPath; Total = $total; Policies = $policyCount; Success = $succ; Warning = $warn; Failed = $fail; Disabled = $disabled }
