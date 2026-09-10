<#
.SYNOPSIS
    VeeamONE All Backup Jobs report - a comprehensive log of every backup job RUN
    (session) with status, type, mode, start / end, duration, data sizes,
    processing rate and failure reasons.

.DESCRIPTION
    Runs on (or next to) the Veeam ONE server. Queries every backup job session
    Veeam ONE collected from the monitored Veeam Backup & Replication servers
    ([monitor].[BpJobSession]) and builds a rich CSV + HTML report:

      - Summary tiles (runs, success / warning / failed, success rate, data moved)
      - Success/failure donut and a daily trend
      - Status-by-job-type breakdown and a "top failing jobs" panel
      - A detailed table: job, type, mode, status, start, end, duration,
        data read, transferred, processing rate and the failure/warning message

.PARAMETER SqlServer / Database / SqlCredential / WindowsCredential
    Connection to the Veeam ONE database (same model as the other reports).

.PARAMETER StartDate / EndDate / StartTime / EndTime
    Reporting window on the session start time (default last 30 days).

.PARAMETER NameLike
    Optional SQL LIKE filter on the job name (e.g. "%SQL%").

.PARAMETER JobType
    Optional one or more numeric job type codes to include (see -Discover).

.PARAMETER StatusFilter
    Optional: keep only Success / Warning / Failed rows.

.PARAMETER OutputFolder / Pdf / Period
    Output location, optional PDF, and relative window for scheduling.

.PARAMETER EmailTo / EmailFrom / SmtpServer / SmtpPort / SmtpUseSsl / SmtpCredential / EmailSubject
    Optional e-mail delivery (attachments: CSV + HTML + PDF).

.PARAMETER Discover
    List the session result codes and job types present (with sample names) and exit.

.PARAMETER DemoData
    Build the report from synthetic data (no DB) to preview the format.

.EXAMPLE
    .\Get-VeeamOneAllBackupsReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)

.EXAMPLE
    # Only failed backups in July, e-mailed as PDF
    .\Get-VeeamOneAllBackupsReport.ps1 -StatusFilter Failed -StartDate 2026-07-01 -EndDate 2026-08-01 `
        -Pdf -EmailTo ops@contoso.com -SmtpServer smtp.contoso.com

.EXAMPLE
    .\Get-VeeamOneAllBackupsReport.ps1 -DemoData
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
    [string]   $NameLike,
    [Parameter(ParameterSetName = 'Report')]
    [int[]]    $JobType,
    [Parameter(ParameterSetName = 'Report')]
    [ValidateSet('Success','Warning','Failed')]
    [string]   $StatusFilter,

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
#    BpJobSession.[result]: 2 = Success, 3 = Warning, 4 = Failed (confirmed).
# ---------------------------------------------------------------------------
$script:Schema = @{
    Table       = '[monitor].[BpJobSession]'
    IdCol       = 'uid'
    JobNameCol  = 'job_name'
    JobTypeCol  = 'job_type'
    IsFullCol   = 'is_full'
    ResultCol   = 'result'
    StartCol    = 'start_time'
    DurationCol = 'duration'
    ReadCol     = 'backedup_size'
    XferCol     = 'transferred_size'
    RateCol     = 'processing_rate'   # bytes / second
    RetryCol    = 'will_be_retried'
    ReasonCol   = 'failure_message'
}
$script:ResultMap  = @{ 2 = 'Success'; 3 = 'Warning'; 4 = 'Failed' }
# Friendly Veeam job-type names (numeric fallback used by Get-JobTypeName).
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
# 1. Connection helpers
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
    $base = "Server=$Server;Database=$Database;Application Name=VeeamOneAllBackupsReport;Connect Timeout=30;"
    if ($Cred) {
        $u = $Cred.UserName; $p = $Cred.GetNetworkCredential().Password
        return "$base User ID=$u;Password=$p;"
    }
    return "$base Integrated Security=SSPI;"
}

function Enter-DomainImpersonation {
    param([pscredential]$Credential)
    if (-not $Credential) { return $null }
    if (-not ([System.Management.Automation.PSTypeName]'VeeamOneAll.NativeLogon').Type) {
        Add-Type -UsingNamespace 'Microsoft.Win32.SafeHandles' -Namespace 'VeeamOneAll' -Name 'NativeLogon' -MemberDefinition @'
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
    $ok = [VeeamOneAll.NativeLogon]::LogonUser($user, $domain, $pw, 9, 3, [ref]$token)
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
function ConvertTo-StatusText {
    param($ResultCode, $Message)
    if ($null -ne $ResultCode -and $ResultCode -isnot [DBNull]) {
        $n = 0
        if ([int]::TryParse([string]$ResultCode, [ref]$n) -and $script:ResultMap.ContainsKey($n)) { return $script:ResultMap[$n] }
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

function Format-Size {
    param($Bytes)
    if ($null -eq $Bytes -or $Bytes -is [DBNull]) { return '' }
    $b = 0.0
    if (-not [double]::TryParse([string]$Bytes, [ref]$b) -or $b -le 0) { return '' }
    $u = 'B','KB','MB','GB','TB','PB'; $i = 0
    while ($b -ge 1024 -and $i -lt $u.Count - 1) { $b /= 1024; $i++ }
    '{0:N2} {1}' -f $b, $u[$i]
}
function Get-Bytes {
    param($Value)
    $b = 0.0
    if ($Value -isnot [DBNull] -and [double]::TryParse([string]$Value, [ref]$b) -and $b -gt 0) { return [double]$b }
    return 0.0
}
function Format-Rate {
    param($BytesPerSec)
    $r = Get-Bytes $BytesPerSec
    if ($r -le 0) { return '' }
    (Format-Size $r) + '/s'
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
    param([hashtable]$Stats, [string]$Range, [string]$Db, [hashtable]$Colors)
    $cell = "padding:6px 14px;border:1px solid #e3e8ee;font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
    $num  = { param($v,$c) "<td style='$cell text-align:center;font-weight:700;color:$c;'>$v</td>" }
    @"
<div style='font-family:Segoe UI,Arial,sans-serif;color:#1a2733;'>
  <h2 style='margin:0 0 4px;'>VeeamONE - All Backup Jobs Report</h2>
  <div style='color:#6b7885;font-size:13px;margin-bottom:14px;'>$Range &middot; DB: $Db</div>
  <table style='border-collapse:collapse;margin-bottom:12px;'>
    <tr>
      <td style='$cell color:#6b7885;'>Runs</td>
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
    if (-not $From) { $From = "VeeamONE-AllBackups-Report@$($env:COMPUTERNAME)" }
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
    Write-Host "DEMO MODE - generating synthetic all-backups report (no database queried)." -ForegroundColor Magenta
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
    $S = $script:Schema
    Write-Host "`n=== Distinct '$($S.ResultCol)' values in $($S.Table) (last 90 days) ===" -ForegroundColor Yellow
    try {
        $sql = @"
SELECT s.[$($S.ResultCol)] AS ResultCode, COUNT(*) AS Sessions, MAX(s.[$($S.ReasonCol)]) AS SampleMessage
FROM $($S.Table) s WHERE s.[$($S.StartCol)] >= @start
GROUP BY s.[$($S.ResultCol)] ORDER BY Sessions DESC
"@
        Invoke-VeeamSql -ConnectionString $connStr -Params @{ start = (Get-Date).AddDays(-90) } -Query $sql |
            Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Could not enumerate results: $($_.Exception.Message)" }

    Write-Host "=== Distinct '$($S.JobTypeCol)' values (with sample job names) ===" -ForegroundColor Yellow
    try {
        $sql = @"
SELECT s.[$($S.JobTypeCol)] AS JobType, COUNT(*) AS Sessions, MAX(s.[$($S.JobNameCol)]) AS SampleJob
FROM $($S.Table) s WHERE s.[$($S.StartCol)] >= @start
GROUP BY s.[$($S.JobTypeCol)] ORDER BY Sessions DESC
"@
        Invoke-VeeamSql -ConnectionString $connStr -Params @{ start = (Get-Date).AddDays(-90) } -Query $sql |
            Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Could not enumerate job types: $($_.Exception.Message)" }
    Write-Host "Discovery complete." -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# 5. Build the record set
# ---------------------------------------------------------------------------
$records = New-Object System.Collections.Generic.List[object]

if ($DemoData) {
    $jobs = @(
        @{ n='BKP-PDC1_PRD_Linux_CLS01 - WEBUSR01'; t=1  },
        @{ n='TSTEDIAPPS01 Backup';                 t=18 },
        @{ n='BKP-PG-PRDSQLDBCLS29 - SQL Server Transaction Log Backup'; t=10 },
        @{ n='BKP-PRDODB09-DMSPP-Full_RMAN';        t=15 },
        @{ n='REP-PRZ_AppDynamics_DR';              t=2  },
        @{ n='CPY-PDC1_PRD_Windows_CLS01';          t=16 },
        @{ n='prdeisap01 SAP backint backup';       t=32 },
        @{ n='prdodb81-scan Oracle backup';         t=33 },
        @{ n='TSTAPISQLDB01 Backup';                t=18 },
        @{ n='PRDODB09 Backup';                     t=18 }
    )
    $rng = New-Object System.Random 20260812
    # Generate across the SELECTED [StartDate, EndDate] window so the demo honors
    # the picked date range (capped to 60 days to keep the sample small).
    $spanDays = [math]::Min(60, [math]::Max(1, [int][math]::Ceiling(($EndDate - $StartDate).TotalDays)))
    for ($d = 0; $d -lt $spanDays; $d++) {
        $day = $StartDate.Date.AddDays($d)
        foreach ($i in 0..($rng.Next(3,9))) {
            $j = $jobs[$rng.Next(0,$jobs.Count)]
            $stt = $day.AddHours($rng.Next(0,23)).AddMinutes($rng.Next(0,59))
            if ($stt -lt $StartDate -or $stt -ge $EndDate) { continue }   # keep within the picked range
            $dur = $rng.Next(20, 7200)
            $roll = $rng.NextDouble()
            $status = if ($roll -gt 0.90) { 'Failed' } elseif ($roll -gt 0.80) { 'Warning' } else { 'Success' }
            $isFull = $rng.Next(0,2)
            $read = [long]($rng.Next(200, 90000)) * 1MB
            $xfer = if ($status -eq 'Failed') { 0 } else { [long]($read * ($rng.NextDouble()*0.4)) }
            $rate = if ($dur -gt 0 -and $status -ne 'Failed') { [long]($read / $dur) } else { 0 }
            $en = $stt.AddSeconds($dur)
            $records.Add([pscustomobject]@{
                JobName     = $j.n
                JobType     = Get-JobTypeName $j.n $j.t
                Mode        = if ($isFull) { 'Full' } else { 'Incremental' }
                Status      = $status
                StartDate   = $stt.ToString('yyyy-MM-dd')
                StartTime   = $stt.ToString('HH:mm:ss')
                EndDate     = $en.ToString('yyyy-MM-dd')
                EndTime     = $en.ToString('HH:mm:ss')
                ElapsedTime = Format-DurationSec $dur
                DataRead    = Format-Size $read
                Transferred = Format-Size $xfer
                Rate        = Format-Rate $rate
                Retried     = if ($status -eq 'Failed' -and $rng.NextDouble() -gt 0.5) { 'Yes' } else { '' }
                Reason      = if ($status -eq 'Failed') { 'Task failed. Error: Failed to connect to the host (demo)' } elseif ($status -eq 'Warning') { 'Completed with warnings (demo)' } else { '' }
                XferBytes   = $xfer
            }) | Out-Null
        }
    }
    # NOTE: $StartDate / $EndDate are the picked values - do NOT overwrite them,
    # so the demo output reflects the selected date range.
}
else {
    $S = $script:Schema
    $filters = New-Object System.Collections.Generic.List[string]
    $sqlParams = @{}
    if ($NameLike) { $filters.Add("s.[$($S.JobNameCol)] LIKE @namelike"); $sqlParams['namelike'] = $NameLike }
    if ($JobType -and $JobType.Count -gt 0) { $filters.Add("s.[$($S.JobTypeCol)] IN ($([string]::Join(',', $JobType)))") }
    if ($StatusFilter) {
        $code = ($script:ResultMap.GetEnumerator() | Where-Object { $_.Value -eq $StatusFilter } | ForEach-Object { $_.Key } | Select-Object -First 1)
        if ($null -ne $code) { $filters.Add("s.[$($S.ResultCol)] = $code") }
    }
    $extra = if ($filters.Count -gt 0) { 'AND ' + ($filters -join ' AND ') } else { '' }

    $sql = @"
SELECT
    s.[$($S.JobNameCol)]  AS JobName,
    s.[$($S.JobTypeCol)]  AS JobType,
    s.[$($S.IsFullCol)]   AS IsFull,
    s.[$($S.ResultCol)]   AS ResultCode,
    s.[$($S.StartCol)]    AS StartDate,
    s.[$($S.DurationCol)] AS DurationSec,
    s.[$($S.ReadCol)]     AS ReadBytes,
    s.[$($S.XferCol)]     AS XferBytes,
    s.[$($S.RateCol)]     AS RateBytes,
    s.[$($S.RetryCol)]    AS WillRetry,
    s.[$($S.ReasonCol)]   AS Reason
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start
  AND s.[$($S.StartCol)] <  @end
  $extra
ORDER BY s.[$($S.StartCol)] DESC
"@
    $endExclusive = if ($EndDate.TimeOfDay -eq [TimeSpan]::Zero) { $EndDate.Date.AddDays(1) } else { $EndDate }
    $sqlParams['start'] = $StartDate; $sqlParams['end'] = $endExclusive

    Write-Host "Querying all backup sessions $($StartDate.ToString('yyyy-MM-dd')) .. $($EndDate.ToString('yyyy-MM-dd'))$(if ($StatusFilter) { " [$StatusFilter only]" })..." -ForegroundColor Cyan
    try {
        $sessions = Invoke-VeeamSql -ConnectionString $connStr -Query $sql -Params $sqlParams
    }
    catch {
        Write-Warning "Backup-session query failed against $($S.Table): $($_.Exception.Message)"
        Write-Warning "Run -Discover to confirm the codes on your build."
        return
    }

    foreach ($s in $sessions.Rows) {
        $start = [datetime]$s['StartDate']
        if (-not (Test-InTimeWindow -When $start -From $StartTime -To $EndTime)) { continue }
        $durSec = 0.0; [double]::TryParse([string]$s['DurationSec'], [ref]$durSec) | Out-Null
        $end = $start.AddSeconds([int]$durSec)
        $isFull = ($s['IsFull'] -isnot [DBNull]) -and ([string]$s['IsFull'] -eq '1' -or [string]$s['IsFull'] -eq 'True')
        $retry  = ($s['WillRetry'] -isnot [DBNull]) -and ([string]$s['WillRetry'] -eq '1' -or [string]$s['WillRetry'] -eq 'True')
        $records.Add([pscustomobject]@{
            JobName     = [string]$s['JobName']
            JobType     = Get-JobTypeName ([string]$s['JobName']) $s['JobType']
            Mode        = if ($isFull) { 'Full' } else { 'Incremental' }
            Status      = ConvertTo-StatusText $s['ResultCode'] $s['Reason']
            StartDate   = $start.ToString('yyyy-MM-dd')
            StartTime   = $start.ToString('HH:mm:ss')
            EndDate     = $end.ToString('yyyy-MM-dd')
            EndTime     = $end.ToString('HH:mm:ss')
            ElapsedTime = Format-DurationSec $durSec
            DataRead    = Format-Size $s['ReadBytes']
            Transferred = Format-Size $s['XferBytes']
            Rate        = Format-Rate $s['RateBytes']
            Retried     = if ($retry) { 'Yes' } else { '' }
            Reason      = if ($s['Reason'] -isnot [DBNull]) { [string]$s['Reason'] } else { '' }
            XferBytes   = Get-Bytes $s['XferBytes']
        }) | Out-Null
    }
}

Write-Host "Collected $($records.Count) backup session(s)." -ForegroundColor Green
if ($records.Count -eq 0) {
    Write-Warning "No backup sessions matched. Widen the date range or adjust the filters (-NameLike / -JobType / -StatusFilter)."
    return
}

# Refine plain "Veeam Agent Backup" rows into SQL vs Oracle/RMAN-host (script).
# VeeamONE does not expose the job's pre/post script, so an agent job on a host
# that ALSO has an RMAN plugin job is inferred to be the script-driven RMAN case;
# a SQL-named host is inferred to be a SQL server.
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

# ---------------------------------------------------------------------------
# 6. Export CSV
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath = Join-Path $OutputFolder "VeeamONE-AllBackups-Report-$stamp.csv"
$records |
    Select-Object JobName, JobType, Mode, Status, StartDate, StartTime, EndDate, EndTime, ElapsedTime, DataRead, Transferred, Rate, Retried, Reason |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV  : $csvPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 7. Build the HTML report
# ---------------------------------------------------------------------------
$total = $records.Count
$succ  = @($records | Where-Object Status -eq 'Success').Count
$warn  = @($records | Where-Object Status -eq 'Warning').Count
$fail  = @($records | Where-Object Status -eq 'Failed').Count
$other = $total - $succ - $warn - $fail
$rate  = if ($total) { [math]::Round(($succ / $total) * 100, 1) } else { 0 }
$totXfer = Format-Size (($records | Measure-Object -Property XferBytes -Sum).Sum)

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

# --- status-by-job-type stacked bar ---
# Full-width job-type label ABOVE each bar so long names (e.g.
# "Oracle RMAN Plugin - full/incremental (unmanaged)") are never clipped by the bars.
function New-TypeBarSvg {
    param($Records, [hashtable]$Colors, [int]$Top = 10)
    $groups = $Records | Group-Object JobType | Sort-Object Count -Descending | Select-Object -First $Top
    if (@($groups).Count -eq 0) { return '' }
    $w = 680; $barMax = $w - 60      # right margin leaves room for the count label
    $labelH = 16; $barH = 15; $block = $labelH + $barH + 12
    $h = (@($groups).Count) * $block + 6
    $max = ($groups | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    if (-not $max) { $max = 1 }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg viewBox='0 0 $w $h' width='100%' height='$h' preserveAspectRatio='xMinYMin meet'>")
    $y = 6
    foreach ($g in $groups) {
        # label on its own line (full width)
        [void]$sb.Append("<text x='2' y='$($y+12)' font-size='11.5' fill='var(--fg)'>$([System.Web.HttpUtility]::HtmlEncode($g.Name))</text>")
        # stacked bar underneath
        $by = $y + $labelH + 1
        $x = 2
        foreach ($st in 'Success','Warning','Failed','Unknown') {
            $c = @($g.Group | Where-Object Status -eq $st).Count
            if ($c -le 0) { continue }
            $bw = $barMax * $c / $max
            $col = if ($Colors.ContainsKey($st)) { $Colors[$st] } else { '#8a8a8a' }
            [void]$sb.Append(("<rect x='{0:N1}' y='$by' width='{1:N1}' height='$barH' rx='2' fill='$col'/>" -f $x,$bw))
            $x += $bw
        }
        [void]$sb.Append("<text x='$([math]::Round($x + 6))' y='$($by+12)' font-size='11' font-weight='600' fill='var(--muted)'>$($g.Count)</text>")
        $y += $block
    }
    [void]$sb.Append('</svg>')
    $sb.ToString()
}
$typeBar = New-TypeBarSvg -Records $records -Colors $statusColors

# --- top failing jobs ---
$failing = $records | Where-Object Status -eq 'Failed' | Group-Object JobName |
    Sort-Object Count -Descending | Select-Object -First 8
$failHtml = [System.Text.StringBuilder]::new()
if (@($failing).Count -gt 0) {
    foreach ($f in $failing) {
        [void]$failHtml.Append("<div class='failrow'><span class='fn'>$(Get-Html $f.Name)</span><span class='fc'>$($f.Count)</span></div>")
    }
} else {
    [void]$failHtml.Append("<div class='sub'>No failed runs in this window.</div>")
}

$legend = @"
<div class='legend'>
  <span><i style='background:$($statusColors.Success)'></i>Success ($succ)</span>
  <span><i style='background:$($statusColors.Warning)'></i>Warning ($warn)</span>
  <span><i style='background:$($statusColors.Failed)'></i>Failed ($fail)</span>
  $(if ($other -gt 0) { "<span><i style='background:$($statusColors.Unknown)'></i>Other ($other)</span>" })
</div>
"@

$rowsHtml = [System.Text.StringBuilder]::new()
foreach ($r in $records) {
    $sc = if ($statusColors.ContainsKey($r.Status)) { $statusColors[$r.Status] } else { '#8a8a8a' }
    [void]$rowsHtml.Append("<tr>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.JobName)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.JobType)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Mode)</td>")
    [void]$rowsHtml.Append("<td><span class='badge' style='background:$sc'>$(Get-Html $r.Status)</span></td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.StartDate) $(Get-Html $r.StartTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.EndDate) $(Get-Html $r.EndTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.ElapsedTime)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.DataRead)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Transferred)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Rate)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Retried)</td>")
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Reason)</td>")
    [void]$rowsHtml.Append("</tr>")
}

$rangeText = "$($StartDate.ToString('yyyy-MM-dd'))  to  $($EndDate.ToString('yyyy-MM-dd'))"
if ($StartTime -or $EndTime) { $rangeText += "   (window $([string]$StartTime)-$([string]$EndTime))" }

$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>VeeamONE All Backup Jobs Report</title>
<style>
  :root { --bg:#f4f6f8; --card:#ffffff; --fg:#1a2733; --muted:#6b7885; --line:#e3e8ee; --accent:#00b336; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#0f1620; --card:#17212e; --fg:#e8eef4; --muted:#8ea0b0; --line:#243546; }
  }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI',Roboto,Arial,sans-serif; background:var(--bg); color:var(--fg); }
  .wrap { max-width:1320px; margin:0 auto; padding:24px; }
  header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }
  h1 { font-size:20px; margin:0; }
  .sub { color:var(--muted); font-size:13px; }
  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:14px; margin:20px 0; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:16px; }
  .card .n { font-size:26px; font-weight:700; }
  .card .l { color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }
  .charts { display:grid; grid-template-columns:240px 1fr; gap:16px; margin-bottom:16px; }
  .charts2 { display:grid; grid-template-columns:1fr 340px; gap:16px; margin-bottom:22px; }
  @media (max-width:900px){ .charts,.charts2{ grid-template-columns:1fr; } }
  .panel { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:16px; }
  .panel h2 { font-size:14px; margin:0 0 10px; color:var(--muted); font-weight:600; }
  .donutwrap { display:flex; flex-direction:column; align-items:center; gap:10px; }
  .legend { display:flex; flex-wrap:wrap; gap:12px; font-size:12px; color:var(--muted); }
  .legend i { display:inline-block; width:10px; height:10px; border-radius:2px; margin-right:5px; vertical-align:middle; }
  .failrow { display:flex; justify-content:space-between; gap:10px; padding:6px 0; border-bottom:1px solid var(--line); font-size:13px; }
  .failrow .fn { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .failrow .fc { color:#e5202e; font-weight:700; }
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
      <h1>VeeamONE - All Backup Jobs Report</h1>
      <div class='sub'>Every backup run &middot; status, type, timing, data &amp; failure reasons &middot; $rangeText &middot; DB: $(Get-Html $connInfo.Database)</div>
    </div>
    <div class='sub'>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
  </header>

  <div class='cards'>
    <div class='card'><div class='n'>$total</div><div class='l'>Total runs</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Success)'>$succ</div><div class='l'>Success</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Warning)'>$warn</div><div class='l'>Warning</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Failed)'>$fail</div><div class='l'>Failed</div></div>
    <div class='card'><div class='n'>$rate%</div><div class='l'>Success rate</div></div>
    <div class='card'><div class='n'>$totXfer</div><div class='l'>Data transferred</div></div>
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

  <div class='charts2'>
    <div class='panel'>
      <h2>Status by job type (top 10)</h2>
      $typeBar
    </div>
    <div class='panel'>
      <h2>Top failing jobs</h2>
      $($failHtml.ToString())
    </div>
  </div>

  <div class='tablewrap'>
    <table>
      <thead><tr>
        <th>Job name</th><th>Type</th><th>Mode</th><th>Status</th>
        <th>Start</th><th>End</th><th>Elapsed</th><th>Data read</th><th>Transferred</th><th>Rate</th><th>Retried</th><th>Reason</th>
      </tr></thead>
      <tbody>
        $($rowsHtml.ToString())
      </tbody>
    </table>
  </div>

  <footer>VeeamONE All Backup Jobs Report &middot; source data from the Veeam ONE database (read-only)</footer>
</div>
</body>
</html>
"@

$htmlPath = Join-Path $OutputFolder "VeeamONE-AllBackups-Report-$stamp.html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "HTML : $htmlPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 8. Optional PDF
# ---------------------------------------------------------------------------
$pdfPath = $null
if ($Pdf) {
    $pdfPath = Join-Path $OutputFolder "VeeamONE-AllBackups-Report-$stamp.pdf"
    try { ConvertTo-Pdf -HtmlPath $htmlPath -PdfPath $pdfPath | Out-Null; Write-Host "PDF  : $pdfPath" -ForegroundColor Green }
    catch { Write-Warning "PDF generation failed: $($_.Exception.Message)"; $pdfPath = $null }
}

# ---------------------------------------------------------------------------
# 9. Optional email delivery
# ---------------------------------------------------------------------------
if ($EmailTo -and $EmailTo.Count -gt 0) {
    $subject = if ($EmailSubject) { $EmailSubject } else {
        "VeeamONE All Backups Report - $rate% success ($total runs, $fail failed) - $($StartDate.ToString('yyyy-MM-dd'))..$($EndDate.ToString('yyyy-MM-dd'))"
    }
    $bodyHtml = New-EmailBodyHtml `
        -Stats @{ Total = $total; Success = $succ; Warning = $warn; Failed = $fail; Rate = $rate } `
        -Range $rangeText -Db ([string]$connInfo.Database) -Colors $statusColors
    $attachments = @($csvPath, $htmlPath, $pdfPath) | Where-Object { $_ }
    try {
        Send-ReportEmail -To $EmailTo -From $EmailFrom -SmtpServer $SmtpServer -SmtpPort $SmtpPort `
            -UseSsl:$SmtpUseSsl -Cred $SmtpCredential -Subject $subject -BodyHtml $bodyHtml -Attachments $attachments
        Write-Host "Email: sent to $($EmailTo -join ', ') via $SmtpServer`:$SmtpPort" -ForegroundColor Green
    }
    catch { Write-Warning "Email delivery failed: $($_.Exception.Message)" }
}

Write-Host "`nSummary: $total runs | $succ success | $warn warning | $fail failed | $rate% success rate | $totXfer transferred" -ForegroundColor Cyan
[pscustomobject]@{ Csv = $csvPath; Html = $htmlPath; Pdf = $pdfPath; Total = $total; Success = $succ; Warning = $warn; Failed = $fail; SuccessRate = $rate }
