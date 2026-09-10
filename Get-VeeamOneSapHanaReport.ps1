<#
.SYNOPSIS
    VeeamONE SAP HANA (backint) plugin backup report - success / warning /
    failure rate for SAP HANA plugin backups, covering BOTH plugin modes:
    VBR-managed (application backup policy) and standalone / unmanaged (backint
    configured on the DB host).

.DESCRIPTION
    Runs on (or next to) the Veeam ONE server. Queries the SAP HANA plugin
    backup sessions Veeam ONE collected ([monitor].[BpJobSession], job name
    matching -JobNameLike, default "%backint%") and builds a CSV + HTML report
    with a success/failure donut, a daily trend, a Managed-vs-Unmanaged split,
    and a detailed table.

    Managed vs unmanaged is derived from the job name: VBR-managed application
    policies carry a "DBBKP-<host>-SAP\..." prefix; bare "<host> SAP backint
    backup (<repo>)" jobs are standalone / unmanaged backint.

.PARAMETER SqlServer / Database / SqlCredential / WindowsCredential
    Connection to the Veeam ONE database (same model as the other reports).

.PARAMETER StartDate / EndDate / StartTime / EndTime
    Reporting window on the session start time (default last 30 days).

.PARAMETER JobNameLike
    SQL LIKE filter identifying SAP HANA plugin jobs. Default "%backint%".

.PARAMETER Deployment
    Optional: keep only 'Managed' or 'Unmanaged' SAP HANA plugin backups.

.PARAMETER OutputFolder / Pdf / Period
    Output location, optional PDF, and relative window for scheduling.

.PARAMETER EmailTo / EmailFrom / SmtpServer / SmtpPort / SmtpUseSsl / SmtpCredential / EmailSubject
    Optional e-mail delivery (attachments: CSV + HTML + PDF).

.PARAMETER Discover
    List the SAP-HANA session result codes present and exit.

.PARAMETER DemoData
    Build the report from synthetic data (no DB) to preview the format.

.EXAMPLE
    .\Get-VeeamOneSapHanaReport.ps1 -SqlServer YOURSQL -SqlCredential (Get-Credential sa)

.EXAMPLE
    # Only standalone/unmanaged backint, last month, e-mailed as PDF
    .\Get-VeeamOneSapHanaReport.ps1 -Deployment Unmanaged -Period Monthly -Pdf `
        -EmailTo ops@contoso.com -SmtpServer smtp.contoso.com

.EXAMPLE
    .\Get-VeeamOneSapHanaReport.ps1 -DemoData
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
    [string]   $JobNameLike = '%backint%',
    [Parameter(ParameterSetName = 'Report')]
    [ValidateSet('Managed','Unmanaged')]
    [string]   $Deployment,

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
# 0. Version-sensitive mapping (adjust after -Discover). BpJobSession.[result]:
#    2 = Success, 3 = Warning, 4 = Failed (confirmed Veeam ONE build).
# ---------------------------------------------------------------------------
$script:Schema = @{
    Table       = '[monitor].[BpJobSession]'
    IdCol       = 'uid'
    JobNameCol  = 'job_name'
    IsFullCol   = 'is_full'
    ResultCol   = 'result'
    StartCol    = 'start_time'
    DurationCol = 'duration'
    ReadCol     = 'backedup_size'
    XferCol     = 'transferred_size'
    ReasonCol   = 'failure_message'
}
$script:ResultMap = @{ 2 = 'Success'; 3 = 'Warning'; 4 = 'Failed' }

# Managed (VBR application policy) vs standalone/unmanaged backint, from the name.
function Get-SapDeployment {
    param([string]$JobName)
    if ($JobName -match '(?i)DBBKP-') { 'Managed (VBR policy)' } else { 'Standalone / unmanaged' }
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
    $base = "Server=$Server;Database=$Database;Application Name=VeeamOneSapHanaReport;Connect Timeout=30;"
    if ($Cred) {
        $u = $Cred.UserName; $p = $Cred.GetNetworkCredential().Password
        return "$base User ID=$u;Password=$p;"
    }
    return "$base Integrated Security=SSPI;"
}

function Enter-DomainImpersonation {
    param([pscredential]$Credential)
    if (-not $Credential) { return $null }
    if (-not ([System.Management.Automation.PSTypeName]'VeeamOneSap.NativeLogon').Type) {
        Add-Type -UsingNamespace 'Microsoft.Win32.SafeHandles' -Namespace 'VeeamOneSap' -Name 'NativeLogon' -MemberDefinition @'
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
    $ok = [VeeamOneSap.NativeLogon]::LogonUser($user, $domain, $pw, 9, 3, [ref]$token)
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
  <h2 style='margin:0 0 4px;'>VeeamONE - SAP HANA (backint) Plugin Report</h2>
  <div style='color:#6b7885;font-size:13px;margin-bottom:14px;'>Managed + unmanaged &middot; $Range &middot; DB: $Db</div>
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
    if (-not $From) { $From = "VeeamONE-SAPHANA-Report@$($env:COMPUTERNAME)" }
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
    Write-Host "DEMO MODE - generating synthetic SAP HANA report (no database queried)." -ForegroundColor Magenta
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
    Write-Host "`n=== SAP HANA (job_name LIKE '$JobNameLike') sessions by result (last 90 days) ===" -ForegroundColor Yellow
    try {
        $sql = @"
SELECT s.[$($S.ResultCol)] AS ResultCode, COUNT(*) AS Sessions, MAX(s.[$($S.JobNameCol)]) AS SampleJob
FROM $($S.Table) s
WHERE s.[$($S.StartCol)] >= @start AND s.[$($S.JobNameCol)] LIKE @jn
GROUP BY s.[$($S.ResultCol)] ORDER BY Sessions DESC
"@
        Invoke-VeeamSql -ConnectionString $connStr -Params @{ start = (Get-Date).AddDays(-90); jn = $JobNameLike } -Query $sql |
            Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
    } catch { Write-Warning "Could not enumerate SAP HANA sessions: $($_.Exception.Message)" }
    Write-Host "Discovery complete." -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# 5. Build the record set
# ---------------------------------------------------------------------------
$records = New-Object System.Collections.Generic.List[object]

if ($DemoData) {
    $jobs = @(
        'DBBKP-PRD4HNBWDB1-SAP\prd4hanabwdb1 SAP backint backup (BKP-SAP-HAN)',
        'DBBKP-STGS4HNAQADB01-SAP\stgs4hnaqadb01 SAP backint backup (BKP-SAP-HAN)',
        'prdeisap01 SAP backint backup (PRDVEEAMLORP03-Nimble04-Repo1)',
        'tsts4hanasnddb1 SAP backint backup (PRDVeeamNRP01-Nimble01-Repo01)',
        'preeradsap01 SAP backint backup (BKP-SAP-HAN)'
    )
    $rng = New-Object System.Random 20260812
    # Generate across the SELECTED [StartDate, EndDate] so the demo honors the picked range.
    $spanDays = [math]::Min(60, [math]::Max(1, [int][math]::Ceiling(($EndDate - $StartDate).TotalDays)))
    for ($d = 0; $d -lt $spanDays; $d++) {
        $day = $StartDate.Date.AddDays($d)
        foreach ($i in 0..($rng.Next(2,6))) {
            $jn = $jobs[$rng.Next(0,$jobs.Count)]
            $stt = $day.AddHours($rng.Next(0,23)).AddMinutes($rng.Next(0,59))
            if ($stt -lt $StartDate -or $stt -ge $EndDate) { continue }
            $dur = $rng.Next(30, 5400)
            $roll = $rng.NextDouble()
            $status = if ($roll -gt 0.92) { 'Failed' } elseif ($roll -gt 0.82) { 'Warning' } else { 'Success' }
            $isFull = $rng.Next(0,2)
            $read = [long]($rng.Next(500, 80000)) * 1MB
            $xfer = if ($status -eq 'Failed') { 0 } else { [long]($read * ($rng.NextDouble()*0.35)) }
            $en = $stt.AddSeconds($dur)
            $records.Add([pscustomobject]@{
                JobName     = $jn
                Deployment  = Get-SapDeployment $jn
                Mode        = if ($isFull) { 'Full' } else { 'Incremental / log' }
                Status      = $status
                StartDate   = $stt.ToString('yyyy-MM-dd')
                StartTime   = $stt.ToString('HH:mm:ss')
                EndDate     = $en.ToString('yyyy-MM-dd')
                EndTime     = $en.ToString('HH:mm:ss')
                ElapsedTime = Format-DurationSec $dur
                DataRead    = Format-Size $read
                Transferred = Format-Size $xfer
                Reason      = if ($status -eq 'Failed') { 'backint: pipe closed unexpectedly (demo)' } elseif ($status -eq 'Warning') { 'Completed with warnings (demo)' } else { '' }
            }) | Out-Null
        }
    }
    # $StartDate / $EndDate kept as picked so the demo reflects the selected range.
}
else {
    $S = $script:Schema
    $sql = @"
SELECT
    s.[$($S.JobNameCol)]  AS JobName,
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
  AND s.[$($S.JobNameCol)] LIKE @jobname
ORDER BY s.[$($S.StartCol)] DESC
"@
    $endExclusive = if ($EndDate.TimeOfDay -eq [TimeSpan]::Zero) { $EndDate.Date.AddDays(1) } else { $EndDate }
    $sqlParams = @{ start = $StartDate; end = $endExclusive; jobname = $JobNameLike }

    Write-Host "Querying SAP HANA sessions $($StartDate.ToString('yyyy-MM-dd')) .. $($EndDate.ToString('yyyy-MM-dd')) [name ~ '$JobNameLike']..." -ForegroundColor Cyan
    try {
        $sessions = Invoke-VeeamSql -ConnectionString $connStr -Query $sql -Params $sqlParams
    }
    catch {
        Write-Warning "SAP HANA session query failed against $($S.Table): $($_.Exception.Message)"
        Write-Warning "Run -Discover to confirm the result codes / job-name filter on your build."
        return
    }

    foreach ($s in $sessions.Rows) {
        $start = [datetime]$s['StartDate']
        if (-not (Test-InTimeWindow -When $start -From $StartTime -To $EndTime)) { continue }
        $dep = Get-SapDeployment ([string]$s['JobName'])
        if ($Deployment -eq 'Managed'   -and $dep -notmatch '(?i)^managed')  { continue }
        if ($Deployment -eq 'Unmanaged' -and $dep -notmatch '(?i)unmanaged') { continue }
        $durSec = 0.0; [double]::TryParse([string]$s['DurationSec'], [ref]$durSec) | Out-Null
        $end = $start.AddSeconds([int]$durSec)
        $isFull = ($s['IsFull'] -isnot [DBNull]) -and ([string]$s['IsFull'] -eq '1' -or [string]$s['IsFull'] -eq 'True')
        $records.Add([pscustomobject]@{
            JobName     = [string]$s['JobName']
            Deployment  = $dep
            Mode        = if ($isFull) { 'Full' } else { 'Incremental / log' }
            Status      = ConvertTo-StatusText $s['ResultCode'] $s['Reason']
            StartDate   = $start.ToString('yyyy-MM-dd')
            StartTime   = $start.ToString('HH:mm:ss')
            EndDate     = $end.ToString('yyyy-MM-dd')
            EndTime     = $end.ToString('HH:mm:ss')
            ElapsedTime = Format-DurationSec $durSec
            DataRead    = Format-Size $s['ReadBytes']
            Transferred = Format-Size $s['XferBytes']
            Reason      = if ($s['Reason'] -isnot [DBNull]) { [string]$s['Reason'] } else { '' }
        }) | Out-Null
    }
}

Write-Host "Collected $($records.Count) SAP HANA session(s)." -ForegroundColor Green
if ($records.Count -eq 0) {
    Write-Warning "No SAP HANA sessions matched. Widen the date range, adjust -JobNameLike (default %backint%), or use -Discover."
    return
}

# ---------------------------------------------------------------------------
# 6. Export CSV
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath = Join-Path $OutputFolder "VeeamONE-SAPHANA-Report-$stamp.csv"
$records |
    Select-Object JobName, Deployment, Mode, Status, StartDate, StartTime, EndDate, EndTime, ElapsedTime, DataRead, Transferred, Reason |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV  : $csvPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 7. Build the HTML report
# ---------------------------------------------------------------------------
$byStatus = $records | Group-Object Status | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Count = $_.Count } }
$total = $records.Count
$succ  = ($byStatus | Where-Object Name -eq 'Success' | ForEach-Object Count) ; if (-not $succ) { $succ = 0 }
$warn  = ($byStatus | Where-Object Name -eq 'Warning' | ForEach-Object Count) ; if (-not $warn) { $warn = 0 }
$fail  = ($byStatus | Where-Object Name -eq 'Failed'  | ForEach-Object Count) ; if (-not $fail) { $fail = 0 }
$other = $total - $succ - $warn - $fail
$rate  = if ($total) { [math]::Round(($succ / $total) * 100, 1) } else { 0 }
$managed   = @($records | Where-Object { $_.Deployment -match '(?i)^managed' }).Count
$unmanaged = $total - $managed

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
    [void]$rowsHtml.Append("<td>$(Get-Html $r.Deployment)</td>")
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

$rangeText = "$($StartDate.ToString('yyyy-MM-dd'))  to  $($EndDate.ToString('yyyy-MM-dd'))"
if ($StartTime -or $EndTime) { $rangeText += "   (window $([string]$StartTime)-$([string]$EndTime))" }

$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>VeeamONE SAP HANA Plugin Report</title>
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
  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:14px; margin:20px 0; }
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
      <h1>VeeamONE - SAP HANA (backint) Plugin Report</h1>
      <div class='sub'>Managed + standalone/unmanaged &middot; $rangeText &middot; DB: $(Get-Html $connInfo.Database)</div>
    </div>
    <div class='sub'>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
  </header>

  <div class='cards'>
    <div class='card'><div class='n'>$total</div><div class='l'>Total runs</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Success)'>$succ</div><div class='l'>Success</div></div>
    <div class='card'><div class='n' style='color:$($statusColors.Failed)'>$fail</div><div class='l'>Failed</div></div>
    <div class='card'><div class='n'>$managed</div><div class='l'>Managed</div></div>
    <div class='card'><div class='n'>$unmanaged</div><div class='l'>Unmanaged</div></div>
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
        <th>Job name</th><th>Deployment</th><th>Mode</th><th>Status</th>
        <th>Start</th><th>End</th><th>Elapsed</th><th>Data read</th><th>Transferred</th><th>Reason</th>
      </tr></thead>
      <tbody>
        $($rowsHtml.ToString())
      </tbody>
    </table>
  </div>

  <footer>VeeamONE SAP HANA Plugin Report &middot; source data from the Veeam ONE database (read-only)</footer>
</div>
</body>
</html>
"@

$htmlPath = Join-Path $OutputFolder "VeeamONE-SAPHANA-Report-$stamp.html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "HTML : $htmlPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 8. Optional PDF
# ---------------------------------------------------------------------------
$pdfPath = $null
if ($Pdf) {
    $pdfPath = Join-Path $OutputFolder "VeeamONE-SAPHANA-Report-$stamp.pdf"
    try { ConvertTo-Pdf -HtmlPath $htmlPath -PdfPath $pdfPath | Out-Null; Write-Host "PDF  : $pdfPath" -ForegroundColor Green }
    catch { Write-Warning "PDF generation failed: $($_.Exception.Message)"; $pdfPath = $null }
}

# ---------------------------------------------------------------------------
# 9. Optional email delivery
# ---------------------------------------------------------------------------
if ($EmailTo -and $EmailTo.Count -gt 0) {
    $subject = if ($EmailSubject) { $EmailSubject } else {
        "VeeamONE SAP HANA Report - $rate% success ($total runs) - $($StartDate.ToString('yyyy-MM-dd'))..$($EndDate.ToString('yyyy-MM-dd'))"
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

Write-Host "`nSummary: $total runs | $succ success | $warn warning | $fail failed | $rate% | managed $managed / unmanaged $unmanaged" -ForegroundColor Cyan
[pscustomobject]@{ Csv = $csvPath; Html = $htmlPath; Pdf = $pdfPath; Total = $total; Success = $succ; Warning = $warn; Failed = $fail; Managed = $managed; Unmanaged = $unmanaged; SuccessRate = $rate }
