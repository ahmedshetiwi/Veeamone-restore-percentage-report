<#
.SYNOPSIS
    Create (or remove) a Windows Scheduled Task that runs the VeeamONE Restore
    Report on a recurring basis and e-mails it out.

.DESCRIPTION
    Registers a scheduled task that calls Get-VeeamOneRestoreReport.ps1 with a
    relative reporting window (-Period), optional PDF output (-Pdf) and e-mail
    delivery. Supports Daily, Weekly, Monthly and Yearly cadences via the Task
    Scheduler COM API (so the executable path and arguments never fight with
    command-line quoting).

    UNATTENDED AUTH: the task runs as the Windows account you pass with
    -RunAsCredential, and the report connects to the Veeam ONE database with that
    account (Windows/Integrated auth). Give that account *read* access to the
    Veeam ONE database and permission to relay through your SMTP server. No SQL or
    SMTP password is ever stored in the task - only the Windows run-as password,
    which Windows keeps in the LSA secret store. If -RunAsCredential is omitted the
    task runs only while you are logged on (interactive token).

.PARAMETER Cadence
    Daily | Weekly | Monthly | Yearly.

.PARAMETER Time
    Time of day to run, 24h "HH:mm" (default 07:00).

.PARAMETER DayOfWeek
    Weekly cadence: which day (default Monday).

.PARAMETER DayOfMonth
    Monthly / Yearly cadence: day-of-month 1-31 (default 1).

.PARAMETER Month
    Yearly cadence: which month (default January).

.PARAMETER SqlServer / Database
    Veeam ONE SQL server + database passed to the report.

.PARAMETER EmailTo / EmailFrom / SmtpServer / SmtpPort / SmtpUseSsl
    E-mail delivery settings passed to the report.

.PARAMETER Pdf
    Attach a PDF as well as CSV + HTML.

.PARAMETER RestoreType
    Optional restore-type LIKE filter(s) passed through to the report.

.PARAMETER OutputFolder
    Where the report writes its files (default .\output next to the report).

.PARAMETER RunAsCredential
    Windows account the task runs as (recommended for unattended runs).

.PARAMETER TaskName
    Scheduled task name (default "VeeamONE Restore Report").

.PARAMETER Unregister
    Remove the named task instead of creating it.

.EXAMPLE
    # Every day at 07:00, e-mail a PDF to the ops team
    .\Register-RestoreReportSchedule.ps1 -Cadence Daily -Time 07:00 `
        -SqlServer WIN-U5BC37Q41JJ -Database VeeamONE -Pdf `
        -EmailTo ops@contoso.com -EmailFrom veeam-reports@contoso.com `
        -SmtpServer smtp.contoso.com -RunAsCredential (Get-Credential CONTOSO\svc_veeam)

.EXAMPLE
    # Every Monday at 06:30
    .\Register-RestoreReportSchedule.ps1 -Cadence Weekly -DayOfWeek Monday -Time 06:30 `
        -SqlServer WIN-U5BC37Q41JJ -EmailTo team@contoso.com -SmtpServer smtp.contoso.com `
        -RunAsCredential (Get-Credential CONTOSO\svc_veeam)

.EXAMPLE
    # First day of every month; and a yearly one every 1 Jan
    .\Register-RestoreReportSchedule.ps1 -Cadence Monthly -DayOfMonth 1 ...
    .\Register-RestoreReportSchedule.ps1 -Cadence Yearly -Month January -DayOfMonth 1 ...

.EXAMPLE
    # Remove a schedule
    .\Register-RestoreReportSchedule.ps1 -Unregister -TaskName "VeeamONE Restore Report"
#>
[CmdletBinding(DefaultParameterSetName = 'Create')]
param(
    [Parameter(ParameterSetName = 'Create', Mandatory)]
    [ValidateSet('Daily','Weekly','Monthly','Yearly')]
    [string] $Cadence,

    [ValidatePattern('^\d{1,2}:\d{2}$')]
    [string] $Time = '07:00',

    [ValidateSet('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')]
    [string] $DayOfWeek = 'Monday',

    [ValidateRange(1,31)]
    [int] $DayOfMonth = 1,

    [ValidateSet('January','February','March','April','May','June',
                 'July','August','September','October','November','December')]
    [string] $Month = 'January',

    [string]   $SqlServer,
    [string]   $Database = 'VeeamONE',

    [string[]] $EmailTo,
    [string]   $EmailFrom,
    [string]   $SmtpServer,
    [int]      $SmtpPort = 25,
    [switch]   $SmtpUseSsl,

    [switch]   $Pdf,
    [string[]] $RestoreType,
    [string]   $OutputFolder,

    [pscredential] $RunAsCredential,

    [string] $TaskName = 'VeeamONE Restore Report',

    [Parameter(ParameterSetName = 'Remove', Mandatory)]
    [switch] $Unregister
)

$ErrorActionPreference = 'Stop'

$reportScript = Join-Path $PSScriptRoot 'Get-VeeamOneRestoreReport.ps1'
if (-not $Unregister -and -not (Test-Path $reportScript)) {
    throw "Cannot find Get-VeeamOneRestoreReport.ps1 next to this script ($reportScript)."
}

# --- Connect to Task Scheduler ---
$svc = New-Object -ComObject 'Schedule.Service'
$svc.Connect()
$root = $svc.GetFolder('\')

# --- Remove mode ---
if ($Unregister) {
    try {
        $root.DeleteTask($TaskName, 0)
        Write-Host "Removed scheduled task '$TaskName'." -ForegroundColor Green
    }
    catch { throw "Could not remove task '$TaskName': $($_.Exception.Message)" }
    return
}

# --- Build the report argument string (single string => no quoting battles) ---
$psexe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$sb = [System.Text.StringBuilder]::new()
[void]$sb.Append('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden ')
[void]$sb.Append(('-File "{0}" ' -f $reportScript))
[void]$sb.Append(('-Period {0} ' -f $Cadence))
if ($SqlServer)    { [void]$sb.Append(('-SqlServer "{0}" ' -f $SqlServer)) }
if ($Database)     { [void]$sb.Append(('-Database "{0}" ' -f $Database)) }
if ($Pdf)          { [void]$sb.Append('-Pdf ') }
if ($OutputFolder) { [void]$sb.Append(('-OutputFolder "{0}" ' -f $OutputFolder)) }
if ($RestoreType)  { [void]$sb.Append(('-RestoreType "{0}" ' -f ($RestoreType -join ','))) }
if ($EmailTo)      { [void]$sb.Append(('-EmailTo "{0}" ' -f ($EmailTo -join ','))) }
if ($EmailFrom)    { [void]$sb.Append(('-EmailFrom "{0}" ' -f $EmailFrom)) }
if ($SmtpServer)   { [void]$sb.Append(('-SmtpServer "{0}" ' -f $SmtpServer)) }
if ($SmtpPort)     { [void]$sb.Append(('-SmtpPort {0} ' -f $SmtpPort)) }
if ($SmtpUseSsl)   { [void]$sb.Append('-SmtpUseSsl ') }
$reportArgs = $sb.ToString().Trim()

# --- Define the task ---
$task = $svc.NewTask(0)
$task.RegistrationInfo.Description = "Runs the VeeamONE Restore Report ($Cadence) and e-mails it."
$task.RegistrationInfo.Author      = "$env:USERDOMAIN\$env:USERNAME"
$task.Settings.Enabled             = $true
$task.Settings.StartWhenAvailable  = $true   # catch up if the machine was off
$task.Settings.ExecutionTimeLimit  = 'PT1H'

# Trigger. COM trigger type ids: 2=Daily, 3=Weekly, 4=Monthly (by day-of-month).
$parts  = $Time.Split(':')
$hh     = [int]$parts[0]; $mm = [int]$parts[1]
$startBoundary = (Get-Date -Hour $hh -Minute $mm -Second 0).ToString('yyyy-MM-ddTHH:mm:ss')

switch ($Cadence) {
    'Daily' {
        $t = $task.Triggers.Create(2)
        $t.DaysInterval = 1
    }
    'Weekly' {
        $t = $task.Triggers.Create(3)
        $t.WeeksInterval = 1
        $dowBit = @{ Sunday=1; Monday=2; Tuesday=4; Wednesday=8; Thursday=16; Friday=32; Saturday=64 }
        $t.DaysOfWeek = $dowBit[$DayOfWeek]
    }
    'Monthly' {
        $t = $task.Triggers.Create(4)
        $t.DaysOfMonth  = [int][math]::Pow(2, $DayOfMonth - 1)   # bit (day-1)
        $t.MonthsOfYear = 4095                                   # all 12 months
    }
    'Yearly' {
        $t = $task.Triggers.Create(4)
        $t.DaysOfMonth = [int][math]::Pow(2, $DayOfMonth - 1)
        $monthNum = [datetime]::ParseExact($Month,'MMMM',[Globalization.CultureInfo]::InvariantCulture).Month
        $t.MonthsOfYear = [int][math]::Pow(2, $monthNum - 1)     # single month => once a year
    }
}
$t.StartBoundary = $startBoundary
$t.Enabled       = $true

# Action
$action = $task.Actions.Create(0)   # 0 = TASK_ACTION_EXEC
$action.Path      = $psexe
$action.Arguments = $reportArgs
$action.WorkingDirectory = $PSScriptRoot

# --- Register ---
# LogonType: 1 = password (runs whether the user is logged on or not),
#            3 = interactive token (runs only while that user is logged on).
if ($RunAsCredential) {
    $user     = $RunAsCredential.UserName
    $password = $RunAsCredential.GetNetworkCredential().Password
    $logon    = 1
    $task.Principal.RunLevel = 1   # TASK_RUNLEVEL_HIGHEST
}
else {
    $user     = "$env:USERDOMAIN\$env:USERNAME"
    $password = $null
    $logon    = 3
    Write-Warning "No -RunAsCredential given: the task will run only while '$user' is logged on. For unattended daily/weekly runs, pass -RunAsCredential for a service account with DB read + SMTP relay access."
}

# 6 = TASK_CREATE_OR_UPDATE
[void]$root.RegisterTaskDefinition($TaskName, $task, 6, $user, $password, $logon)

Write-Host "Scheduled task '$TaskName' registered." -ForegroundColor Green
Write-Host "  Cadence : $Cadence at $Time$(if($Cadence -eq 'Weekly'){" on $DayOfWeek"}elseif($Cadence -eq 'Monthly'){" on day $DayOfMonth"}elseif($Cadence -eq 'Yearly'){" on $Month $DayOfMonth"})" -ForegroundColor Cyan
Write-Host "  Runs as : $user (logon type $logon)" -ForegroundColor Cyan
Write-Host "  Command : $psexe $reportArgs" -ForegroundColor DarkGray
Write-Host "`nManage it in Task Scheduler, or run it now with:  schtasks /Run /TN `"$TaskName`"" -ForegroundColor Gray
