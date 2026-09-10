<#
.SYNOPSIS
    GUI front-end for Get-VeeamOneRestoreReport.ps1.

    Three tabs:
      - Report   : DB connection, auth, date range, restore-type filter, output
                   folder, PDF option, and Run.
      - Email     : optionally e-mail the report (CSV + HTML + PDF) on each run.
      - Schedule  : create/remove a Windows Scheduled Task that runs the report
                   Daily / Weekly / Monthly / Yearly and e-mails it.

    Produces the same CSV + HTML (+ optional PDF) restore success/failure report
    as the command-line script.

.NOTES
    Run on (or next to) the Veeam ONE server:
        powershell -ExecutionPolicy Bypass -STA -File .\Show-RestoreReportForm.ps1
    (Windows PowerShell 5.1. -STA is the default for powershell.exe.)
#>
[CmdletBinding()]
param(
    # Build the form and exit without showing it (smoke test - no display needed).
    [switch] $SelfTest
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportTool = Join-Path $scriptDir 'Get-VeeamOneRestoreReport.ps1'
$schedTool  = Join-Path $scriptDir 'Register-RestoreReportSchedule.ps1'
# Engines for the three added report tabs (Backup Jobs / RMAN / All Backups).
$jobsTool   = Join-Path $scriptDir 'Get-VeeamOneBackupJobReport.ps1'
$rmanTool   = Join-Path $scriptDir 'Get-VeeamOneRmanBackupReport.ps1'
$allTool    = Join-Path $scriptDir 'Get-VeeamOneAllBackupsReport.ps1'
$sapTool    = Join-Path $scriptDir 'Get-VeeamOneSapHanaReport.ps1'
if (-not (Test-Path $reportTool)) {
    [System.Windows.Forms.MessageBox]::Show("Cannot find Get-VeeamOneRestoreReport.ps1 next to this GUI.`n`nExpected: $reportTool",
        'Veeam ONE Restore Report', 'OK', 'Error') | Out-Null
    return
}

# ---- palette ----
$clrBg     = [System.Drawing.Color]::FromArgb(244,246,248)
$clrGreen  = [System.Drawing.Color]::FromArgb(0,179,54)
$clrText   = [System.Drawing.Color]::FromArgb(26,39,51)
$fontLbl   = New-Object System.Drawing.Font('Segoe UI', 9)
$fontHead  = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)

# ---- form ----
$form                = New-Object System.Windows.Forms.Form
$form.Text           = 'VeeamOne Custom Reports'
$form.Size           = New-Object System.Drawing.Size(600, 758)
$form.StartPosition  = 'CenterScreen'
$form.FormBorderStyle= 'FixedDialog'
$form.MaximizeBox    = $false
$form.BackColor      = $clrBg
$form.Font           = $fontLbl

# header
$hdr = New-Object System.Windows.Forms.Label
$hdr.Text = 'VeeamOne Custom Reports'
$hdr.Font = $fontHead; $hdr.ForeColor = $clrText
$hdr.Location = New-Object System.Drawing.Point(18, 12)
$hdr.Size = New-Object System.Drawing.Size(560, 30)
$form.Controls.Add($hdr)

# ---- generic control helpers (parent-aware) ----
function New-Label {
    param($parent, $text, $x, $y, $w = 160)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.Size = New-Object System.Drawing.Size($w, 20); $l.ForeColor = $clrText
    $parent.Controls.Add($l); $l
}
function New-Text {
    param($parent, $x, $y, $w = 330, $val = '')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y)
    $t.Size = New-Object System.Drawing.Size($w, 22); $t.Text = $val
    $parent.Controls.Add($t); $t
}
function New-DateTimePicker {
    param($parent, $x, $y, $val)
    $dtp = New-Object System.Windows.Forms.DateTimePicker
    $dtp.Format = 'Custom'; $dtp.CustomFormat = 'yyyy-MM-dd  HH:mm'
    $dtp.ShowUpDown = $true
    $dtp.Location = New-Object System.Drawing.Point($x, $y)
    $dtp.Size = New-Object System.Drawing.Size(160, 22)
    $dtp.Value = $val
    $parent.Controls.Add($dtp); $dtp
}
# Adds "Start (date & time)" + "End (date & time)" pickers to a tab; returns [start,end].
function Add-DateRange {
    param($parent, [int]$y, $defStart = ((Get-Date).Date.AddDays(-30)), $defEnd = (Get-Date))
    New-Label $parent 'Start (date & time)' 16 ($y+2) | Out-Null
    $s = New-DateTimePicker $parent 180 $y $defStart
    New-Label $parent 'End (date & time)' 16 ($y+34) | Out-Null
    $e = New-DateTimePicker $parent 180 ($y+32) $defEnd
    ,@($s, $e)
}

# ---- tab control ----
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(16, 50)
$tabs.Size     = New-Object System.Drawing.Size(560, 560)
$tabs.Multiline = $true   # wrap the 6 tabs onto two rows so all stay visible
$form.Controls.Add($tabs)

$tabConn     = New-Object System.Windows.Forms.TabPage; $tabConn.Text     = '  Connection  '
$tabReport   = New-Object System.Windows.Forms.TabPage; $tabReport.Text   = '  Restoration report  '
$tabEmail    = New-Object System.Windows.Forms.TabPage; $tabEmail.Text    = '  Email  '
$tabSchedule = New-Object System.Windows.Forms.TabPage; $tabSchedule.Text = '  Schedule  '
$tabJobs     = New-Object System.Windows.Forms.TabPage; $tabJobs.Text     = '  Backup Jobs  '
$tabRman     = New-Object System.Windows.Forms.TabPage; $tabRman.Text     = '  RMAN Backups  '
$tabAll      = New-Object System.Windows.Forms.TabPage; $tabAll.Text      = '  All Backups  '
$tabSap      = New-Object System.Windows.Forms.TabPage; $tabSap.Text      = '  SAP HANA  '
# Order: Connection (shared DB/auth) first, then the reports, then Email + Schedule.
foreach ($tp in $tabConn, $tabReport, $tabJobs, $tabAll, $tabRman, $tabSap, $tabEmail, $tabSchedule) { $tp.BackColor = [System.Drawing.Color]::White; $tabs.TabPages.Add($tp) }

$col2 = 180     # x for input controls inside a tab

# =====================================================================
# TAB - CONNECTION (shared DB + authentication + output, used by ALL reports)
# =====================================================================
$cy = 16
$lblConnIntro = New-Object System.Windows.Forms.Label
$lblConnIntro.Text = 'Database connection & authentication used by EVERY report tab. Set it once here.'
$lblConnIntro.ForeColor = [System.Drawing.Color]::Gray
$lblConnIntro.Location = New-Object System.Drawing.Point(16, $cy)
$lblConnIntro.Size = New-Object System.Drawing.Size(516, 18)
$tabConn.Controls.Add($lblConnIntro)
$cy += 28
New-Label $tabConn 'SQL Server (IP or name)' 16 ($cy+2) | Out-Null
$txtServer = New-Text $tabConn $col2 $cy 250
$cy += 32
New-Label $tabConn 'Database' 16 ($cy+2) | Out-Null
$txtDb = New-Text $tabConn $col2 $cy 250 'VeeamONE'
$cy += 38

# Authentication group
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = 'Authentication'
$grp.Location = New-Object System.Drawing.Point(16, $cy)
$grp.Size = New-Object System.Drawing.Size(516, 150)
$grp.ForeColor = $clrText
$tabConn.Controls.Add($grp)

$rbDomain = New-Object System.Windows.Forms.RadioButton
$rbDomain.Text = 'Windows - domain account (username + password)'
$rbDomain.Location = New-Object System.Drawing.Point(14, 22)
$rbDomain.Size = New-Object System.Drawing.Size(420, 20)
$rbDomain.Checked = $true
$grp.Controls.Add($rbDomain)

$rbSql = New-Object System.Windows.Forms.RadioButton
$rbSql.Text = 'SQL Server login'
$rbSql.Location = New-Object System.Drawing.Point(14, 46)
$rbSql.Size = New-Object System.Drawing.Size(200, 20)
$grp.Controls.Add($rbSql)

$rbCurrent = New-Object System.Windows.Forms.RadioButton
$rbCurrent.Text = 'Windows - current user'
$rbCurrent.Location = New-Object System.Drawing.Point(230, 46)
$rbCurrent.Size = New-Object System.Drawing.Size(250, 20)
$grp.Controls.Add($rbCurrent)

$lblUserHint = New-Object System.Windows.Forms.Label
$lblUserHint.Text = 'e.g. CONTOSO\svc_veeam  (or user@contoso.com)'
$lblUserHint.ForeColor = [System.Drawing.Color]::Gray
$lblUserHint.Location = New-Object System.Drawing.Point(170, 60)
$lblUserHint.Size = New-Object System.Drawing.Size(330, 16)
$grp.Controls.Add($lblUserHint)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = 'Username'
$lblUser.Location = New-Object System.Drawing.Point(14, 82)
$lblUser.Size = New-Object System.Drawing.Size(150, 20)
$grp.Controls.Add($lblUser)
$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(170, 79)
$txtUser.Size = New-Object System.Drawing.Size(326, 22)
$grp.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = 'Password'
$lblPass.Location = New-Object System.Drawing.Point(14, 112)
$lblPass.Size = New-Object System.Drawing.Size(150, 20)
$grp.Controls.Add($lblPass)
$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(170, 109)
$txtPass.Size = New-Object System.Drawing.Size(326, 22)
$txtPass.UseSystemPasswordChar = $true
$grp.Controls.Add($txtPass)

$lblReadOnly = New-Object System.Windows.Forms.Label
$lblReadOnly.Text = 'Tip: use a read-only login (db_datareader). The tool only runs SELECT queries.'
$lblReadOnly.ForeColor = [System.Drawing.Color]::Gray
$lblReadOnly.Location = New-Object System.Drawing.Point(14, 132)
$lblReadOnly.Size = New-Object System.Drawing.Size(494, 16)
$grp.Controls.Add($lblReadOnly)

$cy += 162
New-Label $tabConn 'Output folder' 16 ($cy+2) | Out-Null
$txtOut = New-Text $tabConn $col2 $cy 250 (Join-Path $scriptDir 'output')
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '...'
$btnBrowse.Location = New-Object System.Drawing.Point(437, ($cy-1))
$btnBrowse.Size = New-Object System.Drawing.Size(30, 24)
$tabConn.Controls.Add($btnBrowse)
$cy += 32
$chkPdf = New-Object System.Windows.Forms.CheckBox
$chkPdf.Text = 'Also generate a PDF (uses Edge/Chrome) - applies to all reports'
$chkPdf.Location = New-Object System.Drawing.Point($col2, $cy)
$chkPdf.Size = New-Object System.Drawing.Size(360, 20)
$tabConn.Controls.Add($chkPdf)

# =====================================================================
# TAB - RESTORATION REPORT (filter + own date range; runs via bottom button)
# =====================================================================
$y = 16
$lblRepIntro = New-Object System.Windows.Forms.Label
$lblRepIntro.Text = 'Restore success / warning / failure rate from the Veeam ONE database. Connection is on the Connection tab; Email/Schedule tabs apply.'
$lblRepIntro.ForeColor = [System.Drawing.Color]::Gray
$lblRepIntro.Location = New-Object System.Drawing.Point(16, $y)
$lblRepIntro.Size = New-Object System.Drawing.Size(516, 32)
$tabReport.Controls.Add($lblRepIntro)
$y += 40
New-Label $tabReport 'Restore type filter' 16 ($y+2) | Out-Null
$txtRestoreType = New-Text $tabReport $col2 $y 150
$btnDiscover = New-Object System.Windows.Forms.Button
$btnDiscover.Text = 'Discover...'
$btnDiscover.Location = New-Object System.Drawing.Point(340, ($y-1))
$btnDiscover.Size = New-Object System.Drawing.Size(100, 24)
$tabReport.Controls.Add($btnDiscover)
$y += 30
$hint2 = New-Object System.Windows.Forms.Label
$hint2.Text = 'Optional LIKE patterns, comma-separated. Blank = all.'
$hint2.ForeColor = [System.Drawing.Color]::Gray
$hint2.Location = New-Object System.Drawing.Point($col2, $y)
$hint2.Size = New-Object System.Drawing.Size(350, 16)
$tabReport.Controls.Add($hint2)
$y += 28
$range = Add-DateRange $tabReport $y
$dtpStart = $range[0]; $dtpEnd = $range[1]
$y += 74
$chkRetry = New-Object System.Windows.Forms.CheckBox
$chkRetry.Text = 'Count a failed restore as success if it succeeded on retry within'
$chkRetry.Location = New-Object System.Drawing.Point(16, $y)
$chkRetry.Size = New-Object System.Drawing.Size(400, 20)
$chkRetry.Checked = $true
$tabReport.Controls.Add($chkRetry)
$numRetryHrs = New-Object System.Windows.Forms.NumericUpDown
$numRetryHrs.Location = New-Object System.Drawing.Point(420, ($y-1))
$numRetryHrs.Size = New-Object System.Drawing.Size(50, 22)
$numRetryHrs.Minimum = 1; $numRetryHrs.Maximum = 168; $numRetryHrs.Value = 6
$tabReport.Controls.Add($numRetryHrs)
New-Label $tabReport 'hours' 474 ($y+2) 50 | Out-Null
$y += 28
$chkByJob = New-Object System.Windows.Forms.CheckBox
$chkByJob.Text = 'Count by restore job (session), not per restored object'
$chkByJob.Location = New-Object System.Drawing.Point(16, $y)
$chkByJob.Size = New-Object System.Drawing.Size(460, 20)
$tabReport.Controls.Add($chkByJob)

# =====================================================================
# TAB 2 - EMAIL
# =====================================================================
$ey = 18
$chkEmail = New-Object System.Windows.Forms.CheckBox
$chkEmail.Text = 'E-mail the report each time it runs'
$chkEmail.Location = New-Object System.Drawing.Point(16, $ey)
$chkEmail.Size = New-Object System.Drawing.Size(360, 20)
$tabEmail.Controls.Add($chkEmail)
$ey += 34

New-Label $tabEmail 'To (comma-separated)' 16 ($ey+2) | Out-Null
$txtTo = New-Text $tabEmail $col2 $ey 330
$ey += 32
New-Label $tabEmail 'From' 16 ($ey+2) | Out-Null
$txtFrom = New-Text $tabEmail $col2 $ey 330
$ey += 32
New-Label $tabEmail 'SMTP server' 16 ($ey+2) | Out-Null
$txtSmtp = New-Text $tabEmail $col2 $ey 250
$ey += 32
New-Label $tabEmail 'SMTP port' 16 ($ey+2) | Out-Null
$txtPort = New-Text $tabEmail $col2 $ey 80 '25'
$chkSsl = New-Object System.Windows.Forms.CheckBox
$chkSsl.Text = 'Use SSL/TLS'
$chkSsl.Location = New-Object System.Drawing.Point(280, ($ey+1))
$chkSsl.Size = New-Object System.Drawing.Size(120, 20)
$tabEmail.Controls.Add($chkSsl)
$ey += 38

$grpAuth = New-Object System.Windows.Forms.GroupBox
$grpAuth.Text = 'SMTP authentication (optional - blank = anonymous relay)'
$grpAuth.Location = New-Object System.Drawing.Point(16, $ey)
$grpAuth.Size = New-Object System.Drawing.Size(516, 96)
$grpAuth.ForeColor = $clrText
$tabEmail.Controls.Add($grpAuth)
$l = New-Object System.Windows.Forms.Label; $l.Text='Username'; $l.Location=New-Object System.Drawing.Point(14,28); $l.Size=New-Object System.Drawing.Size(150,20); $grpAuth.Controls.Add($l)
$txtSmtpUser = New-Object System.Windows.Forms.TextBox; $txtSmtpUser.Location=New-Object System.Drawing.Point(164,25); $txtSmtpUser.Size=New-Object System.Drawing.Size(332,22); $grpAuth.Controls.Add($txtSmtpUser)
$l = New-Object System.Windows.Forms.Label; $l.Text='Password'; $l.Location=New-Object System.Drawing.Point(14,58); $l.Size=New-Object System.Drawing.Size(150,20); $grpAuth.Controls.Add($l)
$txtSmtpPass = New-Object System.Windows.Forms.TextBox; $txtSmtpPass.Location=New-Object System.Drawing.Point(164,55); $txtSmtpPass.Size=New-Object System.Drawing.Size(332,22); $txtSmtpPass.UseSystemPasswordChar=$true; $grpAuth.Controls.Add($txtSmtpPass)
$ey += 108

$lblEmailHint = New-Object System.Windows.Forms.Label
$lblEmailHint.Text = 'Two ways to e-mail: (1) tick the box above, then run any report - it is attached automatically; or (2) pick a report below and e-mail it right now.'
$lblEmailHint.ForeColor = [System.Drawing.Color]::Gray
$lblEmailHint.Location = New-Object System.Drawing.Point(16, $ey)
$lblEmailHint.Size = New-Object System.Drawing.Size(516, 44)
$tabEmail.Controls.Add($lblEmailHint)
$ey += 52

# Pick a specific report + e-mail it now (uses that report tab's filters/date range).
New-Label $tabEmail 'Report to e-mail' 16 ($ey+2) | Out-Null
$cboEmailReport = New-Object System.Windows.Forms.ComboBox
$cboEmailReport.DropDownStyle = 'DropDownList'
$cboEmailReport.Location = New-Object System.Drawing.Point($col2, $ey)
$cboEmailReport.Size = New-Object System.Drawing.Size(220, 22)
[void]$cboEmailReport.Items.AddRange(@('Restore','Backup Jobs','All Backups','RMAN Backups','SAP HANA'))
$cboEmailReport.SelectedIndex = 0
$tabEmail.Controls.Add($cboEmailReport)
$ey += 34
$btnEmailNow = New-Object System.Windows.Forms.Button
$btnEmailNow.Text = 'Generate & e-mail now'
$btnEmailNow.Location = New-Object System.Drawing.Point($col2, $ey)
$btnEmailNow.Size = New-Object System.Drawing.Size(200, 30)
$btnEmailNow.BackColor = $clrGreen; $btnEmailNow.ForeColor = [System.Drawing.Color]::White; $btnEmailNow.FlatStyle = 'Flat'
$tabEmail.Controls.Add($btnEmailNow)
$ey += 36
$lblEmailNowHint = New-Object System.Windows.Forms.Label
$lblEmailNowHint.Text = 'The chosen report uses the filters + date range set on its own tab.'
$lblEmailNowHint.ForeColor = [System.Drawing.Color]::Gray
$lblEmailNowHint.Location = New-Object System.Drawing.Point(16, $ey)
$lblEmailNowHint.Size = New-Object System.Drawing.Size(516, 20)
$tabEmail.Controls.Add($lblEmailNowHint)

# =====================================================================
# TAB 3 - SCHEDULE
# =====================================================================
$sy = 18
$lblSchedIntro = New-Object System.Windows.Forms.Label
$lblSchedIntro.Text = 'Create a Windows Scheduled Task that runs this report automatically and e-mails it. Uses the Report + Email tab settings above. Scheduled runs use Windows (Integrated) auth as the run-as account below.'
$lblSchedIntro.ForeColor = [System.Drawing.Color]::Gray
$lblSchedIntro.Location = New-Object System.Drawing.Point(16, $sy)
$lblSchedIntro.Size = New-Object System.Drawing.Size(516, 46)
$tabSchedule.Controls.Add($lblSchedIntro)
$sy += 52

# Which report the schedule (Create/Remove below) targets - the same scheduler
# and Email-tab SMTP settings drive every report.
New-Label $tabSchedule 'Report to schedule' 16 ($sy+2) | Out-Null
$cboSchedReport = New-Object System.Windows.Forms.ComboBox
$cboSchedReport.DropDownStyle = 'DropDownList'
$cboSchedReport.Location = New-Object System.Drawing.Point($col2, $sy)
$cboSchedReport.Size = New-Object System.Drawing.Size(220, 22)
[void]$cboSchedReport.Items.AddRange(@('Restore','Backup Jobs','All Backups','RMAN Backups','SAP HANA'))
$cboSchedReport.SelectedIndex = 0
$tabSchedule.Controls.Add($cboSchedReport)
$sy += 32

New-Label $tabSchedule 'Filter (optional)' 16 ($sy+2) | Out-Null
$txtSchedFilter = New-Text $tabSchedule $col2 $sy 220
$sy += 26
$lblSchedFilterHint = New-Object System.Windows.Forms.Label
$lblSchedFilterHint.Text = 'LIKE filter passed to the chosen report (Restore=item type, RMAN/Jobs/All=job name). Blank = no filter.'
$lblSchedFilterHint.ForeColor = [System.Drawing.Color]::Gray
$lblSchedFilterHint.Location = New-Object System.Drawing.Point($col2, $sy)
$lblSchedFilterHint.Size = New-Object System.Drawing.Size(340, 28)
$tabSchedule.Controls.Add($lblSchedFilterHint)
$sy += 34

New-Label $tabSchedule 'Frequency' 16 ($sy+2) | Out-Null
$cboCadence = New-Object System.Windows.Forms.ComboBox
$cboCadence.DropDownStyle = 'DropDownList'
$cboCadence.Location = New-Object System.Drawing.Point($col2, $sy)
$cboCadence.Size = New-Object System.Drawing.Size(150, 22)
[void]$cboCadence.Items.AddRange(@('Daily','Weekly','Monthly','Yearly'))
$cboCadence.SelectedIndex = 0
$tabSchedule.Controls.Add($cboCadence)
$sy += 32

New-Label $tabSchedule 'Time (HH:mm)' 16 ($sy+2) | Out-Null
$txtSchedTime = New-Text $tabSchedule $col2 $sy 80 '07:00'
$sy += 32

$lblDow = New-Label $tabSchedule 'Day of week' 16 ($sy+2)
$cboDow = New-Object System.Windows.Forms.ComboBox
$cboDow.DropDownStyle = 'DropDownList'
$cboDow.Location = New-Object System.Drawing.Point($col2, $sy)
$cboDow.Size = New-Object System.Drawing.Size(150, 22)
[void]$cboDow.Items.AddRange(@('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'))
$cboDow.SelectedIndex = 0
$tabSchedule.Controls.Add($cboDow)
$sy += 32

$lblDom = New-Label $tabSchedule 'Day of month' 16 ($sy+2)
$numDom = New-Object System.Windows.Forms.NumericUpDown
$numDom.Location = New-Object System.Drawing.Point($col2, $sy)
$numDom.Size = New-Object System.Drawing.Size(80, 22)
$numDom.Minimum = 1; $numDom.Maximum = 31; $numDom.Value = 1
$tabSchedule.Controls.Add($numDom)
$sy += 32

$lblMonth = New-Label $tabSchedule 'Month' 16 ($sy+2)
$cboMonth = New-Object System.Windows.Forms.ComboBox
$cboMonth.DropDownStyle = 'DropDownList'
$cboMonth.Location = New-Object System.Drawing.Point($col2, $sy)
$cboMonth.Size = New-Object System.Drawing.Size(150, 22)
[void]$cboMonth.Items.AddRange(@('January','February','March','April','May','June','July','August','September','October','November','December'))
$cboMonth.SelectedIndex = 0
$tabSchedule.Controls.Add($cboMonth)
$sy += 38

$grpRun = New-Object System.Windows.Forms.GroupBox
$grpRun.Text = 'Run-as account (recommended for unattended runs; blank = only when you are logged on)'
$grpRun.Location = New-Object System.Drawing.Point(16, $sy)
$grpRun.Size = New-Object System.Drawing.Size(516, 96)
$grpRun.ForeColor = $clrText
$tabSchedule.Controls.Add($grpRun)
$l = New-Object System.Windows.Forms.Label; $l.Text='Username'; $l.Location=New-Object System.Drawing.Point(14,28); $l.Size=New-Object System.Drawing.Size(150,20); $grpRun.Controls.Add($l)
$txtRunUser = New-Object System.Windows.Forms.TextBox; $txtRunUser.Location=New-Object System.Drawing.Point(164,25); $txtRunUser.Size=New-Object System.Drawing.Size(332,22); $grpRun.Controls.Add($txtRunUser)
$l = New-Object System.Windows.Forms.Label; $l.Text='Password'; $l.Location=New-Object System.Drawing.Point(14,58); $l.Size=New-Object System.Drawing.Size(150,20); $grpRun.Controls.Add($l)
$txtRunPass = New-Object System.Windows.Forms.TextBox; $txtRunPass.Location=New-Object System.Drawing.Point(164,55); $txtRunPass.Size=New-Object System.Drawing.Size(332,22); $txtRunPass.UseSystemPasswordChar=$true; $grpRun.Controls.Add($txtRunPass)
$sy += 104

$btnSchedule = New-Object System.Windows.Forms.Button
$btnSchedule.Text = 'Create schedule'
$btnSchedule.Location = New-Object System.Drawing.Point(16, $sy)
$btnSchedule.Size = New-Object System.Drawing.Size(140, 28)
$btnSchedule.BackColor = $clrGreen; $btnSchedule.ForeColor = [System.Drawing.Color]::White
$btnSchedule.FlatStyle = 'Flat'
$tabSchedule.Controls.Add($btnSchedule)

$btnUnschedule = New-Object System.Windows.Forms.Button
$btnUnschedule.Text = 'Remove schedule'
$btnUnschedule.Location = New-Object System.Drawing.Point(166, $sy)
$btnUnschedule.Size = New-Object System.Drawing.Size(140, 28)
$tabSchedule.Controls.Add($btnUnschedule)

# =====================================================================
# TAB 4 - BACKUP JOBS (inventory: schedule, type, last run, by policy)
# =====================================================================
$jy = 16
$lblJobsIntro = New-Object System.Windows.Forms.Label
$lblJobsIntro.Text = 'All backup jobs with configured schedule, type details, retention, last run and status - grouped by backup policy. Uses the Connection tab for DB/auth, and the Email tab for delivery.'
$lblJobsIntro.ForeColor = [System.Drawing.Color]::Gray
$lblJobsIntro.Location = New-Object System.Drawing.Point(16, $jy)
$lblJobsIntro.Size = New-Object System.Drawing.Size(516, 46)
$tabJobs.Controls.Add($lblJobsIntro)
$jy += 54
New-Label $tabJobs 'Job name filter (LIKE)' 16 ($jy+2) | Out-Null
$txtJobsName = New-Text $tabJobs $col2 $jy 250
$jy += 32
New-Label $tabJobs 'Policy filter (LIKE)' 16 ($jy+2) | Out-Null
$txtJobsPolicy = New-Text $tabJobs $col2 $jy 250
$jy += 32
$chkJobsEnabled = New-Object System.Windows.Forms.CheckBox
$chkJobsEnabled.Text = 'Enabled (scheduled) jobs only'
$chkJobsEnabled.Location = New-Object System.Drawing.Point($col2, $jy)
$chkJobsEnabled.Size = New-Object System.Drawing.Size(300, 20)
$tabJobs.Controls.Add($chkJobsEnabled)
$jy += 26
$chkJobsCalendar = New-Object System.Windows.Forms.CheckBox
$chkJobsCalendar.Text = 'Also generate A5 backup calendar (schedule & time to run)'
$chkJobsCalendar.Location = New-Object System.Drawing.Point($col2, $jy)
$chkJobsCalendar.Size = New-Object System.Drawing.Size(360, 20)
$tabJobs.Controls.Add($chkJobsCalendar)
$jy += 34
$lblJobsHint = New-Object System.Windows.Forms.Label
$lblJobsHint.Text = 'Tip: filter to a policy or "%RMAN%" to focus. Blank = every job. (Job inventory is current state; the range labels the report.)'
$lblJobsHint.ForeColor = [System.Drawing.Color]::Gray
$lblJobsHint.Location = New-Object System.Drawing.Point(16, $jy)
$lblJobsHint.Size = New-Object System.Drawing.Size(516, 30)
$tabJobs.Controls.Add($lblJobsHint)
$jy += 34
$rangeJobs = Add-DateRange $tabJobs $jy; $dtpJobsStart = $rangeJobs[0]; $dtpJobsEnd = $rangeJobs[1]
$jy += 70
$btnJobsDiscover = New-Object System.Windows.Forms.Button
$btnJobsDiscover.Text = 'Discover types...'
$btnJobsDiscover.Location = New-Object System.Drawing.Point(16, $jy)
$btnJobsDiscover.Size = New-Object System.Drawing.Size(130, 28)
$tabJobs.Controls.Add($btnJobsDiscover)
$lblJobsRunHint = New-Object System.Windows.Forms.Label
$lblJobsRunHint.Text = 'Then click "Run selected report" at the bottom.'
$lblJobsRunHint.ForeColor = [System.Drawing.Color]::Gray
$lblJobsRunHint.Location = New-Object System.Drawing.Point(160, ($jy+6))
$lblJobsRunHint.Size = New-Object System.Drawing.Size(360, 20)
$tabJobs.Controls.Add($lblJobsRunHint)

# =====================================================================
# TAB 5 - RMAN BACKUPS (standalone/unmanaged rescans - success rate)
# =====================================================================
$ry = 16
$lblRmanIntro = New-Object System.Windows.Forms.Label
$lblRmanIntro.Text = 'RMAN plugin (standalone / unmanaged) backup rescans - success / warning / failure rate, plus a per-database summary. Uses the Connection tab for DB/auth and the Email tab for delivery.'
$lblRmanIntro.ForeColor = [System.Drawing.Color]::Gray
$lblRmanIntro.Location = New-Object System.Drawing.Point(16, $ry)
$lblRmanIntro.Size = New-Object System.Drawing.Size(516, 46)
$tabRman.Controls.Add($lblRmanIntro)
$ry += 54
New-Label $tabRman 'Job name filter (LIKE)' 16 ($ry+2) | Out-Null
$txtRmanName = New-Text $tabRman $col2 $ry 250 '%RMAN%'
$ry += 32
$chkRmanAll = New-Object System.Windows.Forms.CheckBox
$chkRmanAll.Text = 'Include ALL backups (ignore the job-name filter)'
$chkRmanAll.Location = New-Object System.Drawing.Point($col2, $ry)
$chkRmanAll.Size = New-Object System.Drawing.Size(340, 20)
$tabRman.Controls.Add($chkRmanAll)
$ry += 34
$lblRmanHint = New-Object System.Windows.Forms.Label
$lblRmanHint.Text = 'Default keeps RMAN plugin jobs (job name LIKE %RMAN%). Uses the date range below.'
$lblRmanHint.ForeColor = [System.Drawing.Color]::Gray
$lblRmanHint.Location = New-Object System.Drawing.Point(16, $ry)
$lblRmanHint.Size = New-Object System.Drawing.Size(516, 30)
$tabRman.Controls.Add($lblRmanHint)
$ry += 34
$rangeRman = Add-DateRange $tabRman $ry; $dtpRmanStart = $rangeRman[0]; $dtpRmanEnd = $rangeRman[1]
$ry += 74
$lblRmanRunHint = New-Object System.Windows.Forms.Label
$lblRmanRunHint.Text = 'Click "Run selected report" at the bottom to generate this report.'
$lblRmanRunHint.ForeColor = [System.Drawing.Color]::Gray
$lblRmanRunHint.Location = New-Object System.Drawing.Point(16, $ry)
$lblRmanRunHint.Size = New-Object System.Drawing.Size(500, 20)
$tabRman.Controls.Add($lblRmanRunHint)

# =====================================================================
# TAB 6 - ALL BACKUPS (every backup run, rich detail + failure reasons)
# =====================================================================
$ay = 16
$lblAllIntro = New-Object System.Windows.Forms.Label
$lblAllIntro.Text = 'Every backup job run with status, type, mode, timing, data sizes, processing rate and failure reasons. Uses the Connection tab for DB/auth and the Email tab for delivery.'
$lblAllIntro.ForeColor = [System.Drawing.Color]::Gray
$lblAllIntro.Location = New-Object System.Drawing.Point(16, $ay)
$lblAllIntro.Size = New-Object System.Drawing.Size(516, 46)
$tabAll.Controls.Add($lblAllIntro)
$ay += 54
New-Label $tabAll 'Job name filter (LIKE)' 16 ($ay+2) | Out-Null
$txtAllName = New-Text $tabAll $col2 $ay 250
$ay += 32
New-Label $tabAll 'Status filter' 16 ($ay+2) | Out-Null
$cboAllStatus = New-Object System.Windows.Forms.ComboBox
$cboAllStatus.DropDownStyle = 'DropDownList'
$cboAllStatus.Location = New-Object System.Drawing.Point($col2, $ay)
$cboAllStatus.Size = New-Object System.Drawing.Size(160, 22)
[void]$cboAllStatus.Items.AddRange(@('All','Success','Warning','Failed'))
$cboAllStatus.SelectedIndex = 0
$tabAll.Controls.Add($cboAllStatus)
$ay += 32
$lblAllHint = New-Object System.Windows.Forms.Label
$lblAllHint.Text = 'Blank name = every backup. Uses the date range below.'
$lblAllHint.ForeColor = [System.Drawing.Color]::Gray
$lblAllHint.Location = New-Object System.Drawing.Point(16, $ay)
$lblAllHint.Size = New-Object System.Drawing.Size(516, 30)
$tabAll.Controls.Add($lblAllHint)
$ay += 34
$rangeAll = Add-DateRange $tabAll $ay; $dtpAllStart = $rangeAll[0]; $dtpAllEnd = $rangeAll[1]
$ay += 74
$lblAllRunHint = New-Object System.Windows.Forms.Label
$lblAllRunHint.Text = 'Click "Run selected report" at the bottom to generate this report.'
$lblAllRunHint.ForeColor = [System.Drawing.Color]::Gray
$lblAllRunHint.Location = New-Object System.Drawing.Point(16, $ay)
$lblAllRunHint.Size = New-Object System.Drawing.Size(500, 20)
$tabAll.Controls.Add($lblAllRunHint)

# =====================================================================
# TAB 7 - SAP HANA (backint plugin: managed + standalone/unmanaged)
# =====================================================================
$sy2 = 16
$lblSapIntro = New-Object System.Windows.Forms.Label
$lblSapIntro.Text = 'SAP HANA (backint) plugin backups - success / failure rate, covering BOTH plugin modes: VBR-managed (application policy) and standalone / unmanaged. Uses the Connection tab for DB/auth and the Email tab for delivery.'
$lblSapIntro.ForeColor = [System.Drawing.Color]::Gray
$lblSapIntro.Location = New-Object System.Drawing.Point(16, $sy2)
$lblSapIntro.Size = New-Object System.Drawing.Size(516, 46)
$tabSap.Controls.Add($lblSapIntro)
$sy2 += 54
New-Label $tabSap 'Job name filter (LIKE)' 16 ($sy2+2) | Out-Null
$txtSapName = New-Text $tabSap $col2 $sy2 250 '%backint%'
$sy2 += 32
New-Label $tabSap 'Deployment' 16 ($sy2+2) | Out-Null
$cboSapDeploy = New-Object System.Windows.Forms.ComboBox
$cboSapDeploy.DropDownStyle = 'DropDownList'
$cboSapDeploy.Location = New-Object System.Drawing.Point($col2, $sy2)
$cboSapDeploy.Size = New-Object System.Drawing.Size(200, 22)
[void]$cboSapDeploy.Items.AddRange(@('Managed + Unmanaged','Managed only','Unmanaged only'))
$cboSapDeploy.SelectedIndex = 0
$tabSap.Controls.Add($cboSapDeploy)
$sy2 += 32
$lblSapHint = New-Object System.Windows.Forms.Label
$lblSapHint.Text = 'Default keeps SAP HANA plugin jobs (name LIKE %backint%). Managed jobs carry a DBBKP- prefix; the rest are standalone/unmanaged. Uses the date range below.'
$lblSapHint.ForeColor = [System.Drawing.Color]::Gray
$lblSapHint.Location = New-Object System.Drawing.Point(16, $sy2)
$lblSapHint.Size = New-Object System.Drawing.Size(516, 30)
$tabSap.Controls.Add($lblSapHint)
$sy2 += 34
$rangeSap = Add-DateRange $tabSap $sy2; $dtpSapStart = $rangeSap[0]; $dtpSapEnd = $rangeSap[1]
$sy2 += 74
$lblSapRunHint = New-Object System.Windows.Forms.Label
$lblSapRunHint.Text = 'Click "Run selected report" at the bottom to generate this report.'
$lblSapRunHint.ForeColor = [System.Drawing.Color]::Gray
$lblSapRunHint.Location = New-Object System.Drawing.Point(16, $sy2)
$lblSapRunHint.Size = New-Object System.Drawing.Size(500, 20)
$tabSap.Controls.Add($lblSapRunHint)

# maps the schedule-tab report picker label -> scheduler -ReportType value
$reportTypeArgMap = @{ 'Restore' = 'Restore'; 'Backup Jobs' = 'BackupJobs'; 'All Backups' = 'AllBackups'; 'RMAN Backups' = 'RmanBackups'; 'SAP HANA' = 'SapHana' }

# =====================================================================
# Bottom bar (outside tabs)
# =====================================================================
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 620)
$lblStatus.Size = New-Object System.Drawing.Size(430, 20)
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblStatus)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run selected report'
$btnRun.Location = New-Object System.Drawing.Point(290, 645)
$btnRun.Size = New-Object System.Drawing.Size(170, 32)
$btnRun.BackColor = $clrGreen; $btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = 'Flat'
$form.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(470, 645)
$btnClose.Size = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnClose)

# ---- disclaimer (tool only - never rendered into the reports) ----
$disclaimer = New-Object System.Windows.Forms.Label
$disclaimer.Text = 'Not affiliated with Veeam. Read-only community tool - test before production use.'
$disclaimer.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$disclaimer.ForeColor = [System.Drawing.Color]::Gray
$disclaimer.Location = New-Object System.Drawing.Point(20, 690)
$disclaimer.Size = New-Object System.Drawing.Size(556, 32)
$form.Controls.Add($disclaimer)

# ---- show the bottom "Run selected report" button only on report tabs ----
# It has nothing to run on Connection / Email / Schedule (those have their own
# controls), so hide it there to avoid confusion.
$reportTabs = @($tabReport, $tabJobs, $tabAll, $tabRman, $tabSap)
$syncRunBtn = {
    $btnRun.Visible = ($reportTabs -contains $tabs.SelectedTab)
}
$tabs.Add_SelectedIndexChanged($syncRunBtn)
& $syncRunBtn

# ---- credential-field enable/disable ----
$syncAuth = {
    $needCreds = ($rbDomain.Checked -or $rbSql.Checked)
    $txtUser.Enabled = $needCreds; $txtPass.Enabled = $needCreds
    $lblUser.Enabled = $needCreds; $lblPass.Enabled = $needCreds
    $lblUserHint.Visible = $rbDomain.Checked
}
$rbDomain.Add_CheckedChanged($syncAuth)
$rbSql.Add_CheckedChanged($syncAuth)
$rbCurrent.Add_CheckedChanged($syncAuth)
& $syncAuth

# ---- cadence-field enable/disable ----
$syncCadence = {
    $c = $cboCadence.SelectedItem
    $cboDow.Enabled   = ($c -eq 'Weekly');  $lblDow.Enabled   = $cboDow.Enabled
    $numDom.Enabled   = ($c -eq 'Monthly' -or $c -eq 'Yearly'); $lblDom.Enabled = $numDom.Enabled
    $cboMonth.Enabled = ($c -eq 'Yearly');  $lblMonth.Enabled = $cboMonth.Enabled
}
$cboCadence.Add_SelectedIndexChanged($syncCadence)
& $syncCadence

# ---- helpers ----
function Get-EnteredCredential {
    if ($txtUser.Text.Trim() -eq '' -or $txtPass.Text -eq '') { return $null }
    $sec = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
    New-Object System.Management.Automation.PSCredential($txtUser.Text.Trim(), $sec)
}
function Get-CommonArgs {
    $a = @{ Database = $txtDb.Text.Trim() }
    if ($txtServer.Text.Trim()) { $a.SqlServer = $txtServer.Text.Trim() }
    if ($rbDomain.Checked)  { $a.WindowsCredential = Get-EnteredCredential }
    elseif ($rbSql.Checked) { $a.SqlCredential     = Get-EnteredCredential }
    $a
}
function Add-EmailArgs {
    param([hashtable]$a)
    if (-not $chkEmail.Checked) { return }
    $a.EmailTo    = ($txtTo.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($txtFrom.Text.Trim()) { $a.EmailFrom = $txtFrom.Text.Trim() }
    $a.SmtpServer = $txtSmtp.Text.Trim()
    $p = 25; [int]::TryParse($txtPort.Text.Trim(), [ref]$p) | Out-Null; $a.SmtpPort = $p
    if ($chkSsl.Checked) { $a.SmtpUseSsl = $true }
    if ($txtSmtpUser.Text.Trim() -ne '' -and $txtSmtpPass.Text -ne '') {
        $sec = ConvertTo-SecureString $txtSmtpPass.Text -AsPlainText -Force
        $a.SmtpCredential = New-Object System.Management.Automation.PSCredential($txtSmtpUser.Text.Trim(), $sec)
    }
}
function Test-Creds {
    if (($rbDomain.Checked -or $rbSql.Checked) -and -not (Get-EnteredCredential)) {
        [System.Windows.Forms.MessageBox]::Show('Enter a username and password for the selected authentication mode.',
            'Missing credentials','OK','Warning') | Out-Null
        return $false
    }
    $true
}

# ---- Browse ----
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if (Test-Path $txtOut.Text) { $dlg.SelectedPath = $txtOut.Text }
    if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.SelectedPath }
})

# ---- Discover ----
$btnDiscover.Add_Click({
    if (-not $txtServer.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server (IP or name) first.','Discover','OK','Warning') | Out-Null
        return
    }
    if (-not (Test-Creds)) { return }
    $form.Cursor = 'WaitCursor'; $lblStatus.Text = 'Discovering restore states & types...'; $form.Refresh()
    try {
        $a = Get-CommonArgs
        $out = & $reportTool @a -Discover *>&1 | Out-String
        $lblStatus.Text = 'Discovery complete.'
        $df = New-Object System.Windows.Forms.Form
        $df.Text = 'Discover - restore states & types'; $df.Size = New-Object System.Drawing.Size(820,560)
        $df.StartPosition = 'CenterParent'
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $true; $tb.ScrollBars = 'Both'; $tb.WordWrap = $false; $tb.ReadOnly = $true
        $tb.Dock = 'Fill'; $tb.Font = New-Object System.Drawing.Font('Consolas',9)
        $tb.Text = $out
        $df.Controls.Add($tb)
        $df.ShowDialog() | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Discovery failed:`n`n$($_.Exception.Message)",'Discover','OK','Error') | Out-Null
        $lblStatus.Text = 'Discovery failed.'
    }
    finally { $form.Cursor = 'Default' }
})

# ---- validate the Email tab when enabled ----
function Test-EmailReady {
    if (-not $chkEmail.Checked) { return $true }
    if (-not ($txtTo.Text.Trim())) {
        [System.Windows.Forms.MessageBox]::Show('Enter at least one recipient in the Email tab (To).','Email','OK','Warning') | Out-Null; return $false
    }
    if (-not ($txtSmtp.Text.Trim())) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SMTP server in the Email tab.','Email','OK','Warning') | Out-Null; return $false
    }
    $true
}

# ---- Run (bottom button runs the CURRENTLY SELECTED report tab) ----
$btnRun.Add_Click({
    # Dispatch to whichever report tab is showing, so this button is never
    # "always the Restore report". Each report tab also has its own Run button.
    $sel = $tabs.SelectedTab
    if ($sel -eq $tabJobs) { Invoke-JobsReport; return }
    if ($sel -eq $tabAll)  { Invoke-AllReport;  return }
    if ($sel -eq $tabRman) { Invoke-RmanReport; return }
    if ($sel -eq $tabSap)  { Invoke-SapReport;  return }
    if ($sel -eq $tabConn -or $sel -eq $tabEmail -or $sel -eq $tabSchedule) {
        [System.Windows.Forms.MessageBox]::Show(
            'Select a report tab first (Restoration report, Backup Jobs, All Backups, RMAN Backups, or SAP HANA), then click Run.',
            'Run report','OK','Information') | Out-Null
        return
    }
    # Otherwise the Restoration report tab is selected.
    Invoke-RestoreReport
})

# ---- Create schedule ----
$btnSchedule.Add_Click({
    if (-not (Test-Path $schedTool)) {
        [System.Windows.Forms.MessageBox]::Show("Cannot find Register-RestoreReportSchedule.ps1 next to this GUI.",'Schedule','OK','Error') | Out-Null; return
    }
    if (-not $txtServer.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server on the Connection tab first.','Schedule','OK','Warning') | Out-Null; return
    }
    if (-not (Test-EmailReady)) { return }

    $sa = @{
        Cadence   = [string]$cboCadence.SelectedItem
        Time      = $txtSchedTime.Text.Trim()
        SqlServer = $txtServer.Text.Trim()
        Database  = $txtDb.Text.Trim()
    }
    switch ($sa.Cadence) {
        'Weekly'  { $sa.DayOfWeek  = [string]$cboDow.SelectedItem }
        'Monthly' { $sa.DayOfMonth = [int]$numDom.Value }
        'Yearly'  { $sa.DayOfMonth = [int]$numDom.Value; $sa.Month = [string]$cboMonth.SelectedItem }
    }
    if ($chkPdf.Checked) { $sa.Pdf = $true }
    # Target the report chosen in the "Report to schedule" picker (same scheduler for all).
    $sa.ReportType = $reportTypeArgMap[[string]$cboSchedReport.SelectedItem]
    if ($txtSchedFilter.Text.Trim()) {
        $sa.Filter = $txtSchedFilter.Text.Trim()
    }
    elseif ([string]$cboSchedReport.SelectedItem -eq 'Restore' -and $txtRestoreType.Text.Trim()) {
        $sa.RestoreType = ($txtRestoreType.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if ($txtOut.Text.Trim()) { $sa.OutputFolder = $txtOut.Text.Trim() }
    if ($chkEmail.Checked) {
        $sa.EmailTo    = ($txtTo.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($txtFrom.Text.Trim()) { $sa.EmailFrom = $txtFrom.Text.Trim() }
        $sa.SmtpServer = $txtSmtp.Text.Trim()
        $p = 25; [int]::TryParse($txtPort.Text.Trim(), [ref]$p) | Out-Null; $sa.SmtpPort = $p
        if ($chkSsl.Checked) { $sa.SmtpUseSsl = $true }
    }
    if ($txtRunUser.Text.Trim() -ne '' -and $txtRunPass.Text -ne '') {
        $sec = ConvertTo-SecureString $txtRunPass.Text -AsPlainText -Force
        $sa.RunAsCredential = New-Object System.Management.Automation.PSCredential($txtRunUser.Text.Trim(), $sec)
    }

    $form.Cursor = 'WaitCursor'; $lblStatus.Text = 'Creating scheduled task...'; $form.Refresh()
    try {
        $out = & $schedTool @sa *>&1 | Out-String
        $lblStatus.Text = 'Scheduled task created.'
        [System.Windows.Forms.MessageBox]::Show("Scheduled task created.`n`n$out",'Schedule','OK','Information') | Out-Null
    }
    catch {
        $lblStatus.Text = 'Schedule failed.'
        [System.Windows.Forms.MessageBox]::Show("Could not create the schedule:`n`n$($_.Exception.Message)",'Schedule','OK','Error') | Out-Null
    }
    finally { $form.Cursor = 'Default' }
})

# ---- Remove schedule ----
$btnUnschedule.Add_Click({
    if (-not (Test-Path $schedTool)) { return }
    $form.Cursor = 'WaitCursor'; $lblStatus.Text = 'Removing scheduled task...'; $form.Refresh()
    try {
        $rt = $reportTypeArgMap[[string]$cboSchedReport.SelectedItem]
        $out = & $schedTool -Unregister -ReportType $rt *>&1 | Out-String
        $lblStatus.Text = 'Scheduled task removed.'
        [System.Windows.Forms.MessageBox]::Show($out,'Schedule','OK','Information') | Out-Null
    }
    catch {
        $lblStatus.Text = 'Remove failed.'
        [System.Windows.Forms.MessageBox]::Show("Could not remove the schedule:`n`n$($_.Exception.Message)",'Schedule','OK','Error') | Out-Null
    }
    finally { $form.Cursor = 'Default' }
})

# ---- shared runner for the added report tabs (reuses connection + email + pdf) ----
function Invoke-NewReport {
    param([string]$ToolPath, [hashtable]$Extra, $StartCtl = $null, $EndCtl = $null)
    if (-not (Test-Path $ToolPath)) {
        [System.Windows.Forms.MessageBox]::Show("Cannot find the report engine:`n$ToolPath",'Run report','OK','Error') | Out-Null; return
    }
    if (-not $txtServer.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server (IP or name) on the Connection tab.','Run report','OK','Warning') | Out-Null; return
    }
    if (-not (Test-Creds)) { return }
    if (-not (Test-EmailReady)) { return }
    if ($StartCtl -and $EndCtl -and $EndCtl.Value -le $StartCtl.Value) {
        [System.Windows.Forms.MessageBox]::Show('End date/time must be after the start.','Run report','OK','Warning') | Out-Null; return
    }
    $form.Cursor = 'WaitCursor'; $btnRun.Enabled = $false
    $lblStatus.Text = 'Running report...'; $form.Refresh()
    try {
        $a = Get-CommonArgs
        if ($StartCtl -and $EndCtl) { $a.StartDate = $StartCtl.Value; $a.EndDate = $EndCtl.Value }
        if ($txtOut.Text.Trim()) { $a.OutputFolder = $txtOut.Text.Trim() }
        if ($chkPdf.Checked) { $a.Pdf = $true }
        Add-EmailArgs $a
        foreach ($k in $Extra.Keys) { $a[$k] = $Extra[$k] }
        $result = @(& $ToolPath @a) | Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties['Html'] } | Select-Object -Last 1
        if ($result) {
            $lblStatus.Text = 'Done - report generated.'
            $summary = ($result.PSObject.Properties | Where-Object { $_.Name -notin @('Csv','Html','Pdf') } | ForEach-Object { "{0,-13}: {1}" -f $_.Name, $_.Value }) -join "`n"
            $pdfNote   = if ($result.PSObject.Properties['Pdf'] -and $result.Pdf) { "`nPDF : $($result.Pdf)" } else { '' }
            $emailNote = if ($chkEmail.Checked) { "`nEmailed  : $($txtTo.Text.Trim())" } else { '' }
            $msg = "Report generated.`n`n$summary$emailNote`n`nCSV : $($result.Csv)`nHTML: $($result.Html)$pdfNote`n`nOpen the HTML report now?"
            if ([System.Windows.Forms.MessageBox]::Show($msg,'Report complete','YesNo','Information') -eq 'Yes') { Start-Process $result.Html }
        }
        else {
            $lblStatus.Text = 'Completed - no matching rows.'
            [System.Windows.Forms.MessageBox]::Show('The run completed but returned no data. Check the date range / filters (or use Discover).','Run report','OK','Information') | Out-Null
        }
    }
    catch {
        $lblStatus.Text = 'Failed.'
        [System.Windows.Forms.MessageBox]::Show("Report failed:`n`n$($_.Exception.Message)",'Run report','OK','Error') | Out-Null
    }
    finally { $form.Cursor = 'Default'; $btnRun.Enabled = $true }
}

# Per-report run logic - called by the single bottom "Run selected report" button.
function Invoke-JobsReport {
    $extra = @{}
    if ($txtJobsName.Text.Trim())   { $extra.NameLike    = $txtJobsName.Text.Trim() }
    if ($txtJobsPolicy.Text.Trim()) { $extra.PolicyLike  = $txtJobsPolicy.Text.Trim() }
    if ($chkJobsEnabled.Checked)    { $extra.EnabledOnly = $true }
    if ($chkJobsCalendar.Checked)   { $extra.Calendar    = $true }
    Invoke-NewReport -ToolPath $jobsTool -Extra $extra -StartCtl $dtpJobsStart -EndCtl $dtpJobsEnd
}
function Invoke-RmanReport {
    $extra = @{}
    if ($chkRmanAll.Checked) { $extra.IncludeAllBackups = $true }
    elseif ($txtRmanName.Text.Trim()) { $extra.JobNameLike = $txtRmanName.Text.Trim() }
    Invoke-NewReport -ToolPath $rmanTool -Extra $extra -StartCtl $dtpRmanStart -EndCtl $dtpRmanEnd
}
function Invoke-AllReport {
    $extra = @{}
    if ($txtAllName.Text.Trim()) { $extra.NameLike = $txtAllName.Text.Trim() }
    if ($cboAllStatus.SelectedItem -and [string]$cboAllStatus.SelectedItem -ne 'All') { $extra.StatusFilter = [string]$cboAllStatus.SelectedItem }
    Invoke-NewReport -ToolPath $allTool -Extra $extra -StartCtl $dtpAllStart -EndCtl $dtpAllEnd
}
function Invoke-SapReport {
    $extra = @{}
    if ($txtSapName.Text.Trim()) { $extra.JobNameLike = $txtSapName.Text.Trim() }
    switch ([string]$cboSapDeploy.SelectedItem) {
        'Managed only'   { $extra.Deployment = 'Managed' }
        'Unmanaged only' { $extra.Deployment = 'Unmanaged' }
    }
    Invoke-NewReport -ToolPath $sapTool -Extra $extra -StartCtl $dtpSapStart -EndCtl $dtpSapEnd
}
function Invoke-RestoreReport {
    if (-not $txtServer.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server (IP or name) on the Connection tab.','Run report','OK','Warning') | Out-Null; return
    }
    if (-not (Test-Creds)) { return }
    if (-not (Test-EmailReady)) { return }
    if ($dtpEnd.Value -le $dtpStart.Value) {
        [System.Windows.Forms.MessageBox]::Show('End date/time must be after the start (Restoration report tab).','Run report','OK','Warning') | Out-Null; return
    }
    $rt = @()
    foreach ($p in ($txtRestoreType.Text -split ',')) { $p = $p.Trim(); if ($p -ne '') { $rt += $p } }
    $form.Cursor = 'WaitCursor'; $btnRun.Enabled = $false
    $lblStatus.Text = 'Running report...'; $form.Refresh()
    try {
        $a = Get-CommonArgs
        if ($rt.Count -gt 0) { $a.RestoreType = [string[]]$rt }
        $a.StartDate    = $dtpStart.Value
        $a.EndDate      = $dtpEnd.Value
        $a.OutputFolder = $txtOut.Text.Trim()
        $a.RetryWindowHours = if ($chkRetry.Checked) { [int]$numRetryHrs.Value } else { 0 }
        if ($chkByJob.Checked) { $a.ByRestoreJob = $true }
        if ($chkPdf.Checked) { $a.Pdf = $true }
        Add-EmailArgs $a
        $result = @(& $reportTool @a) | Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties['Html'] } | Select-Object -Last 1
        if ($result) {
            $emailNote = if ($chkEmail.Checked) { "`nEmailed  : $($txtTo.Text.Trim())" } else { '' }
            $pdfNote   = if ($result.PSObject.Properties['Pdf'] -and $result.Pdf) { "`nPDF : $($result.Pdf)" } else { '' }
            $lblStatus.Text = "Done: $($result.Total) restores, $($result.SuccessRate)% success."
            $msg = "Report generated.`n`nRestores : $($result.Total)`nSuccess  : $($result.Success)`nFailed   : $($result.Failed)`nSkipped  : $($result.Skipped)`nSuccess rate: $($result.SuccessRate)%$emailNote`n`nCSV : $($result.Csv)`nHTML: $($result.Html)$pdfNote`n`nOpen the HTML report now?"
            if ([System.Windows.Forms.MessageBox]::Show($msg,'Report complete','YesNo','Information') -eq 'Yes') { Start-Process $result.Html }
        }
        else {
            $lblStatus.Text = 'Completed - no matching restores (check date range / type / mapping).'
            [System.Windows.Forms.MessageBox]::Show('The run completed but returned no restores. Check the date range, restore-type filter, and (via Discover) the state/type mapping.','Run report','OK','Information') | Out-Null
        }
    }
    catch {
        $lblStatus.Text = 'Failed.'
        [System.Windows.Forms.MessageBox]::Show("Report failed:`n`n$($_.Exception.Message)",'Run report','OK','Error') | Out-Null
    }
    finally { $form.Cursor = 'Default'; $btnRun.Enabled = $true }
}
# Run one report by its friendly name (used by the Email tab's "e-mail now").
function Invoke-ReportByName {
    param([string]$Name)
    switch ($Name) {
        'Restore'      { Invoke-RestoreReport }
        'Backup Jobs'  { Invoke-JobsReport }
        'All Backups'  { Invoke-AllReport }
        'RMAN Backups' { Invoke-RmanReport }
        'SAP HANA'     { Invoke-SapReport }
    }
}
$btnJobsDiscover.Add_Click({
    if (-not $txtServer.Text.Trim()) { [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server on the Connection tab first.','Discover','OK','Warning') | Out-Null; return }
    if (-not (Test-Creds)) { return }
    $form.Cursor = 'WaitCursor'; $lblStatus.Text = 'Discovering job types...'; $form.Refresh()
    try {
        $a = Get-CommonArgs
        $out = & $jobsTool @a -Discover *>&1 | Out-String
        $lblStatus.Text = 'Discovery complete.'
        $df = New-Object System.Windows.Forms.Form
        $df.Text = 'Discover - job types'; $df.Size = New-Object System.Drawing.Size(820,540); $df.StartPosition = 'CenterParent'
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline=$true; $tb.ScrollBars='Both'; $tb.WordWrap=$false; $tb.ReadOnly=$true; $tb.Dock='Fill'
        $tb.Font = New-Object System.Drawing.Font('Consolas',9); $tb.Text = $out
        $df.Controls.Add($tb); $df.ShowDialog() | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Discovery failed:`n`n$($_.Exception.Message)",'Discover','OK','Error') | Out-Null; $lblStatus.Text = 'Discovery failed.'
    }
    finally { $form.Cursor = 'Default' }
})

# ---- Email tab: generate the chosen report and e-mail it now ----
$btnEmailNow.Add_Click({
    if (-not $txtServer.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server (IP or name) on the Connection tab.','Email now','OK','Warning') | Out-Null; return
    }
    if (-not $txtTo.Text.Trim() -or -not $txtSmtp.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the recipients (To) and the SMTP server on this Email tab first.','Email now','OK','Warning') | Out-Null; return
    }
    # Force e-mail on for this run regardless of the "each time it runs" checkbox.
    $prev = $chkEmail.Checked
    $chkEmail.Checked = $true
    try   { Invoke-ReportByName ([string]$cboEmailReport.SelectedItem) }
    finally { $chkEmail.Checked = $prev }
})

$btnClose.Add_Click({ $form.Close() })

if ($SelfTest) {
    $tabCount = $tabs.TabPages.Count
    Write-Host "SelfTest OK - form built with $($form.Controls.Count) top-level controls, $tabCount tabs; auth + cadence sync + handlers wired." -ForegroundColor Green
    if ($env:RESTORE_FORM_SHOT) {
        if ($env:RESTORE_FORM_TAB) { $tabs.SelectedIndex = [int]$env:RESTORE_FORM_TAB }
        $form.StartPosition = 'Manual'
        $form.Location = New-Object System.Drawing.Point(-3000, -3000)
        $form.Show(); $form.Refresh()
        $bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
        $form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0,0,$form.Width,$form.Height)))
        $bmp.Save($env:RESTORE_FORM_SHOT, [System.Drawing.Imaging.ImageFormat]::Png)
        $form.Hide()
        Write-Host "Saved form preview: $env:RESTORE_FORM_SHOT"
    }
    $form.Dispose()
    return
}

[void]$form.ShowDialog()
$form.Dispose()
