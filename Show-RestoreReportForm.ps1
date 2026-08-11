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
$form.Text           = 'Veeam ONE - Restore Report'
$form.Size           = New-Object System.Drawing.Size(600, 758)
$form.StartPosition  = 'CenterScreen'
$form.FormBorderStyle= 'FixedDialog'
$form.MaximizeBox    = $false
$form.BackColor      = $clrBg
$form.Font           = $fontLbl

# header
$hdr = New-Object System.Windows.Forms.Label
$hdr.Text = 'VeeamONE Restoration Status Percentage Report'
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

# ---- tab control ----
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(16, 50)
$tabs.Size     = New-Object System.Drawing.Size(560, 560)
$form.Controls.Add($tabs)

$tabReport   = New-Object System.Windows.Forms.TabPage; $tabReport.Text   = '  Report  '
$tabEmail    = New-Object System.Windows.Forms.TabPage; $tabEmail.Text    = '  Email  '
$tabSchedule = New-Object System.Windows.Forms.TabPage; $tabSchedule.Text = '  Schedule  '
foreach ($tp in $tabReport, $tabEmail, $tabSchedule) { $tp.BackColor = [System.Drawing.Color]::White; $tabs.TabPages.Add($tp) }

$col2 = 180     # x for input controls inside a tab

# =====================================================================
# TAB 1 - REPORT
# =====================================================================
$y = 18
New-Label $tabReport 'SQL Server (IP or name)' 16 ($y+2) | Out-Null
$txtServer = New-Text $tabReport $col2 $y 250
$y += 32
New-Label $tabReport 'Database' 16 ($y+2) | Out-Null
$txtDb = New-Text $tabReport $col2 $y 250 'VeeamONE'
$y += 38

# Authentication group
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = 'Authentication'
$grp.Location = New-Object System.Drawing.Point(16, $y)
$grp.Size = New-Object System.Drawing.Size(516, 150)
$grp.ForeColor = $clrText
$tabReport.Controls.Add($grp)

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

$y += 162

# Restore type filter + discover
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

# Date/time pickers
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
New-Label $tabReport 'Start (date & time)' 16 ($y+2) | Out-Null
$dtpStart = New-DateTimePicker $tabReport $col2 $y ((Get-Date).Date.AddDays(-30))
$y += 32
New-Label $tabReport 'End (date & time)' 16 ($y+2) | Out-Null
$dtpEnd = New-DateTimePicker $tabReport $col2 $y (Get-Date)
$y += 38

# Output folder + PDF
New-Label $tabReport 'Output folder' 16 ($y+2) | Out-Null
$txtOut = New-Text $tabReport $col2 $y 250 (Join-Path $scriptDir 'output')
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '...'
$btnBrowse.Location = New-Object System.Drawing.Point(437, ($y-1))
$btnBrowse.Size = New-Object System.Drawing.Size(30, 24)
$tabReport.Controls.Add($btnBrowse)
$y += 32
$chkPdf = New-Object System.Windows.Forms.CheckBox
$chkPdf.Text = 'Also generate a PDF (uses Edge/Chrome)'
$chkPdf.Location = New-Object System.Drawing.Point($col2, $y)
$chkPdf.Size = New-Object System.Drawing.Size(320, 20)
$tabReport.Controls.Add($chkPdf)

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
$lblEmailHint.Text = 'When enabled, the CSV + HTML (+ PDF if selected) are attached to the e-mail.'
$lblEmailHint.ForeColor = [System.Drawing.Color]::Gray
$lblEmailHint.Location = New-Object System.Drawing.Point(16, $ey)
$lblEmailHint.Size = New-Object System.Drawing.Size(516, 32)
$tabEmail.Controls.Add($lblEmailHint)

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
# Bottom bar (outside tabs)
# =====================================================================
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 620)
$lblStatus.Size = New-Object System.Drawing.Size(430, 20)
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblStatus)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run report'
$btnRun.Location = New-Object System.Drawing.Point(340, 645)
$btnRun.Size = New-Object System.Drawing.Size(120, 32)
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

# ---- Run ----
$btnRun.Add_Click({
    if (-not $txtServer.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server (IP or name) on the Report tab.','Run report','OK','Warning') | Out-Null; return
    }
    if (-not (Test-Creds)) { return }
    if (-not (Test-EmailReady)) { return }
    if ($dtpEnd.Value -le $dtpStart.Value) {
        [System.Windows.Forms.MessageBox]::Show('End date/time must be after the start.','Run report','OK','Warning') | Out-Null; return
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
        if ($chkPdf.Checked) { $a.Pdf = $true }
        Add-EmailArgs $a

        $result = @(& $reportTool @a) | Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties['Html'] } | Select-Object -Last 1

        if ($result) {
            $emailNote = if ($chkEmail.Checked) { "`nEmailed  : $($txtTo.Text.Trim())" } else { '' }
            $pdfNote   = if ($result.PSObject.Properties['Pdf'] -and $result.Pdf) { "`nPDF : $($result.Pdf)" } else { '' }
            $lblStatus.Text = "Done: $($result.Total) restores, $($result.SuccessRate)% success."
            $msg = "Report generated.`n`nRestores : $($result.Total)`nSuccess  : $($result.Success)`nFailed   : $($result.Failed)`nSkipped  : $($result.Skipped)`nSuccess rate: $($result.SuccessRate)%$emailNote`n`nCSV : $($result.Csv)`nHTML: $($result.Html)$pdfNote`n`nOpen the HTML report now?"
            if ([System.Windows.Forms.MessageBox]::Show($msg,'Report complete','YesNo','Information') -eq 'Yes') {
                Start-Process $result.Html
            }
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
})

# ---- Create schedule ----
$btnSchedule.Add_Click({
    if (-not (Test-Path $schedTool)) {
        [System.Windows.Forms.MessageBox]::Show("Cannot find Register-RestoreReportSchedule.ps1 next to this GUI.",'Schedule','OK','Error') | Out-Null; return
    }
    if (-not $txtServer.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show('Enter the SQL Server on the Report tab first.','Schedule','OK','Warning') | Out-Null; return
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
    if ($txtRestoreType.Text.Trim()) { $sa.RestoreType = ($txtRestoreType.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
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
        $out = & $schedTool -Unregister *>&1 | Out-String
        $lblStatus.Text = 'Scheduled task removed.'
        [System.Windows.Forms.MessageBox]::Show($out,'Schedule','OK','Information') | Out-Null
    }
    catch {
        $lblStatus.Text = 'Remove failed.'
        [System.Windows.Forms.MessageBox]::Show("Could not remove the schedule:`n`n$($_.Exception.Message)",'Schedule','OK','Error') | Out-Null
    }
    finally { $form.Cursor = 'Default' }
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
