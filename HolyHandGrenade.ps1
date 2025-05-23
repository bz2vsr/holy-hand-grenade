Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# === Predefined OUs: Used to populate the Target OU dropdown ===
$predefinedOUs = @(
    "OU=General Use,OU=Workstations,DC=corp,DC=domain,DC=com",
    "OU=Disposed,OU=Workstations,DC=corp,DC=domain,DC=com"
)

# Define Disposed OU Path: Used to delete computers from AD
$DisposedPathPrefix = "OU=Disposed,OU=Workstations"

# === MECM Action Configuration Settings ===
$CollectionId = 'COL00001'
$SiteCode = "SITE"
$ProviderMachineName = "server.domain.com"

function Show-ErrorPopup($message, $title = "Error") {
    [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Show-YesNoPopup($message, $title = "Missing Dependency") {
    return [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
}

# === PowerShell Version Check ===
if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
    Show-ErrorPopup "This script requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
    return
}

# === Admin Check (with auto-elevation and hidden console) ===
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    $scriptPath = $MyInvocation.MyCommand.Path
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
    $psi.Verb = 'runas'
    $psi.WindowStyle = 'Hidden'
    try {
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    } catch {
        Show-ErrorPopup "This script must be run as Administrator. Relaunch cancelled." "Administrator Required"
    }
    exit
}

# === Hide Console Window ===
Add-Type -Name Win -Namespace Console -MemberDefinition @"
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@

$consolePtr = [Console.Win]::GetConsoleWindow()
[Console.Win]::ShowWindow($consolePtr, 0)  # 0 = Hide, 5 = Show

# === RSAT Check ===
$rsatInstalled = Get-WindowsCapability -Online | Where-Object {
    $_.Name -like "Rsat.ActiveDirectory.DS-LDS.Tools*" -and $_.State -eq "Installed"
}

if (-not $rsatInstalled) {
    $choice = Show-YesNoPopup "The RSAT Active Directory tools are not installed.`nWould you like to install them now?"

    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Write-Host "Installing RSAT..."
            Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop | Out-Null
            Start-Sleep -Seconds 5
        }
        catch {
            Show-ErrorPopup "RSAT installation failed: $($_.Exception.Message)"
            return
        }
    } else {
        Show-ErrorPopup "RSAT is required to continue. Exiting..."
        return
    }
}

# === Configuration Manager Check ===
$cmModuleAvailable = $false

try {
    # Check if ConfigurationManager module is available
    $cmModulePath = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    if (Test-Path $cmModulePath) {
        Import-Module $cmModulePath -ErrorAction Stop
        $cmModuleAvailable = $true
        
        # Check/Create PSDrive
        if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName -ErrorAction Stop
        }
    } else {
        throw "Configuration Manager console not found"
    }
} catch {
    $choice = Show-YesNoPopup "Configuration Manager Console (CMConsole) not detected. Ensure CMConsole is installed and restart this application.`n`nIf you proceed, MECM-related actions will not work. Continue anyways?" "CMConsole Not Found"
    if ($choice -eq [System.Windows.Forms.DialogResult]::No) {
        return
    }
}

# === GUI Form Setup (Dark Mode + Custom Title Bar + Border Panel) ===
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'None'
$form.ForeColor = [System.Drawing.Color]::White
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.Size = New-Object System.Drawing.Size(600, 750)
$form.StartPosition = "CenterScreen"

# === Custom Title Bar ===
$titleBar = New-Object System.Windows.Forms.Panel
$titleBar.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$titleBar.Dock = 'Top'
$titleBar.Height = 30
$titleBar.BringToFront()

# Add grenade icon 
$grenadeBox = New-Object System.Windows.Forms.PictureBox
$grenadeBox.Size = New-Object System.Drawing.Size(22, 22)
$grenadeBox.Location = New-Object System.Drawing.Point(3, 4)
$grenadeBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$grenadeBox.Visible = $true
# Courtesy of Claude Sonnet :-)
$grenadeBox.Add_Paint({
    $g = $_.Graphics
    $gold = [System.Drawing.Brushes]::Gold
    $g.FillEllipse($gold, 2, 4, 14, 14)
    $g.DrawEllipse([System.Drawing.Pens]::Black, 2, 4, 14, 14)
    $g.FillRectangle([System.Drawing.Brushes]::DarkGray, 7, 0, 4, 6)
    $g.DrawRectangle([System.Drawing.Pens]::Black, 7, 0, 4, 6)
    $crossPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2)
    $g.DrawLine($crossPen, 9, 1, 11, 5) 
    $g.DrawLine($crossPen, 8, 3, 12, 3) 
    $crossPen.Dispose()
    $g.DrawArc([System.Drawing.Pens]::Black, 10, 0, 7, 7, 0, 270)
})
$titleBar.Controls.Add($grenadeBox)

# Title bar with close button
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "HolyHandGrenade - MECM && AD Device Manager"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(23, 7)
$titleBar.Controls.Add($titleLabel)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "X"
$closeButton.ForeColor = [System.Drawing.Color]::White
$closeButton.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$closeButton.FlatStyle = 'Flat'
$closeButton.FlatAppearance.BorderSize = 0
$closeButton.Size = New-Object System.Drawing.Size(30, 25)
$closeButton.Dock = 'Right'
$closeButton.Add_MouseEnter({ $closeButton.BackColor = [System.Drawing.Color]::DarkRed })
$closeButton.Add_MouseLeave({ $closeButton.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48) })
$closeButton.Add_Click({ $form.Close() })
$titleBar.Controls.Add($closeButton)

# Enable dragging the form
Add-Type -Namespace Win32Functions -Name NativeMethods -MemberDefinition @"
    [DllImport("user32.dll")]
    public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
    [DllImport("user32.dll")]
    public static extern bool ReleaseCapture();
"@

$titleBar.Add_MouseDown({
    [Win32Functions.NativeMethods]::ReleaseCapture() | Out-Null
    [Win32Functions.NativeMethods]::SendMessage($form.Handle, 0xA1, 0x2, 0) | Out-Null
})

$form.Controls.Add($titleBar)

# === Warning Label (for ADWS error) ===
$labelWarning = New-Object System.Windows.Forms.Label
$labelWarning.Margin = '0,10,0,0'
$labelWarning.Location = New-Object System.Drawing.Point(10, 40)
$labelWarning.ForeColor = [System.Drawing.Color]::OrangeRed
$labelWarning.AutoSize = $true
$form.Controls.Add($labelWarning)

# === Try importing the AD module ===
try {
    Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
    # Test Get-ADDomain to verify AD Web Services are up
    Get-ADDomain -ErrorAction Stop | Out-Null
} catch {
    $labelWarning.Text = "WARNING: Unable to connect to Active Directory Web Services." + [Environment]::NewLine + "Make sure your network connection is active and your password is still valid (not expired)."
}

# === Target OU Dropdown with Placeholder ===
$labelOU = New-Object System.Windows.Forms.Label
$labelOU.ForeColor = [System.Drawing.Color]::White
$labelOU.Text = "Target OU:"
$labelOU.AutoSize = $true
$labelOU.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($labelOU)

$comboBoxOU = New-Object System.Windows.Forms.ComboBox
$comboBoxOU.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$comboBoxOU.ForeColor = [System.Drawing.Color]::White
$comboBoxOU.Size = New-Object System.Drawing.Size(560, 20)
$comboBoxOU.Location = New-Object System.Drawing.Point(10, 80)
$comboBoxOU.DropDownStyle = 'DropDown'
$placeholderOU = "<Enter OU or select from dropdown...>"
$comboBoxOU.Items.Add($placeholderOU)

$comboBoxOU.Items.AddRange($predefinedOUs)
$comboBoxOU.SelectedIndex = 0
$form.Controls.Add($comboBoxOU)

$comboBoxOU.Add_GotFocus({
    if ($comboBoxOU.Text -eq $placeholderOU) {
        $comboBoxOU.Text = ""
    }
})

# === Computer Input + File Import ===
$labelComputers = New-Object System.Windows.Forms.Label
$labelComputers.ForeColor = [System.Drawing.Color]::White
$labelComputers.Text = "Enter or import computer names (one per line):"
$labelComputers.AutoSize = $true
$labelComputers.Location = New-Object System.Drawing.Point(10, 115)
$form.Controls.Add($labelComputers)

$buttonImport = New-Object System.Windows.Forms.Button
$buttonImport.Text = "Import from File..."
$buttonImport.Size = New-Object System.Drawing.Size(130, 24)
$buttonImport.Location = New-Object System.Drawing.Point(440, 110)
$form.Controls.Add($buttonImport)

$textBoxComputers = New-Object System.Windows.Forms.TextBox
$textBoxComputers.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$textBoxComputers.ForeColor = [System.Drawing.Color]::White
$textBoxComputers.Multiline = $true
$textBoxComputers.ScrollBars = "Vertical"
$textBoxComputers.Size = New-Object System.Drawing.Size(560, 200)
$textBoxComputers.Location = New-Object System.Drawing.Point(10, 140)
$form.Controls.Add($textBoxComputers)

$labelNote = New-Object System.Windows.Forms.Label
$labelNote.ForeColor = [System.Drawing.Color]::OrangeRed
$labelNote.Text = "Note: Computer names must be line-separated (one per line)."
$labelNote.AutoSize = $true
$labelNote.Location = New-Object System.Drawing.Point(10, 345)
$form.Controls.Add($labelNote)

# === Log File Paths ===
$tempDir = Join-Path $env:TEMP "HolyHandGrenade"
if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFileName = "HHGLog_${timestamp}.txt"
$logFilePath = Join-Path $tempDir $logFileName
$appLogPath = $logFilePath
Start-Transcript -Path $appLogPath -Append -NoClobber | Out-Null

# === Log Label + Output Box ===
$labelLog = New-Object System.Windows.Forms.Label
$labelLog.ForeColor = [System.Drawing.Color]::White
$labelLog.Text = "Results:"
$labelLog.AutoSize = $true
$labelLog.Location = New-Object System.Drawing.Point(10, 370)
$form.Controls.Add($labelLog)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$logBox.ForeColor = [System.Drawing.Color]::White
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.WordWrap = $false
$logBox.Size = New-Object System.Drawing.Size(560, 200)
$logBox.Location = New-Object System.Drawing.Point(10, 390)
$form.Controls.Add($logBox)

# === Buttons: Copy Log + Clear Results ===
$buttonCopyLog = New-Object System.Windows.Forms.Button
$buttonCopyLog.Text = "Copy Results"
$buttonCopyLog.Size = New-Object System.Drawing.Size(100, 30)
$buttonCopyLog.Location = New-Object System.Drawing.Point(330, 607)
$form.Controls.Add($buttonCopyLog)

$buttonClearResults = New-Object System.Windows.Forms.Button
$buttonClearResults.Text = "Clear Results"
$buttonClearResults.Size = New-Object System.Drawing.Size(100, 30)
$buttonClearResults.Location = New-Object System.Drawing.Point(440, 607)
$form.Controls.Add($buttonClearResults)

# === Footer FlowLayoutPanel with Main Log and Link ===
$footerPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$footerPanel.Dock = 'Bottom'
$footerPanel.Height = 28
$footerPanel.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$footerPanel.WrapContents = $false
$footerPanel.AutoSize = $false
$footerPanel.FlowDirection = 'LeftToRight'
$form.Controls.Add($footerPanel)

$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = "Main Log:"
$footerLabel.AutoSize = $true
$footerLabel.ForeColor = [System.Drawing.Color]::DarkGray
$footerLabel.Margin = '3,6,0,6'
$footerPanel.Controls.Add($footerLabel)

$footerLogLink = New-Object System.Windows.Forms.Label
$footerLogLink.Text = "%TEMP%\HolyHandGrenade\$logFileName"
$footerLogLink.AutoSize = $true
$footerLogLink.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$footerLogLink.Font = New-Object System.Drawing.Font($footerLabel.Font, [System.Drawing.FontStyle]::Underline)
$footerLogLink.Cursor = [System.Windows.Forms.Cursors]::Hand
$footerLogLink.Margin = '0,6,0,6'
$footerPanel.Controls.Add($footerLogLink)

# Add tooltip to show resolved log file path
$footerLogTooltip = New-Object System.Windows.Forms.ToolTip
$footerLogTooltip.SetToolTip($footerLogLink, $logFilePath)

$footerLogLink.Add_Click({
    if (Test-Path $logFilePath) {
        Start-Process notepad.exe $logFilePath
    } else {
        Show-ErrorPopup "Log file not found." + [Environment]::NewLine + "($logFilePath)" "File Missing"
    }
})

function Write-Log {
    param (
        [string]$message,
        [bool]$isFailure = $false,
        [bool]$isWarning = $false
    )
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.SelectionLength = 0

    if ($isFailure) {
        $logBox.SelectionColor = [System.Drawing.Color]::OrangeRed
    } elseif ($isWarning) {
        $logBox.SelectionColor = [System.Drawing.Color]::Orange
    } else {
        $logBox.SelectionColor = [System.Drawing.Color]::White
    }
    $logBox.AppendText("$message`r`n")
    $logBox.SelectionColor = $logBox.ForeColor  # Reset color
    Write-Host $message
}

function Get-OUFromDN {
    param (
        [string]$distinguishedName
    )
    # Extract only the OU portion from Distinguished Name
    # Example: "CN=Computer,OU=General Use,OU=Workstations,DC=corp,DC=domain,DC=com" 
    # Returns: "OU=General Use,OU=Workstations"
    try {
        # Remove CN= portion and anything after DC=
        $cleaned = $distinguishedName -replace '^CN=[^,]*,', '' -replace ',DC=.*$', ''
        return $cleaned
    } catch {
        # Fallback: return the original DN if parsing fails
        return $distinguishedName
    }
}

$buttonImport.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "Text files (*.txt)|*.txt"
    $fileDialog.Title = "Select a .txt file with computer names"

    if ($fileDialog.ShowDialog() -eq "OK") {
        try {
            $lines = Get-Content $fileDialog.FileName
            $textBoxComputers.Lines = $lines
        } catch {
            Show-ErrorPopup "Failed to load file: $($_.Exception.Message)"
        }
    }
})

$buttonRun.Enabled = $false
$buttonRun.BackColor = [System.Drawing.Color]::Gray
$buttonRun.ForeColor = [System.Drawing.Color]::FromArgb(30,30,30)  # Dark text for contrast

# Add Action ComboBox and Run button in its place
$labelAction = New-Object System.Windows.Forms.Label
$labelAction.ForeColor = [System.Drawing.Color]::White
$labelAction.Text = "Action:"
$labelAction.AutoSize = $true
$labelAction.Location = New-Object System.Drawing.Point(10, 610)
$form.Controls.Add($labelAction)

$comboBoxAction = New-Object System.Windows.Forms.ComboBox
$comboBoxAction.Items.Clear()
$comboBoxAction.Items.AddRange(@("Move and Enable", "Move and Disable", "Move Only", "Enable Only", "Delete - AD Only", "Delete - MECM Only", "Delete - MECM & AD", "Analyze"))
$comboBoxAction.SelectedIndex = 0
$comboBoxAction.DropDownStyle = 'DropDownList'
$comboBoxAction.Size = New-Object System.Drawing.Size(140, 30)
$comboBoxAction.Location = New-Object System.Drawing.Point(65, 607)
$form.Controls.Add($comboBoxAction)

$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = "Run"
$buttonRun.Size = New-Object System.Drawing.Size(100, 30)
$buttonRun.Location = New-Object System.Drawing.Point(220, 607)
$form.Controls.Add($buttonRun)

# Move Copy and Clear buttons to the right
$buttonCopyLog.Location = New-Object System.Drawing.Point(330, 607)
$buttonClearResults.Location = New-Object System.Drawing.Point(440, 607)

# Update Update-RunButtonState to account for action
function Update-RunButtonState {
    $action = $comboBoxAction.SelectedItem
    # Disable OU selection for Enable Only, Delete, and Analyze
    if (($action -eq "Enable Only") -or ($action -eq "Delete - AD Only") -or ($action -eq "Delete - MECM Only") -or ($action -eq "Delete - MECM & AD") -or ($action -eq "Analyze")) {
        $comboBoxOU.Enabled = $false
        $labelOU.Enabled = $false
    } else {
        $comboBoxOU.Enabled = $true
        $labelOU.Enabled = $true
    }
    # Run button is always enabled and blue (or red for Delete)
    $buttonRun.Enabled = $true
    $buttonRun.ForeColor = [System.Drawing.Color]::White
    if (($action -eq "Delete - AD Only") -or ($action -eq "Delete - MECM Only") -or ($action -eq "Delete - MECM & AD")) {
        $buttonRun.BackColor = [System.Drawing.Color]::DarkRed
    } else {
        $buttonRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    }
}

$comboBoxOU.Add_TextChanged({ Update-RunButtonState })
$comboBoxOU.Add_SelectedIndexChanged({ Update-RunButtonState })
$textBoxComputers.Add_TextChanged({ Update-RunButtonState })
$comboBoxAction.Add_SelectedIndexChanged({ Update-RunButtonState })

# Call once at startup to set initial state
Update-RunButtonState

$buttonRun.Add_Click({
    $ou = $comboBoxOU.Text.Trim()
    $computerNames = $textBoxComputers.Lines | Where-Object { $_.Trim() -ne "" }
    $action = $comboBoxAction.SelectedItem

    # Validation logic for each action
    if (($action -eq "Move and Enable") -or ($action -eq "Move Only") -or ($action -eq "Move and Disable")) {
        if (-not $ou -or $ou -eq $placeholderOU) {
            Show-ErrorPopup "Please enter or select a valid target OU."
            return
        }
        if (-not $computerNames) {
            Show-ErrorPopup "Please enter or import at least one computer name."
            return
        }
    } elseif (($action -eq "Analyze") -or ($action -eq "Enable Only") -or ($action -eq "Delete - AD Only") -or ($action -eq "Delete - MECM Only") -or ($action -eq "Delete - MECM & AD")) {
        if (-not $computerNames) {
            Show-ErrorPopup "Please enter or import at least one computer name."
            return
        }
    }

    # Check MECM availability for Delete - MECM action
    if (($action -eq "Delete - MECM Only" -or $action -eq "Delete - MECM & AD") -and -not $cmModuleAvailable) {
        Show-ErrorPopup "Configuration Manager Console (CMConsole) is not available.`nCannot perform MECM operations.`nPlease install CMConsole and restart the application."
        return
    }

    # Confirmation for Delete actions
    if ($action -eq "Delete - AD Only") {
        $pcList = $computerNames -join "`n"
        $msg = "You are about to delete the following PCs from Active Directory:`n`n$pcList`n`nAre you sure?"
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    } elseif ($action -eq "Delete - MECM Only") {
        $pcList = $computerNames -join "`n"
        $msg = "You are about to delete the following PCs from MECM/Configuration Manager:`n`n$pcList`n`nAre you sure?"
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm MECM Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    } elseif ($action -eq "Delete - MECM & AD") {
        $pcList = $computerNames -join "`n"
        $msg = "You are about to delete the following PCs from BOTH MECM and Active Directory:`n`n$pcList`n`nAre you sure?"
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Combined Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    foreach ($name in $computerNames) {
        if ($action -eq "Delete - MECM Only") {
            # Handle MECM-only deletion without AD lookup
            if (-not $cmModuleAvailable) {
                Write-Log "✗ Failed: $name — CMConsole not available. Install Configuration Manager Console." $true
            } else {
                # Set location to CM drive
                $currentLocation = Get-Location
                Set-Location "$($SiteCode):\"
                
                try {
                    # Check if device exists in collection
                    $device = Get-CMDevice -Name $name.Trim() -Fast -CollectionID $CollectionId -ErrorAction SilentlyContinue
                    if ($null -eq $device) {
                        Write-Log "✗ Failed: $name — Not Found in MECM." $true
                    } else {
                        try {
                            Remove-CMDevice -Name $name.Trim() -Force -ErrorAction Stop
                            # Verify removal
                            $deviceCheck = Get-CMDevice -Name $name.Trim() -Fast -CollectionID $CollectionId -ErrorAction SilentlyContinue
                            if ($null -eq $deviceCheck) {
                                Write-Log "✓ Success: $name removed from MECM."
                            } else {
                                Write-Log "✗ Failed: $name — MECM removal failed." $true
                            }
                        } catch {
                            Write-Log "✗ Failed: $name — MECM removal error: $($_.Exception.Message)" $true
                        }
                    }
                } finally {
                    # Restore original location
                    Set-Location $currentLocation
                }
            }
        } elseif ($action -eq "Delete - MECM & AD") {
            # Handle combined MECM and AD deletion
            $mecmResult = ""
            $adResult = ""
            
            # Try MECM deletion first
            if (-not $cmModuleAvailable) {
                $mecmResult = "MECM unavailable"
            } else {
                $currentLocation = Get-Location
                Set-Location "$($SiteCode):\"
                
                try {
                    $device = Get-CMDevice -Name $name.Trim() -Fast -CollectionID $CollectionId -ErrorAction SilentlyContinue
                    if ($null -eq $device) {
                        $mecmResult = "not found in MECM"
                    } else {
                        try {
                            Remove-CMDevice -Name $name.Trim() -Force -ErrorAction Stop
                            $deviceCheck = Get-CMDevice -Name $name.Trim() -Fast -CollectionID $CollectionId -ErrorAction SilentlyContinue
                            if ($null -eq $deviceCheck) {
                                $mecmResult = "deleted from MECM"
                            } else {
                                $mecmResult = "MECM deletion failed"
                            }
                        } catch {
                            $mecmResult = "MECM error: $($_.Exception.Message)"
                        }
                    }
                } finally {
                    Set-Location $currentLocation
                }
            }
            
            # Try AD deletion
            try {
                $comp = Get-ADComputer -Filter "Name -eq '$name'" -Properties DistinguishedName,Enabled -ErrorAction Stop
                if ($comp) {
                    # Parse domain from DN
                    $dn = $comp.DistinguishedName
                    $domainMatch = [regex]::Match($dn, "DC=.*").Value
                    $domainName = ($domainMatch -replace "DC=", "") -replace ",", "."
                    # Build Disposed OU path
                    $disposedPath = $DisposedPathPrefix + "," + $domainMatch
                    # Get all DCs
                    $dcs = (Get-ADDomain -Server $domainName).ReplicaDirectoryServers
                    $removalCounter = 0
                    foreach ($dc in $dcs) {
                        try {
                            $adComp = Get-ADComputer -Identity $comp.Name -Server $dc -ErrorAction Stop
                            if ($adComp -and ($adComp.DistinguishedName -notmatch [regex]::Escape($DisposedPathPrefix))) {
                                Move-ADObject -Identity $adComp.DistinguishedName -Server $dc -TargetPath $disposedPath -Confirm:$false -ErrorAction Stop
                                Start-Sleep -Seconds 3
                                $adComp = Get-ADComputer -Identity $comp.Name -Server $dc -ErrorAction Stop
                            }
                            Set-ADObject -Identity $adComp.DistinguishedName -Server $dc -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
                            Remove-ADObject -Identity $adComp.DistinguishedName -Server $dc -Recursive -Confirm:$false -ErrorAction Stop
                            $removalCounter++
                        } catch {
                            # Continue trying other DCs
                        }
                    }
                    if ($removalCounter -gt 0) {
                        $adResult = "deleted from AD"
                    } else {
                        $adResult = "AD deletion failed"
                    }
                } else {
                    $adResult = "not found in AD"
                }
            } catch {
                $adResult = "AD error: $($_.Exception.Message)"
            }
            
            # Combine results and log appropriate message
            if (($mecmResult -eq "deleted from MECM") -and ($adResult -eq "deleted from AD")) {
                Write-Log "✓ Success: $name deleted from both MECM and AD."
            } elseif (($mecmResult -eq "deleted from MECM") -and ($adResult -like "*not found*")) {
                Write-Log "✓ Partial: $name deleted from MECM only ($adResult)." -isWarning $true
            } elseif (($adResult -eq "deleted from AD") -and ($mecmResult -like "*not found*")) {
                Write-Log "✓ Partial: $name deleted from AD only ($mecmResult)." -isWarning $true
            } elseif (($mecmResult -like "*not found*") -and ($adResult -like "*not found*")) {
                Write-Log "✗ Failed: $name — Not found in MECM or AD." $true
            } elseif (($mecmResult -eq "deleted from MECM") -and ($adResult -like "*failed*" -or $adResult -like "*error*")) {
                Write-Log "✓ Partial: $name deleted from MECM only. AD issue: $adResult." -isWarning $true
            } elseif (($adResult -eq "deleted from AD") -and ($mecmResult -like "*failed*" -or $mecmResult -like "*error*")) {
                Write-Log "✓ Partial: $name deleted from AD only. MECM issue: $mecmResult." -isWarning $true
            } else {
                Write-Log "✗ Failed: $name — MECM: $mecmResult, AD: $adResult" $true
            }
        } elseif ($action -eq "Analyze") {
            # Handle independent analysis of both AD and MECM
            
            # Check AD status
            try {
                $comp = Get-ADComputer -Filter "Name -eq '$name'" -Properties Enabled, DistinguishedName -ErrorAction Stop
                if ($comp) {
                    $adStatus = if ($comp.Enabled) { "Enabled" } else { "Disabled" }
                    $cleanOU = Get-OUFromDN -distinguishedName $comp.DistinguishedName
                    Write-Log "[Analyze] $name — AD: $adStatus — $cleanOU"
                } else {
                    Write-Log "[Analyze] $name — AD: Not found" $true
                }
            } catch {
                Write-Log "[Analyze] $name — AD: Error - $($_.Exception.Message)" $true
            }
            
            # Check MECM status
            if (-not $cmModuleAvailable) {
                Write-Log "[Analyze] $name — MECM: Console unavailable" $true
            } else {
                $currentLocation = Get-Location
                Set-Location "$($SiteCode):\"
                
                try {
                    $device = Get-CMDevice -Name $name.Trim() -Fast -CollectionID $CollectionId -ErrorAction SilentlyContinue
                    if ($null -eq $device) {
                        Write-Log "[Analyze] $name — MECM: Not found" $true
                    } else {
                        Write-Log "[Analyze] $name — MECM: Found in collection $CollectionId"
                    }
                } catch {
                    Write-Log "[Analyze] $name — MECM: Error - $($_.Exception.Message)" $true
                } finally {
                    Set-Location $currentLocation
                }
            }
            
            # Add visual separator between computers
            Write-Log ""
        } else {
            # Handle all other actions that require AD lookup
            try {
                $comp = Get-ADComputer -Filter "Name -eq '$name'" -Properties Enabled, DistinguishedName -ErrorAction Stop
                if ($comp) {
                    switch ($action) {
                        "Move and Enable" {
                            Move-ADObject -Identity $comp.DistinguishedName -TargetPath $ou
                            Start-Sleep -Seconds 3
                            $compNew = Get-ADComputer -Filter "Name -eq '$name'" -ErrorAction Stop
                            Enable-ADAccount -Identity $compNew.DistinguishedName
                            Write-Log "✓ Success: $name moved and enabled."
                        }
                        "Move Only" {
                            Move-ADObject -Identity $comp.DistinguishedName -TargetPath $ou
                            Write-Log "✓ Success: $name moved."
                        }
                        "Move and Disable" {
                            Move-ADObject -Identity $comp.DistinguishedName -TargetPath $ou
                            Start-Sleep -Seconds 3
                            $compNew = Get-ADComputer -Filter "Name -eq '$name'" -ErrorAction Stop
                            Disable-ADAccount -Identity $compNew.DistinguishedName
                            Write-Log "✓ Success: $name moved and disabled."
                        }
                        "Enable Only" {
                            Enable-ADAccount -Identity $comp.DistinguishedName
                            Write-Log "✓ Success: $name enabled."
                        }
                        "Delete - AD Only" {
                            $comp = Get-ADComputer -Filter "Name -eq '$name'" -Properties DistinguishedName,Enabled -ErrorAction Stop
                            if ($comp) {
                                # Parse domain from DN
                                $dn = $comp.DistinguishedName
                                $domainMatch = [regex]::Match($dn, "DC=.*").Value
                                $domainName = ($domainMatch -replace "DC=", "") -replace ",", "."
                                # Build Disposed OU path
                                $disposedPath = $DisposedPathPrefix + "," + $domainMatch
                                # Get all DCs
                                $dcs = (Get-ADDomain -Server $domainName).ReplicaDirectoryServers
                                $removalCounter = 0
                                foreach ($dc in $dcs) {
                                    try {
                                        $adComp = Get-ADComputer -Identity $comp.Name -Server $dc -ErrorAction Stop
                                        if ($adComp -and ($adComp.DistinguishedName -notmatch [regex]::Escape($DisposedPathPrefix))) {
                                            Move-ADObject -Identity $adComp.DistinguishedName -Server $dc -TargetPath $disposedPath -Confirm:$false -ErrorAction Stop
                                            Start-Sleep -Seconds 3
                                            $adComp = Get-ADComputer -Identity $comp.Name -Server $dc -ErrorAction Stop
                                        }
                                        Set-ADObject -Identity $adComp.DistinguishedName -Server $dc -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
                                        Remove-ADObject -Identity $adComp.DistinguishedName -Server $dc -Recursive -Confirm:$false -ErrorAction Stop
                                        $removalCounter++
                                    } catch {
                                        Write-Log "✗ Failed: $name on $dc — $($_.Exception.Message)" $true
                                    }
                                }
                                if ($removalCounter -gt 0) {
                                    Write-Log "✓ Success: $name deleted from $removalCounter DC(s)."
                                } else {
                                    Write-Log "✗ Failed: $name — Could not be deleted from any DC." $true
                                }
                            } else {
                                Write-Log "✗ Failed: $name — Not found in AD." $true
                            }
                        }
                    }
                } else {
                    Write-Log "✗ Failed: $name — Not found in AD." $true
                }
            }
            catch {
                Write-Log "✗ Failed: $name — $($_.Exception.Message)" $true
            }
        }
    }
    [System.Windows.Forms.MessageBox]::Show("Operation complete.","Done",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
})

$buttonCopyLog.Add_Click({
    if ($logBox.Text.Trim()) {
        [System.Windows.Forms.Clipboard]::SetText($logBox.Text)
        [System.Windows.Forms.MessageBox]::Show("Results copied to clipboard.","Copied",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("No results to copy.","Nothing to Copy",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$buttonClearResults.Add_Click({
    $logBox.Clear()
})

# Draw border in form's Paint event
$form.Add_Paint({
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(50, 50, 50), 1)
    $_.Graphics.DrawRectangle($pen, 0, 0, $form.Width - 1, $form.Height - 1)
    $pen.Dispose()
})

[void]$form.ShowDialog()
Stop-Transcript | Out-Null
