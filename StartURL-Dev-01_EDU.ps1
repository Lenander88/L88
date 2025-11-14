#================================================
#   [PreOS] Perimeters
#================================================
$Params = @{
    OSVersion = "Windows 11"
    OSBuild = "24H2"
    OSEdition = "Enterprise"
    OSLanguage = "en-us"
    OSLicense = "Retail"
    ZTI = $true
    Firmware = $false
}

# Group all Add-Type calls together for clarity
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

#=======================================================================
#   [PreOS] EDU Build Selection
#=======================================================================
Write-Host -BackgroundColor Black -ForegroundColor Green "Starting EDU Build Selection"
Start-Sleep -Seconds 5

$bodyMessage = [PSCustomObject] @{}
Clear-Variable serialNumber -ErrorAction:SilentlyContinue
$serialNumber = Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber

# Path to CSV file
$csvPath = ".\EDU.csv"
$csvUrl = 'https://raw.githubusercontent.com/Lenander88/L88/main/EDU_Dev.csv'
if (!(Test-Path $csvPath) -or ((Get-Item $csvPath).LastWriteTime -lt (Get-Date).AddDays(-1))) {
    Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath
}

$options = Import-CSV $csvPath

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select EDU Build"
$form.Size = New-Object System.Drawing.Size(300,150)
$form.StartPosition = "CenterScreen"

$textLabel2 = New-Object "System.Windows.Forms.Label";
$textLabel2.Left = 25;
$textLabel2.Top = 45;
$textLabel2.Text = "EDU Build";

# Create ComboBox
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(50,20)
$comboBox.Size = New-Object System.Drawing.Size(200,20)
$comboBox.DropDownStyle = 'DropDownList'  # Prevent typing, only select

# Populate ComboBox with OptionName from CSV
foreach ($item in $options) {
    $comboBox.Items.Add($item.OptionName)
}

$form.Controls.Add($comboBox)

# Create OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100,60)

$okButtonClickHandler = {
    $selectedOption = $comboBox.SelectedItem
    if ($selectedOption) {
        # Assign corresponding Value to $edu (hidden from user)
        $global:edu = ($options | Where-Object { $_.OptionName -eq $selectedOption }).Value
        $form.Close()
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select an option.")
    }
}

$okButton.Add_Click($okButtonClickHandler)

$form.Controls.Add($okButton)

# Show Form
$form.ShowDialog()


#=======================================================================
#   [PreOS] Detect Serial Number and Prepare for AutoPilot
#=======================================================================
if ($serialNumber) {

    $bodyMessage | Add-Member -MemberType NoteProperty -Name "serialNumber" -Value $serialNumber

} else {

    $infoMessage = "We were unable to locate the serial number of your device, so the process cannot proceed. The computer will shut down when this window is closed."
    Write-Host -BackgroundColor Black -ForegroundColor Red $infoMessage
    [System.Windows.MessageBox]::Show($infoMessage, 'OSDCloud', 'OK', 'Error') | Out-Null
    wpeutil shutdown
}

#=======================================================================
#   [PreOS] AutoPilot Verification and Group Tag Selection
#=======================================================================

Write-Host -BackgroundColor Black -ForegroundColor Green "Start AutoPilot Verification"
$body = $bodyMessage | ConvertTo-Json -Depth 5; $uri = 'https://prod-145.westus.logic.azure.com:443/workflows/dadfcaca1bcc4b069c998a99e82ee728/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=n0urWoGWa2OXN-4ba0U7UwfEM8i9vwTuSHx2PrSVtvU'
$result = Invoke-RestMethod -Uri $uri -Method POST -Body $body -ContentType "application/json; charset=utf-8" -UseBasicParsing    

if ($result.Response -eq 0) {

    Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/PCPKsp.dll' -OutFile X:\Windows\System32\PCPKsp.dll
    rundll32 X:\Windows\System32\PCPKsp.dll,DllInstall

    Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/OA3.cfg' -OutFile OA3.cfg
    Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/oa3tool.exe' -OutFile oa3tool.exe
    Remove-Item .\OA3.xml -ErrorAction:SilentlyContinue
    .\oa3tool.exe /Report /ConfigFile=.\OA3.cfg /NoKeyCheck

    if (Test-Path .\OA3.xml) {        
        # Path to CSV file
        $csvPath = ".\grouptags.csv"
        $csvUrl = 'https://raw.githubusercontent.com/Lenander88/L88/main/grouptags_dev.csv'
        if (!(Test-Path $csvPath) -or ((Get-Item $csvPath).LastWriteTime -lt (Get-Date).AddDays(-1))) {
            Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath
        }

        $options = Import-CSV $csvPath

        # Create Form
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Digital Workplace Group Tag"
        $form.Size = New-Object System.Drawing.Size(300,150)
        $form.StartPosition = "CenterScreen"
        $form.ControlBox = $True

        $textLabel2 = New-Object "System.Windows.Forms.Label";
        $textLabel2.Left = 25;
        $textLabel2.Top = 45;
        $textLabel2.Text = "Digital Workplace Group Tag";

        # Create ComboBox
        $comboBox = New-Object System.Windows.Forms.ComboBox
        $comboBox.Location = New-Object System.Drawing.Point(50,20)
        $comboBox.Size = New-Object System.Drawing.Size(200,20)
        $comboBox.DropDownStyle = 'DropDownList'  # Prevent typing, only select

        # Populate ComboBox with OptionName from CSV
        foreach ($item in $options) {
            $comboBox.Items.Add($item.OptionName)
        }

        $form.Controls.Add($comboBox)

        # Create OK Button
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(100,60)

        $okButtonClickHandler = {
            $selectedOption = $comboBox.SelectedItem
            if ($selectedOption) {
                # Assign corresponding Value to $grouptag (hidden from user)
                $global:grouptag = ($options | Where-Object { $_.OptionName -eq $selectedOption }).Value
                $form.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please select an option.")
            }
        }

        $okButton.Add_Click($okButtonClickHandler)

        $form.Controls.Add($okButton)

        # Show Form
        $form.ShowDialog()

        [xml]$xmlhash = Get-Content -Path .\OA3.xml
        $hash=$xmlhash.Key.HardwareHash

        $computers = @(); $product = ""

        $c = New-Object psobject -Property @{
            "Device Serial Number" = $serialNumber
            "Windows Product ID" = $product
            "Hardware Hash" = $hash
            "Group Tag" = $grouptag
        }

        $computers += $c
        $computers | Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" | ConvertTo-CSV -NoTypeInformation | ForEach-Object {$_ -replace '"',''} | Out-File AutopilotHWID.csv
        
        $usbMedia = Get-WmiObject -Namespace "root\cimv2" -Query "SELECT * FROM Win32_LogicalDisk WHERE DriveType = 2"
        foreach ($disk in $usbMedia) {
            Copy-Item -Path .\AutopilotHWID.csv -Destination "$($disk.DeviceID)\$($serialNumber).csv" -Force -ErrorAction:SilentlyContinue
        }
        Copy-Item -Path .\AutopilotHWID.csv -Destination "C:\$($serialNumber).csv" -Force -ErrorAction:SilentlyContinue
    }


    $infoMessage = "You cannot continue because the device is not ready for Windows AutoPilot. The HWHash has been generated and placed on the USB-stick, upload HWHash, reinsert USB-stick and click OK to start deployment."
    Write-Host -BackgroundColor Black -ForegroundColor Yellow $infoMessage
    [System.Windows.MessageBox]::Show($infoMessage, 'HWHash', 'OK', 'Warning') | Out-Null
 
#=======================================================================
#   [OS] Start-OSDCloud and update OSD Module
#=======================================================================
    Write-Host -BackgroundColor Black -ForegroundColor Green "Updating OSD PowerShell Module"
    Install-Module OSD -Force

    Write-Host -BackgroundColor Black -ForegroundColor Green "Importing OSD PowerShell Module"
    Import-Module OSD -Force   


    Write-Host -BackgroundColor Black -ForegroundColor Green "Start OSDCloud"
    Start-OSDCloud @Params
#================================================
#  [PostOS] SetupComplete CMD Command Line
#================================================

Write-Host -BackgroundColor Black -ForegroundColor Green "Stage SetupComplete"

# Ensure PSWindowsUpdate is staged for post-boot use
Save-Module -Name PSWindowsUpdate -Path 'C:\Program Files\WindowsPowerShell\Modules' -Force

# Ensure SetupComplete folder exists
$setupPath = 'C:\OSDCloud\Scripts\SetupComplete'
if (-not (Test-Path $setupPath)) {
    New-Item -Path $setupPath -ItemType Directory -Force | Out-Null
}

# Run the selected EDU command (from ComboBox selection)
Invoke-Expression $edu
# Download SetupComplete.cmd
Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/SetupComplete.cmd' `
    -OutFile "$setupPath\SetupComplete.cmd"

# Download Install-LCU.ps1
Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/Install-LCU.ps1' `
    -OutFile "$setupPath\Install-LCU.ps1"

#=======================================================================
#   Restart-Computer
#=======================================================================   
    Write-Host -BackgroundColor Black -ForegroundColor Green "Restart in 20 seconds"
    Start-Sleep -Seconds 20
    wpeutil reboot

} else {

#=======================================================================
#   [OS] Start-OSDCloud and update OSD Module
#=======================================================================
    Write-Host -BackgroundColor Black -ForegroundColor Green "Updating OSD PowerShell Module"
    Install-Module OSD -Force

    Write-Host -BackgroundColor Black -ForegroundColor Green "Importing OSD PowerShell Module"
    Import-Module OSD -Force

    Write-Host -BackgroundColor Black -ForegroundColor Green "Start OSDCloud"
    Start-OSDCloud @Params

#================================================
#  [PostOS] SetupComplete CMD Command Line
#================================================

Write-Host -BackgroundColor Black -ForegroundColor Green "Stage SetupComplete"

# Ensure PSWindowsUpdate is staged for post-boot use
Save-Module -Name PSWindowsUpdate -Path 'C:\Program Files\WindowsPowerShell\Modules' -Force

# Ensure SetupComplete folder exists
$setupPath = 'C:\OSDCloud\Scripts\SetupComplete'
if (-not (Test-Path $setupPath)) {
    New-Item -Path $setupPath -ItemType Directory -Force | Out-Null
}

# Run the selected EDU command (from ComboBox selection)
Invoke-Expression $edu

# Download SetupComplete.cmd
Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/SetupComplete.cmd' `
    -OutFile "$setupPath\SetupComplete.cmd"

# Download Install-LCU.ps1
Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/Install-LCU.ps1' `
    -OutFile "$setupPath\Install-LCU.ps1"
#=======================================================================
#   Restart-Computer
#=======================================================================    
    Write-Host -BackgroundColor Black -ForegroundColor Green "Restart in 20 seconds"
    Start-Sleep -Seconds 20
    wpeutil reboot
}
