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

#=======================================================================
#   [PreOS] EDU Build Selection
#=======================================================================
Write-Host -BackgroundColor Black -ForegroundColor Green "Starting EDU Build Selection"
Start-Sleep -Seconds 5

Add-Type -AssemblyName PresentationFramework
$bodyMessage = [PSCustomObject] @{}; Clear-Variable serialNumber -ErrorAction:SilentlyContinue
$serialNumber = Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber


[void][System.Reflection.Assembly]::LoadWithPartialName( "System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName( "Microsoft.VisualBasic")

    $form = New-Object "System.Windows.Forms.Form";
    $form.Width = 500;
    $form.Height = 150;
    $form.Text = "EDU Build";
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $form.ControlBox = $True

    $textLabel2 = New-Object "System.Windows.Forms.Label";
    $textLabel2.Left = 25;
    $textLabel2.Top = 45;
    $textLabel2.Text = "EDU Build";

    $cBox2 = New-Object "System.Windows.Forms.combobox";
    $cBox2.Left = 150;
    $cBox2.Top = 45;
    $cBox2.width = 200;
    $cBox2.Text = "EDU Build"

    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Lenander88/L88/main/EDU.csv' -Outfile EDU.csv 

# Create a hashtable to map names to URLs
    $eduMap = @{}

    Import-Csv ".\EDU.csv" | ForEach-Object {
    $eduMap[$_.EDU] = $_.Command
    $cBox2.Items.Add($_.EDU) | Out-Null
    }



    $button = New-Object "System.Windows.Forms.Button";
    $button.Left = 360;
    $button.Top = 45;
    $button.Width = 100;
    $button.Text = "OK";
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $eventHandler = [System.EventHandler]{
    $cBox2.Text;
    $form.Close();};
    $button.Add_Click($eventHandler) ;

    $form.Controls.Add($button);
    $form.Controls.Add($textLabel2);
    $form.Controls.Add($cBox2);


    $button.add_Click({
        $selectedName = $cBox2.SelectedItem
        $script:locationResult = $eduMap[$selectedName]
        $form.Close()
    })


    $form.Controls.Add($button)
    $form.Controls.Add($cBox2)

    $form.ShowDialog()

    $EDU = $script:locationResult
    Write-Host "Selected command: $EDU"

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
$body = $bodyMessage | ConvertTo-Json -Depth 5; $uri = 'https://defaultf0bdc1c951484f86ac40edd976e181.4c.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/dadfcaca1bcc4b069c998a99e82ee728/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=CZt0ePxAkBD147HaaTMZLxjZ9SuByOfVf-RYc5Ckl14'
$result = Invoke-RestMethod -Uri $uri -Method POST -Body $body -ContentType "application/json; charset=utf-8" -UseBasicParsing    

if ($result.Response -eq 0) {

    Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/PCPKsp.dll' -OutFile X:\Windows\System32\PCPKsp.dll
    rundll32 X:\Windows\System32\PCPKsp.dll,DllInstall

    Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/OA3.cfg' -OutFile OA3.cfg
    Invoke-WebRequest -Uri 'https://github.com/Lenander88/L88/raw/main/oa3tool.exe' -OutFile oa3tool.exe
    Remove-Item .\OA3.xml -ErrorAction:SilentlyContinue
    .\oa3tool.exe /Report /ConfigFile=.\OA3.cfg /NoKeyCheck

    if (Test-Path .\OA3.xml) {
        [void][System.Reflection.Assembly]::LoadWithPartialName( "System.Windows.Forms")
        [void][System.Reflection.Assembly]::LoadWithPartialName( "Microsoft.VisualBasic")
        
        $form = New-Object "System.Windows.Forms.Form";
        $form.Width = 500;
        $form.Height = 150;
        $form.Text = "Digital Workplace Group Tag";
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $form.ControlBox = $True

        $textLabel2 = New-Object "System.Windows.Forms.Label";
        $textLabel2.Left = 25;
        $textLabel2.Top = 45;
        $textLabel2.Text = "Group tag";
        
        $cBox2 = New-Object "System.Windows.Forms.combobox";
        $cBox2.Left = 150;
        $cBox2.Top = 45;
        $cBox2.width = 200;
        $cBox2.Text = "Choose group tag"

        Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Lenander88/L88/main/grouptags.csv' -Outfile grouptags.csv 
        Import-CSV ".\grouptags.csv" | ForEach-Object {
            $cBox2.Items.Add($_.grouptags)| out-null
            
        }

        $button = New-Object "System.Windows.Forms.Button";
        $button.Left = 360;
        $button.Top = 45;
        $button.Width = 100;
        $button.Text = "OK";
        $Button.Cursor = [System.Windows.Forms.Cursors]::Hand

        $eventHandler = [System.EventHandler]{
        $cBox2.Text;
        $form.Close();};
        $button.Add_Click($eventHandler) ;

        $form.Controls.Add($button);
        $form.Controls.Add($textLabel2);
        $form.Controls.Add($cBox2);

        $button.add_Click({    

            $script:locationResult = $cBox2.selectedItem 
        })
  
        $form.Controls.Add($button)
        $form.Controls.Add($cBox2)
  
        $form.ShowDialog()
  
        $grouptag = $script:locationResult
        Write-Host $grouptag

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
Invoke-Expression $EDU

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
Invoke-Expression $EDU

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
