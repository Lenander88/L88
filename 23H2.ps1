Write-Host -BackgroundColor Black -ForegroundColor Green "Start OSDCloud ZTI"
Start-Sleep -Seconds 5

Add-Type -AssemblyName PresentationFramework
$bodyMessage = [PSCustomObject] @{}; Clear-Variable serialNumber -ErrorAction:SilentlyContinue
$serialNumber = Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber

if ($serialNumber) {

    $bodyMessage | Add-Member -MemberType NoteProperty -Name "serialNumber" -Value $serialNumber

} else {

    $infoMessage = "We were unable to locate the serial number of your device, so the process cannot proceed. The computer will shut down when this window is closed."
    Write-Host -BackgroundColor Black -ForegroundColor Red $infoMessage
    [System.Windows.MessageBox]::Show($infoMessage, 'OSDCloud', 'OK', 'Error') | Out-Null
    wpeutil shutdown
}

###################Load Assembly for creating form & button######
  
[void][System.Reflection.Assembly]::LoadWithPartialName( "System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName( "Microsoft.VisualBasic")
  
  
#####Define the form size & placement
  
$form = New-Object "System.Windows.Forms.Form";
$form.Width = 500;
$form.Height = 150;
$form.Text = "Digital Workplace Group Tag";
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$form.ControlBox = $True
  
  
##############Define text label2
  
$textLabel2 = New-Object "System.Windows.Forms.Label";
$textLabel2.Left = 25;
$textLabel2.Top = 45;
$textLabel2.Text = "Group tag";
  
  
############Define text box2 for input
  
$cBox2 = New-Object "System.Windows.Forms.combobox";
$cBox2.Left = 150;
$cBox2.Top = 45;
$cBox2.width = 200;
$cBox2.Text = "Choose group tag"
  
  
###############"Add descriptions to combo box"##############

Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Lenander88/L88/main/grouptags.csv' -Outfile grouptags.csv 
Import-CSV ".\grouptags.csv" | ForEach-Object {
    $cBox2.Items.Add($_.grouptags)
      
}
  
  
#############define OK button
$button = New-Object "System.Windows.Forms.Button";
$button.Left = 360;
$button.Top = 45;
$button.Width = 100;
$button.Text = "OK";
$Button.Cursor = [System.Windows.Forms.Cursors]::Hand
#$Button.Font = New-Object System.Drawing.Font("Times New Roman",10,[System.Drawing.FontStyle]::Regular)
############# This is when you have to close the form after getting values
$eventHandler = [System.EventHandler]{
$cBox2.Text;
$form.Close();};
$button.Add_Click($eventHandler) ;
  
#############Add controls to all the above objects defined
$form.Controls.Add($button);
$form.Controls.Add($textLabel2);
$form.Controls.Add($cBox2);
#$ret = $form.ShowDialog();
  
#################return values
$button.add_Click({
      
    #Set-Variable -Name locationResult -Value $combobox1.selectedItem -Force -Scope Script # Use this
    $script:locationResult = $cBox2.selectedItem # or this to retrieve the user selection
})
  
$form.Controls.Add($button)
$form.Controls.Add($cBox2)
  
$form.ShowDialog()
  
$grouptag = $script:locationResult
Write-Output $grouptag

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
#    wpeutil shutdown
    
Write-Host -BackgroundColor Black -ForegroundColor Yellow -NoNewLine 'Press any key to continue...'
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

} else {

    Write-Host -BackgroundColor Black -ForegroundColor Green "Update OSD PowerShell Module"
    Install-Module OSD -Force -SkippublisherCheck

    Write-Host -BackgroundColor Black -ForegroundColor Green "Import OSD PowerShell Module"
    Import-Module OSD -Force

    Write-Host -BackgroundColor Black -ForegroundColor Green "Start OSDCloud"
    Start-OSDCloud -ZTI -OSVersion 'Windows 11' -OSBuild 23H2 -OSEdition Enterprise -OSLanguage en-us -OSLicense Retail

    Write-Host -BackgroundColor Black -ForegroundColor Green "Restart in 20 seconds"
    Start-Sleep -Seconds 20
    wpeutil reboot
}