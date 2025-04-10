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

(Write-Host -BackgroundColor Black -ForegroundColor Green "Start AutoPilot Verification"
$body = $bodyMessage | ConvertTo-Json -Depth 5; $uri = "https://prod-145.westus.logic.azure.com:443/workflows/dadfcaca1bcc4b069c998a99e82ee728/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=n0urWoGWa2OXN-4ba0U7UwfEM8i9vwTuSHx2PrSVtvU"
$result = Invoke-RestMethod -Uri $uri -Method POST -Body $body -ContentType "application/json; charset=utf-8" -UseBasicParsing    

##if ($result.Response -eq 0) {

    Invoke-WebRequest -Uri "https://github.com/dwp-lab/OSDCloud/raw/main/PCPKsp.dll" -OutFile X:\Windows\System32\PCPKsp.dll
    rundll32 X:\Windows\System32\PCPKsp.dll,DllInstall

    Invoke-WebRequest -Uri "https://github.com/Lenander88/L88/main/OA3.cfg" -OutFile OA3.cfg
    Invoke-WebRequest -Uri "https://github.com/Lenander88/L88/main/oa3tool.exe" -OutFile oa3tool.exe
    Remove-Item .\OA3.xml -ErrorAction:SilentlyContinue
    .\oa3tool.exe /Report /ConfigFile=.\OA3.cfg /NoKeyCheck

    Write-Host -BackgroundColor Black -ForegroundColor Yellow -NoNewLine 'Press any key to continue...';
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

##    if (Test-Path .\OA3.xml) {

##        [xml]$xmlhash = Get-Content -Path .\OA3.xml
##        $hash=$xmlhash.Key.HardwareHash

##        $computers = @(); $product = ""

##        $c = New-Object psobject -Property @{
##            "Device Serial Number" = $serialNumber
##            "Windows Product ID" = $product
##            "Hardware Hash" = $hash
##            "Group Tag" = 'MEM-Global-Standard'
##        }

##        $computers += $c
##        $computers | Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" | ConvertTo-CSV -NoTypeInformation | ForEach-Object {$_ -replace '"',''} | Out-File AutopilotHWID.csv
        
##        $usbMedia = Get-WmiObject -Namespace "root\cimv2" -Query "SELECT * FROM Win32_LogicalDisk WHERE DriveType = 2"
##        foreach ($disk in $usbMedia) {
##            Copy-Item -Path .\AutopilotHWID.csv -Destination "$($disk.DeviceID)\$($serialNumber).csv" -Force -ErrorAction:SilentlyContinue
##        }
##        Copy-Item -Path .\AutopilotHWID.csv -Destination "C:\$($serialNumber).csv" -Force -ErrorAction:SilentlyContinue
##    }

##    $infoMessage = "You cannot continue because the device is not ready for Windows AutoPilot. The computer will shut down when this window is closed."
##    Write-Host -BackgroundColor Black -ForegroundColor Red $infoMessage
##    [System.Windows.MessageBox]::Show($infoMessage, 'OSDCloud', 'OK', 'Error') | Out-Null
##    wpeutil shutdown
    
##} else {

    Write-Host -BackgroundColor Black -ForegroundColor Green "Update OSD PowerShell Module"
    Install-Module OSD -Force -SkippublisherCheck

    Write-Host -BackgroundColor Black -ForegroundColor Green "Import OSD PowerShell Module"
    Import-Module OSD -Force

    Write-Host -BackgroundColor Black -ForegroundColor Green "Start OSDCloud"
    Start-OSDCloud -ZTI -OSVersion 'Windows 11' -OSBuild 23H2 -OSEdition Enterprise -OSLanguage en-us -OSLicense Retail

    Write-Host -BackgroundColor Black -ForegroundColor Green "Restart in 20 seconds"
    Start-Sleep -Seconds 20
    wpeutil reboot
##}