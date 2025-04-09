Write-Host -BackgroundColor Black -ForegroundColor Green "Start OSDCloud ZTI"
Start-Sleep -Seconds 5
Start-OSDCloud -ZTI -OSVersion 'Windows 11' -OSBuild 23H2 -OSEdition Enterprise -OSLanguage en-us -OSLicense Retail

Write-Host -BackgroundColor Black -ForegroundColor Green "Update OSD PowerShell Module"
Install-Module OSD -Force -SkippublisherCheck

Write-Host -BackgroundColor Black -ForegroundColor Green "Import OSD PowerShell Module"
Import-Module OSD -Force

Write-Host -BackgroundColor Black -ForegroundColor Green "Start OSDCloud"
Start-OSDCloud -ZTI -OSVersion 'Windows 11' -OSBuild 23H2 -OSEdition Enterprise -OSLanguage en-us -OSLicense Retail

Write-Host -BackgroundColor Black -ForegroundColor Green "Restart in 20 seconds"
Start-Sleep -Seconds 20
wpeutil reboot