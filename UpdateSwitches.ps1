#----------------------------------------
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
Unblock-File C:\_BUILD\CheckAppsAndInstallLatest2.ps1
Update-Evergreen
#----------------------------------------

#REPORTING
#Report – Displays versions of all defined apps
.\CheckAppsAndInstallLatest2.ps1 -NoReports

#Report only in powershell console, creates HTML report
.\CheckAppsAndInstallLatest2.ps1

#Report only HTML when all apps are green (Update)
.\CheckAppsAndInstallLatest2.ps1 -HtmlOnlyWhenGreen

#UPDATE SWITCHES

# Update – TEST RUN Apps and Windows updates
.\CheckAppsAndInstallLatest2.ps1 -Upgrade -IncludeWindowsUpdate -WhatIf

# Update – Full run: Windows OS patches + app upgrades
.\CheckAppsAndInstallLatest2.ps1 -Upgrade -IncludeWindowsUpdate -NoReports

# Update Installed Apps only
.\CheckAppsAndInstallLatest2.ps1 -Upgrade -NoReports

# Update – Windows OS only
.\CheckAppsAndInstallLatest2.ps1 -Upgrade -IncludeWindowsUpdate -WindowsUpdateOnly -NoReports

# Update – Disable the MStore repo, may display errors when trying to pull the latest Evergreen apps from Mstore, disabling defaults to winget only source
.\CheckAppsAndInstallLatest2.ps1 -Upgrade -HtmlOnlyWhenGreen -DisableMsStore

#Useful Modules
Install-Module -name PSWindowsUpdate -Force
Import-Module PSWindowsUpdate
