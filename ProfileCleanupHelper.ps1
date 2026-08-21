# (optional) allow scripts just for this session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 1) Run BOTH regions as a DRY RUN (safe preview)
& "\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-ProfileCleanups.ps1" -Region Both -DryRun

# 2) Run SOUTH only, show table, skip two folders, 45-day threshold
& "\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-ProfileCleanups.ps1" `
   -Region South `
   -DaysInactive 3 `
   -ExcludeFolders @('Profile_Redirections','Template') `
   -ShowTable 1 `
   -DryRun

   # 3) Run WEST only, show table, skip two folders, 45-day threshold
& "\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-ProfileCleanups.ps1" `
   -Region West `
   -DaysInactive 10 `
   -ExcludeFolders @('Profile_Redirections','Template') `
   -ShowTable 1 `
   -DryRun

# 3) Run WEST for real (no -DryRun), hide the table
& "\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-ProfileCleanups.ps1" -Region West -ShowTable 0
