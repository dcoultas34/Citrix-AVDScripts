#Currently set to look at profiles that have not been used in 60 days

#Dry run South
\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-PruneFSLogix.ps1 -Region South
#Dry run West
\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-PruneFSLogix.ps1 -Region West
#Dry run for both profile shares
\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-PruneFSLogix.ps1


#Live deletion
#South
\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-PruneFSLogix.ps1 -Region South -Live -SkipIfLocked
#West
\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-PruneFSLogix.ps1 -Region West -Live -SkipIfLocked
#Both
\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Run-PruneFSLogix.ps1 -Live -SkipIfLocked







