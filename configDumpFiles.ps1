
# configDumpFiles.ps1 by David Sirrine @ CNS Computer Services, LLC
# latest revision 2024-04-22

# Sets the Crash (BSoD) settings to create full, kernel-only, small, or
# no memory dumps and keep all or only the last memory dump.
#
# Defaults to "Complete" for debugging purposes,
# and "overwrite" to save disk space.


#default dump level "Complete"
if {-not $dumpLevel) { $dumpLevel = "Complete" }

#hash to decode human-readable dump levels
$lvlHash = @{
    "None"     = 0
    "Complete" = 1
    "Kernel"   = 2
    "Small"    = 3
}
$lvl = $lvlHash[$dumpLevel]

#set the crash memory dump level
if (-not $lvl) {
    Write-Host "Invalid dump level specified."
} else {
	Write-Host "$dumpLevel crash dump files will be saved."
    Get-WmiObject -Class Win32_OSRecoveryConfiguration -EnableAllPrivileges | Set-WmiInstance -Arguments @{ DebugInfoType=$lvl }
}

#overwrite memory dump files unless specified
if ($overwriteDebugFiles="no") {
	Write-Host "ALL dump files will be retained."
	$CrashBehaviour | Set-WmiInstance -Arguments @{ OverwriteExistingDebugFile=$False }
	###warn if "Complete" dump files will be retained
	if ($lvl=1) { Write-Host "WARNING: Be aware that `"Complete`" dump files can quickly fill the drive!" }
} else {
	Write-Host "only the last dump file will be saved."
	$CrashBehaviour | Set-WmiInstance -Arguments @{ OverwriteExistingDebugFile=$True }
}
