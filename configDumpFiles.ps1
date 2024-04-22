
# configDumpFiles.ps1 by David Sirrine @ CNS Computer Services, LLC
# source: https://github.com/dasirrine/Syncro/blob/main/configDumpFiles.ps1
# latest revision 2024-04-22

# Sets Windows Crash (BSoD) settings to create full, kernel-only, small, or
# no memory dumps and keep all or only the last memory dump.
#
# Defaults to "Complete" for debugging purposes,
# and "overwrite" to save disk space.


#set default dump level to "Complete"
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
if (($lvl) -and ($lvl -in 0..3)) {
	"Crash dump files will be set to level: `"$dumpLevel`""
	Get-WmiObject -Class Win32_OSRecoveryConfiguration -EnableAllPrivileges | Set-WmiInstance -Arguments @{ DebugInfoType=$lvl }
} else {
    Write-Host "Invalid dump level specified."
}

#overwrite memory dump files unless specified
if ($overwriteDebugFiles="no") {
	Write-Host "ALL dump files will be retained."
	$CrashBehaviour | Set-WmiInstance -Arguments @{ OverwriteExistingDebugFile=$False }
	###warn if "Complete" dump files will be retained
	if ($lvl=1) { Write-Host "WARNING: Saving all `"Complete`" dump files can quickly fill the drive!" }
} else {
	Write-Host "only the last dump file will be saved."
	$CrashBehaviour | Set-WmiInstance -Arguments @{ OverwriteExistingDebugFile=$True }
}
