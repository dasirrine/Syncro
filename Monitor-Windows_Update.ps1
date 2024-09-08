Import-Module $env:SyncroModule -DisableNameChecking
$ErrorActionPreference = 'Stop'

<#
## Windows Update/Version Monitor ##
	Author: David Sirrine
	Source: https://github.com/dasirrine/Syncro/blob/main/Monitor-Windows_Update.ps1
	Forked from: NullZilla's: script "Monitor - Windows Update"
https://pastebin.com/TWr17QBg
	[ Nullzilla's Todo: properly handle less common Win10 editions LTSB/LTSC & ESU
	https://learn.microsoft.com/en-us/windows/release-health/release-information ]
	Potential addition - create ticket to log time?
#>

# Maximum age of builds to support in months
[int]$Win10_MaxAge = "18" # EOS for Home/Pro is 18 months except for 22H2 which we account for later
[int]$Win11_HomePro_MaxAge = "24" # EOS for Home/Pro is 24 months
[int]$Win11_EntEduIoT_MaxAge = "36" # EOS for Enterprise/Education/IoT Enterprise is 36 months

## set this variable if you don't want this script to enable Windows updates
## (e.g. if you're controlling Windows updates via Syncro policy
##  or a third-party patch manager)
#$enableWindowsUpdates = $false

# Set number of days to consider an update recent (if not provided at runtime)
if (!$Recent) { $Recent = '50' }

# set UTC Time Zone Offset to local time (if not provided at runtime)
# # reference to find your current UTC offset:
# # https://www.timeanddate.com/time/difference/timezone/utc
if (!$Offset) {
	$TimeZone = [System.TimeZoneInfo]::Local
	$Offset = -($TimeZone.BaseUtcOffset.TotalHours)
	if ($TimeZone.IsDaylightSavingTime([datetime]::UtcNow)) {
		$Offset -= 1
	}
}

# get Windows name and version number
$OSname = (Get-CimInstance Win32_OperatingSystem).Caption
if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion) {
	$Version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
} else { 
	$Version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId
}
if (!$OSname -or !$Version) {
	"Error retrieving OS name and/or version."
	exit 0
}
# set the max age for version based on OS
if ($OSname -match 'Windows 10') {
	$MaxAge = if ($Version -eq '22H2') { 36 } else { $Win10_MaxAge }
} elseif ($OSname -match 'Windows 11 (Home|Pro)') {
    $MaxAge = $Win11_HomePro_MaxAge
} elseif ($OSname -match 'Windows 11 (Enterprise|Education|IoT Enterprise)') {
	$MaxAge = $Win11_EntEduIoT_MaxAge
} else {
	"Unsupported OS: $OSname`nThis script only supports Windows 10 & 11, exiting..."
	exit 0
}
# Convert current date to comparable format and subtract $MaxAge
$CurrentDate = (Get-Date).AddMonths(-$MaxAge).ToString("yyMM")
$VersionNumerical = ($Version).replace('H1', '05').replace('H2', '10')
$Diff = $VersionNumerical - $CurrentDate


# Check 1: if services needed for Windows update are disabled; raise Alert if needed
$DisabledServices = Get-Service WUAUServ, BITS, CryptSvc, RPCSS, EventLog | Where-Object StartType -eq 'Disabled'
if ($DisabledServices) {
	$Failure = "Disabled Services: $($DisabledServices | Select-Object -ExpandProperty Name)"
	##$Failure
	Rmm-Alert -Category "Monitor - Windows Update" -Body $Failure -ErrorAction 'SilentlyContinue'
	##exit 1
	$script:Results = $Failure ##
} else {
	$script:Results = "Services needed for Windows Update are not disabled."
	##Close-Rmm-Alert -Category "Monitor - Windows Update" -ErrorAction 'SilentlyContinue'
}


# Check 2: if Microsoft Update Service is disabled; re-create if needed
$UpdateService = New-Object -ComObject Microsoft.Update.ServiceManager
$MicrosoftUpdateService = $UpdateService.Services | Where-Object { $_.ServiceId -eq '7971f918-a847-4430-9279-4a52d1efe18d' }
if (!$MicrosoftUpdateService) {
	$Failure = "Windows Update Service is disabled or not present"
	##$Failure
	$script:Results += "`n$Failure" ##

	# enable Windows updates unless explicitly directed otherwise
	if ($enableWindowsUpdates -ne $false) {
		$UpdateService.AddService2('7971f918-a847-4430-9279-4a52d1efe18d', 7, '')
		$script:Results += "`nEnabled Windows Update service"
	}
}


# Check 3: if OS version is current; raise Alert if needed
if ($Diff -lt 0) {
	$Failure = "$OSname $Version is over $MaxAge months old, needs upgrading"
	#$Failure
	Rmm-Alert -Category "Monitor - Windows Update" $Failure -ErrorAction 
	'SilentlyContinue'
	##exit 1
	$script:Results += "`n$Failure" ##
} else {
	$script:Results += "`n$OSname $Version is less than the max age of $MaxAge months"
	##Close-Rmm-Alert -Category "Monitor - Windows Update" -ErrorAction 'SilentlyContinue'
}


# Check 4: for Recent Updates; raise Alert if needed
try {
	$AutoUpdate = New-Object -ComObject Microsoft.Update.AutoUpdate
	$UpdateResults = $AutoUpdate.Results
	$SSDLastDate = ([datetime]$UpdateResults.LastSearchSuccessDate).AddHours($Offset)
	$SSDDays = (New-TimeSpan -Start $SSDLastDate -End (Get-Date)).Days

	$ISDLastDate = ([datetime]$UpdateResults.LastInstallationSuccessDate).AddHours($Offset)
	$ISDDays = (New-TimeSpan -Start $ISDLastDate -End (Get-Date)).Days
	$script:Results += "`nLast Search Success: $SSDLastDate ($SSDDays days ago)"
	$script:Results += "`nLast Installation Success: $ISDLastDate ($ISDDays days ago)"

	$Searcher = (New-Object -ComObject 'Microsoft.Update.Session').CreateUpdateSearcher()
	$UpdateHistory = $Searcher.QueryHistory(0, $Searcher.GetTotalHistoryCount()) | 
		Where-Object { $_.Operation -eq 1 -and $_.ResultCode -match '[123]' } | 
		Select-Object -ExpandProperty Title

	$LastMonth = (Get-Date).AddMonths(-1).ToString("yyyy-MM")
	$ThisMonth = (Get-Date).ToString("yyyy-MM")

	$RecentUpdates = $UpdateHistory | Where-Object { $_ -match "($LastMonth|$ThisMonth) (Security Monthly Quality Rollup for Windows|Cumulative Update for Windows)" -or $_ -match "Feature update" }
 
	if (!$RecentUpdates -and $ISDDays -gt $Recent) {
		$Failure = "WARNING - No recent rollup / cumulative update / feature update detected"
		##$Failure
		Rmm-Alert -Category "Monitor - Windows Update" $Failure -ErrorAction 
		'SilentlyContinue'
		##exit 1
		$script:Results += "`n$Failure" ##
	} else {
		$script:Results += "`nRecent rollup or cumulative update detected"
		#Close-Rmm-Alert -Category "Monitor - Windows Update" -ErrorAction 'SilentlyContinue'
	}

	if ($RecentUpdates) {
		$script:Results += "`nLast updates:`n"
		$script:Results += $RecentUpdates | Select-Object -ExpandProperty Title -First 1
	}
}
catch {
	$Failure = "Error checking for recent updates."
	##$Failure	
	Rmm-Alert -Category "Monitor - Windows Update" $Failure -ErrorAction 'SilentlyContinue'
	$script:Results += "`n$Failure" ##
}

#Display all results
$script:Results
