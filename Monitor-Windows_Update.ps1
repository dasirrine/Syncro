Import-Module $env:SyncroModule -DisableNameChecking
<#
## Windows Update Monitor ##
Author: David Sirrine
Source: 
Forked from: NullZilla's: script "Monitor - Windows Update"
https://pastebin.com/TWr17QBg
[ Nullzilla's Todo: properly handle less common Win10 editions LTSB/LTSC & ESU
	https://learn.microsoft.com/en-us/windows/release-health/release-information ]
#>

# Maximum age of builds to support in months
[int]$Win10_MaxAge = "18" # EOS for Home/Pro is 18 months except for 22H2 which we account for later
[int]$Win11_HomePro_MaxAge = "24" # EOS for Home/Pro is 24 months
[int]$Win11_EntEduIoT_MaxAge = "36" # EOS for Enterprise/Education/IoT Enterprise is 36 months

# Number of days to consider an update recent
$Recent = '50'

# set UTC Time Zone Offset to local time if not overriden at runtime
## reference to find your current UTC offset:
## https://www.timeanddate.com/time/difference/timezone/utc
if (-not $Offset) {
	$TimeZone = [System.TimeZoneInfo]::Local
	$Offset = -($TimeZone.BaseUtcOffset.TotalHours)
	if ($TimeZone.IsDaylightSavingTime([datetime]::UtcNow)) {
		$Offset -= 1
	}
}


# Function to Check Update Services
function Check-ServiceStatus {
    $DisabledServices = Get-Service WUAUServ, BITS, CryptSvc, RPCSS, EventLog | Where-Object StartType -eq 'Disabled'
    if ($DisabledServices) {
        #Write-Output "Disabled Services: $($DisabledServices | Select-Object -ExpandProperty Name)"
        #Rmm-Alert -Category "Monitor - Windows Update" -Body "Disabled Services: $($DisabledServices | Select-Object -ExpandProperty Name)"
		"Disabled Services:"
		$DisabledServices
		Rmm-Alert -Category "Monitor - Windows Update" -Body "Disabled Services: $DisabledService"
        exit 1
    } else {
        Close-Rmm-Alert -Category "Monitor - Windows Update"
    }
}

# Function to Ensure Microsoft Update Service is Enabled
function Ensure-MicrosoftUpdateService {
    $UpdateService = New-Object -ComObject Microsoft.Update.ServiceManager
    $MicrosoftUpdateService = $UpdateService.Services | Where-Object { $_.ServiceId -eq '7971f918-a847-4430-9279-4a52d1efe18d' }
    if (!$MicrosoftUpdateService) {
        $UpdateService.AddService2('7971f918-a847-4430-9279-4a52d1efe18d', 7, '')
    }
}


# get version number
$OSname = (Get-CimInstance Win32_OperatingSystem).Caption
if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion) {
	$Version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
} else { 
	$Version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId
}

# set max age based on OS version
if ($OSname -match 'Windows 10') {
	$MaxAge = if ($Version -eq '22H2') { 36 } else { $Win10_MaxAge }
} elseif ($OSname -match 'Windows 11 (Home|Pro)') {
    $MaxAge = $Win11_HomePro_MaxAge
} elseif ($OSname -match 'Windows 11 (Enterprise|Education|IoT Enterprise)') {
	$MaxAge = $Win11_EntEduIoT_MaxAge
} else {
	Write-Host "Unsupported OS: $OSname`nThis script only supports Windows 10 & 11, exiting..."
	exit 0
}

# Convert current date to comparable format and subtract $MaxAge
$CurrentDate = (Get-Date).AddMonths(-$MaxAge).ToString("yyMM")
$VersionNumerical = ($Version).replace('H1', '05').replace('H2', '10')
$Diff = $VersionNumerical - $CurrentDate

if ($Diff -lt 0) {
	Write-Output "$OSname $Version is over $MaxAge months old, needs upgrading"
	Rmm-Alert -Category "Monitor - Windows Update" -Body "$OSname $Version is over $MaxAge months old, needs upgrading"
	exit 1
} else {
	Write-Output "$OSname $Version is less than the max age of $MaxAge months"
	Close-Rmm-Alert -Category "Monitor - Windows Update"
}


# Check for Disabled Services
Check-ServiceStatus

# Ensure Microsoft Update Service is Enabled
Ensure-MicrosoftUpdateService

# Check for Recent Updates
$AutoUpdate = New-Object -ComObject Microsoft.Update.AutoUpdate
$UpdateResults = $AutoUpdate.Results
$SSDLastDate = ([datetime]$UpdateResults.LastSearchSuccessDate).AddHours($Offset)
$SSDDays = (New-TimeSpan -Start $SSDLastDate -End (Get-Date)).Days

$ISDLastDate = ([datetime]$UpdateResults.LastInstallationSuccessDate).AddHours($Offset)
$ISDDays = (New-TimeSpan -Start $ISDLastDate -End (Get-Date)).Days

Write-Output "Last Search Success: $SSDLastDate ($SSDDays days ago)"
Write-Output "Last Installation Success: $ISDLastDate ($ISDDays days ago)"

$Searcher = (New-Object -ComObject 'Microsoft.Update.Session').CreateUpdateSearcher()
$UpdateHistory = $Searcher.QueryHistory(0, $Searcher.GetTotalHistoryCount()) | 
    Where-Object { $_.Operation -eq 1 -and $_.ResultCode -match '[123]' } | 
    Select-Object -ExpandProperty Title

$LastMonth = (Get-Date).AddMonths(-1).ToString("yyyy-MM")
$ThisMonth = (Get-Date).ToString("yyyy-MM")

$RecentUpdates = $UpdateHistory | Where-Object { $_ -match "($LastMonth|$ThisMonth) (Security Monthly Quality Rollup for Windows|Cumulative Update for Windows)" -or $_ -match "Feature update" }

if (!$RecentUpdates -and $ISDDays -gt $Recent) {
    Write-Output "WARNING - No recent rollup/cumulative/feature update detected"
	Write-Output "Last updates:"
    $RecentUpdates | Select-Object -ExpandProperty Title -First 1
    Rmm-Alert -Category "Monitor - Windows Update" -Body "WARNING - No recent rollup/cumulative/feature update detected"
    exit 1
} else {
    Write-Output "Recent rollup or cumulative update detected"
    $RecentUpdates | Select-Object -ExpandProperty Title -First 1
    Close-Rmm-Alert -Category "Monitor - Windows Update"
}
