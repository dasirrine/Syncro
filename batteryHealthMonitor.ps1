
# batteryHealthMonitor.ps1 by David Sirrine @ CNS Computer Services, LLC
# source: https://github.com/dasirrine/Syncro/blob/main/batteryHealthMonitor.ps1
# latest revision 2024-04-18
#
# Monitors battery health by comparing the full charge capacity to the design
# capacity. Specified percentage is the threshold below which a battery is
# considered "unhealthy," e.g. inputting 40% would trigger the alert if the
# battery's full charge capacity is less than 40% of its design capacity.
# optionally updates a custom Asset field "Battery Health" with the current
# capacity (hat tip to Jackson Wyatt).
#
# optional:
# 1) Target capacity defaults to 50% if no value is provided at runtime.
# 2) script writes to an Asset custom field "Battery Health" if present.
#
# NOTE: batteries with an ID containing "UPS" are excluded, as the status of
#       UPSs does not appear to be reported correctly by POWERCFG.
# 
# [based on the script "CyberDrain.com - Battery Health Monitor"
# from the Community Library.]


Import-Module $env:SyncroModule -WarningAction SilentlyContinue

 ### set default health threshold if not provided at runtime ###
if (-not $alertPercent) { $alertPercent = 50 }

### location and name of XML results file ###
$tmpFolder = "C:\temp"
$reportFile = "battery_report.xml"


# Exit if the script is not running on a mobile device
$isMobile = (Get-WmiObject -Class Win32_ComputerSystem).PCSystemType -eq 2
if (-not $isMobile) {
    Write-Host "This is not a mobile device. Aborting."
    exit
}

#create temp folder if it doesn't exist
If (!(Test-Path $tmpFolder)) { New-Item -ItemType Directory -Force -Path $tmpFolder | Out-Null }

#generate XML report of battery status
Start-Process "powercfg.exe" -ArgumentList "/batteryreport /XML /OUTPUT `"$tmpfolder\$reportFile`"" -Wait -WindowStyle Hidden

try {
    #gather battery data from POWERCFG output
    [xml]$Report = Get-Content "$tmpfolder\$reportFile" -ErrorAction Stop
} catch {
	#exit if the XML file didn't load (e.g. if it failed to generate)
	Write-Host "This device does not have a battery to monitor, or the status of the batteries could not be found.`n"
	$_
    exit
}

foreach ($batt in $Report.BatteryReport.Batteries.Battery) {

	#exclude UPS batteries
	if ($batt.ID -notlike "*UPS*") {

		#calculate the capacity percentage and round for readability
		$designCap = [Math]::Round($batt.DesignCapacity / 1000, 2)
		$fullCap = [Math]::Round($batt.FullChargeCapacity / 1000, 2)
		$fullPercent = [int64]$batt.FullChargeCapacity * 100 / [int64]$batt.DesignCapacity
		$fullPercent = [Math]::Round($fullPercent, 2)
		
		if ($fullPercent -lt $alertPercent) {
			#battery capacity is below the specified threshold
			$aboveORbelow = "below"
			$andORbut = "but"
			$script:uploadXML = $true
			#raise an alert per failed battery
			Rmm-Alert -Category "Battery_Health" -Body "Battery `"$($batt.id)`" charge capacity is $fullPercent% of rated capacity. Charge capacity/Rated capacity is $fullCap/$designCap wH." -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
		} else {
			#battery has sufficient capacity
			$aboveORbelow = "above"
			$andORbut = "and"
		}

		#collect battery status for Asset custom field
		if ($battStatus) { $battStatus = " / " }
		$battStatus = $battStatus + "Battery `"$($batt.id)`": $fullPercent% ($fullCap/$designCap wH)"

		#display results *per battery*
		Write-Host "Charge capacity of battery `"$($batt.id)`" is $aboveORbelow the specified threshold of $alertPercent%."
		Write-Host "Rated capacity of the battery is $designCap wH $andORbut the current maximum charge is $fullCap wH ($fullPercent%)."
		
	}
}

#timestamp battery status and write to Asset custom field
$battStatus = (Get-Date -Format "M/d/yyyy h:mmtt") + " - " + $battStatus
Set-Asset-Field -Name "Battery Health" -Value $battStatus -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

#upload the results XML file if *any* battery failed
if ( $script:uploadXML ) { Upload-File -FilePath "$tmpfolder\$reportFile" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }

