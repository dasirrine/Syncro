Import-Module $env:SyncroModule -WarningAction SilentlyContinue

<#
.VARIABLES
$channel {"values"??"current-channel","monthly-enterprise-channel","semi-annual-enterprise-channel"],"default":"current-channel","index":0}
$forceUpdate {"values"??"No","Yes"],"default":"No","index":0}
.SCRIPT
Language : PowerShell
Run As: Logged in User
.SYNOPSIS
Office365 Updater for Syncro
.DESCRIPTION
Checks  Office365 build and updates to latest version based on selected channel
.NOTES
Version : 1.2.2
Author: Mar Szt
(expanded upon by David Sirrine @ CNS 2023-03-15)
.EXAMPLE
Update-Office365 -channel "monthly-enterprise-channel"
Checks for latest version in monthly enterprise channel and updates if needed
Update-Office365 -forceUpdate $true
Checks for latest version in current channel (default) and force close all office apps if needed
#>

#set default O365 updater parameters
if ($null -eq $channel) {$channel = "current-channel"}
if ($null -eq $forceUpdate) { $forceUpdate = "no" }
$updater = "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe"
$updater_args = "/update user"
if ($forceUpdate -eq 'yes') { $updater_args += " forceappshutdown=true" }

#get O365 build currently installed
$build_installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').VersionToReport

#check if O365 is installed; log O365 version in Syncro
if ( $null -eq $build_installed ) {
    Write-Host "update not required; MS Office 365 not installed"
	Set-Asset-Field -Name "O365 version" -Value "n/a"
} else {
	#log O365 version installed before updates
	Set-Asset-Field -Name "O365 version" -Value $build_installed

	#get latest MS Office 365 build number
	$page_html = Invoke-RestMethod "https://learn.microsoft.com/en-us/officeupdates/$channel"
	$reg_exp = '<em>Version (?<version>.*) \(Build (?<build>.*)\)<\/em>'
	$versions = ($page_html | Select-String $reg_exp -AllMatches).Matches
	$build_latest = "16.0."+($versions[0].Groups.Where{$_.Name -like 'build'}).Value

	if ( $null -eq $build_latest ) { $build_latest = "[build version not found]" }

	#check if O365 update is needed (or latest build couldn't be retrieved)
	if ( $build_installed -ne $build_latest ) {
		$updateComplete = $false

		#check for update if O365 updater exists
		if (Test-Path $updater) {

			#user broadcast message
			$message = "Your Microsoft Office needs an urgent update to improve its performance and security. After the update finishes, please restart your system at your earliest convenience."
			Broadcast-Message -Title "Updates Notification" -Message $message -LogActivity "true"

			#start update
			Write-Host "Updating Office from build $build_installed to $build_latest"
			Start-Process -FilePath $updater -ArgumentList $updater_args -Wait
			if ( -not $error ) {
				$updateComplete = $true
				Start-Sleep -Seconds 90
				$build_installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').VersionToReport
				#log update in Syncro
				Log-Activity -Message "critical update" -EventName "Office 365 updated to $build_installed"
				Set-Asset-Field -Name "O365 version" -Value $build_installed
			}
		}

		#log an error and create a ticket if the update didn't work
		if ( $updateComplete = $false ) {
			Log-Activity -Message "critical update" -EventName "error updating Office 365 $build_installed to $build_latest"
			Create-Syncro-Ticket -Subject "error updating Office 365 on $env:computername" -IssueType "Regular Maintenance" -Status "New"
		} 
		
	} else {
		Write-Host "Office update not required - build installed is $build_installed"
	}
}
