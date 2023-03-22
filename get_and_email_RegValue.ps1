#Import-Module $env:SyncroModule

$regPath = "HKLM:\Software\ThreatLocker"
$valueName = "ComputerId"

function Test-RegistryValue {
	param (
		[parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] $Path,
		[parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()] $Value
	)
	If ( Test-Path -Path $Path ) {
		try {
			Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
			return $true
		}
		catch {
			return $false
		}
	} Else { return $false }
}

If ( Test-RegistryValue -Path $regPath -Value $valueName ) {
    $value = Get-ItemPropertyValue -Path $regPath -Name $valueName
    Write-Output "The value of $regPath\$valueName is $value"
	Send-Email -To "you@yourmsp.com" -Subject "ThreatLocker info for $env:COMPUTERNAME" -Body "The value of $regPath\$valueName is $value"
} Else {
    Write-Output "The Registry key $regpath and/or the value $valueName was not found."
	Rmm-Alert -Category "Script Error" -Body "An error occurred while retrieving the registry value $regPath\$valueName"
}
