#run System File Checker
$SFCLogFile = "${env:windir}\Temp\SFC.log"
get-date >> $SFCLogFile
Start-Process -FilePath "${env:Windir}\System32\SFC.EXE" -ArgumentList '/scannow' -Wait -Verb RunAs >> $SFCLogFile

#check the Windows Component Store, Restore if needed
Repair-WindowsImage -Online -ScanHealth
if ($ComputerProperties = Repair-WindowsImage -Online -ScanHealth | Select-Object ImageHealthState | Where-Object {($_.ImageHealthState -notlike "Healthy")}) {
     Repair-WindowsImage -Online -RestoreHealth}





#make a note of SFC repairs from CBS.LOG
findstr /c:"[SR]" "${env:windir}\logs\cbs\cbs.log" >> $SFCLogFile

#attach repair logs to Syncro Asset
Import-Module $env:SyncroModule -WarningAction SilentlyContinue
Upload-File -FilePath $SFCLogFile
Upload-File -FilePath "${env:windir}\logs\dism\DISM.log"
