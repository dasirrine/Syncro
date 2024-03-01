Import-Module $env:SyncroModule -WarningAction SilentlyContinue

#make a zip of the minidump folder
$dumpZip = "$env:TEMP\minidump.zip"
$minidumpFolder = "$env:windir\Minidump"
$error.clear()
Write-Host "Compressing all files in the MiniDump folder"
Compress-Archive -Path $minidumpFolder -DestinationPath $dumpZip -Update -CompressionLevel Optimal -ErrorAction SilentlyContinue
If ( -Not $error ) {

    #attach minidumps to Syncro Asset
    Write-Host "Uploading MiniDump archive to Asset record"
    $error.clear()
	Upload-File -FilePath $dumpZip -ErrorAction SilentlyContinue -ErrorVariable UploadError

		#delete the zip [-and empty the minidump folder]
		If ( -Not $error ) {
			Write-Host "Deleting MiniDump archive"
            Remove-Item -path $dumpZip
			#Remove-Item -path $minidumpFolder\*.* -recurse
		}
} Else {
    Write-Host "Error while creating MiniDump archive"
}
