# Monthly Settings Backup script for Rapid Recovery #

# Import RR module and declare variables #
Import-Module RapidRecoveryPowerShellModule
$DateStamp = Get-Date -format "MMddyy-HHmm"
$LocalCopyNewName = "AppRecoveryCoreConfigurationBackup$DateStamp.xml"
$LPath = "D:\AppAssure\SettingsBackup"
$RPath = "\\srvrdc42\filestorage\documentation\projects\servers\rapidrecovery\settingsexport\"
$RRegPath = "\\srvrdc42\filestorage\documentation\projects\servers\rapidrecovery\registryexport\"
$LStopWatch = New-Object -TypeName System.Diagnostics.Stopwatch

# Archive all previous configuration backups #
Move-Item -Path "$LPath\*.xml" -Destination "$LPath\Archive"

# Export backup of configuration #
Start-BackupSettings -LocalPath $LPath

# Start the timer to ensure the previous command has enough time to complete #
$LStopWatch.Start()
$LSeconds = $LStopWatch.Elapsed.Seconds

# Check the directory to make sure the XML file is present, waiting up to 30 seconds for the process to complete #
Do {
    $LDirCheck = Get-ChildItem -Path "$LPath" -Name -File
} Until (($LDirCheck -like "*.xml") -or ($LSeconds -gt 30))
$LStopWatch.Stop()
$LStopWatch.Reset()

# Rename the backup to reflect the current date, then archive the XML files on the remote storage #
Rename-Item -Path "$LPath\*.xml" -NewName $LocalCopyNewName
Move-Item -Path "$RPath\*.xml" -Destination "$RPath\Archive"

# Copy the backup to the remote location, where it can be offloaded to tape as part of the monthly archive #
Copy-Item -Path "$LPath\$LocalCopyNewName" -Destination "$RPath"

# Archive any old registry backups and export the current file #
Move-Item -Path "$RRegPath\*.reg" -Destination "$RRegPath\Archive"
reg export "HKLM\Software\AppRecovery" "$RRegPath\apprecovery$DateStamp.reg"