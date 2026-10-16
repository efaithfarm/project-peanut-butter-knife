# Instantiate Arrays #
$mountedRPs = @()
#$rPoints = @()
$deltaPtSizeMatches = @()
$sourcePtSizeMatches = @()
$tracker = 0
$sessionTracker = 0
$serversVerified = @()
$coreServersVerified = @()
$sourceBESFileList = @()
$deltaBESFileList = @()
$deltaItems = @()


#function Process-Verifications {
# Check status of script #
while ($tracker -lt 1) {

    # Connect to RR Core Server #
    $coreServer = Read-Host "Enter Core server name"
    $coreServersVerified += $coreServer
    Invoke-Command -ComputerName $coreServer -ScriptBlock {
        
        # Check status of session #
        while ($sessionTracker -lt 1) {
            
            # Load RR Module #
            try {
                Import-Module RapidRecoveryPowerShellModule -ea Stop
            } 
            catch {
                Write-Host "Sorry, loading the module failed. That server might not be a Rapid Recovery core."
                break
            }
            
            # Prompt for server, clear any previous mounts #
            $server = Read-Host "Enter server name"
            $serversVerified += $server
            Remove-Mounts

            # Get the list of recovery points for the server and collect data for the volume being verified #  
            $rPoints = Get-RecoveryPoints -protectedServer $server -number all
            $deltaPt = $rPoints[0]
            $deltaRPN = $deltaPt.number
            $deltaRPTimeStamp = $deltaPt.dateTimeStamp
            $deltaPtVIS = $deltaPt.volumeImageSummaries
            $deltaPtSizeMatches = $deltaPtVIS.Where{(($_.rawCapacity -gt 1000000000) -and ($_.volumeDisplayName -ne "(Volume Labeled 'RECOVERY')"))}
            $deltaVDN = $deltaPtSizeMatches[0].volumeDisplayName
            
            # Look through the list to find the next-latest recovery point for the volume #
            forEach ($sourcePt in $rPoints) {
                $sourceRPN = $sourcePt.number
                $sourceRPTimeStamp = $sourcePt.dateTimeStamp
                $sourceVIS = $sourcePt.volumeImageSummaries
                $sourcePtSizeMatches = $sourceVIS.Where{(($_.rawCapacity -gt 1000000000) -and ($_.volumeDisplayName -ne "(Volume Labeled 'RECOVERY')"))}
                $sourceVDN = $sourcePtSizeMatches[0].volumeDisplayName
                if (($sourceVDN -eq $deltaVDN) -and ($sourceRPTimeStamp -lt $deltaRPTimeStamp)) {
                    break
                }
            }

            # Display info on recovery points to be mounted #
            $displayDateTimeStamp = Get-Date -format "HH:mm"
            Write-Host "The $sourceVDN volume from $sourceRPTimeStamp will be mounted from recovery point number $sourceRPN as the source at $displayDateTimeStamp."
            Write-Host "The $deltaVDN volume from $deltaRPTimeStamp will be mounted from recovery point number $deltaRPN as the delta at $displayDateTimeStamp."

            # Mount the recovery points #
            New-Mount -protectedServer $server -rpn $deltaRPN -path "C:\mounts\delta" -volumes $deltaVDN
            New-Mount -protectedServer $server -rpn $sourceRPN -path "C:\mounts" -volumes $sourceVDN
            
            # Generate list of files on each recovery point #
            $mountedRPs = get-mounts -protectedServer $server
            $deltaRP = $mountedRPs[$mountedRPs.count -2]
            $sourceRP = $mountedRPs[$mountedRPs.count -1]
            $sourcePath = $sourceRP.mountedVolumeDisplayName
            $deltaPath = $deltaRP.mountedVolumeDisplayName

            # Check for presence of BigFix log files to use as directed comparison files #
            $besLogPath = "Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
            $besLogPathx64 = "Program Files\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
            $sourceBESPath = $sourcePath + $besLogPath
            $deltaBESPath = $deltaPath + $besLogPath
            $sourceBESPathx64 = $sourcePath + $besLogPathx64
            $deltaBESPathx64 = $deltaPath + $besLogPathx64
            try {
                $sourceBESFileList = Get-ChildItem -path $sourceBESPath -file -ea Stop
                Write-Host "BES log detected, parsing files."
            } catch {
                try {
                    $sourceBESFileList = Get-ChildItem -path $sourceBESPathx64 -file -ea Stop
                    Write-Host "BES log detected, parsing files."
                } catch {
                    Write-Host "No log files detected; proceeding with full drive scan."
                }
            }
            try {
                $deltaBESFileList = Get-ChildItem -path $deltaBESPath -file -ea Stop
                Write-Host "BES log detected, parsing files."
            } catch {
                try {
                    $deltaBESFileList = Get-ChildItem -path $deltaBESPathx64 -file -ea Stop
                    Write-Host "BES log detected, parsing files."
                } catch {
                    Write-Host "No log files detected; proceeding with full drive scan."
                }
            }
            $sourceBESFileListNo = $sourceBESFileList.count
            $deltaBESFileListNo = $deltaBESFileList.count
            Write-Host "$sourceBESFileListNo BES items for source"
            Write-Host "$deltaBESFileListNo BES items for delta"
            
            # Compare lists and collect differencing files #
            $propertyList = @("LastWriteTime","CreationTime")
            $deltaItems = ForEach ($property in $propertyList) {
                if (($sourceBESFileList.count -gt 0) -or ($deltaBESFileList.count -gt 0)) {
                    Compare-Object $sourceBESFileList $deltaBESFileList -Property "$property" -PassThru | Where-Object {$_.FullName.StartsWith("$deltaPath")} | Select-Object Name,Length,Directory,FullName,CreationTime,LastWriteTime | Sort-Object -Property FullName
                } else {
                    $sourceFileList = Get-ChildItem -path $sourcePath -recurse -file
                    $deltaFileList = Get-ChildItem -path $deltaPath -recurse -file
                    Compare-Object $sourceFileList $deltaFileList -Property "$property" -PassThru | Where-Object {$_.FullName.StartsWith("$deltaPath")} | Select-Object Name,Length,Directory,FullName,CreationTime,LastWriteTime | Sort-Object -Property FullName
                }
            }
            
            # Generate CSV file from comparison output #
            $csvDateStamp = Get-Date -format "MMddyy-HHmm"
            $csvFileName = "$server--$csvDateStamp.csv"
            $csvPath = "c:\mounts\$csvFileName"
            $deltaItems | Export-Csv -Path $csvPath
            
            # Remove any mounted recovery points #
            Remove-Mount -ProtectedServer $server

            # Prompt user to close session with Core server or pick another protected machine on that Core #
            $coreSessionClose = Read-Host "Ready to close Core session? Y or N" 
            if (($coreSessionClose -eq "y") -or ($coreSessionClose -eq "yes")) {
                $sessionTracker ++
            }
        }
    } 
    
    # Move output CSV to centralized location #
    try {
        Move-Item "\\$coreServer\C$\mounts\*.csv" "\\srvrdc42\filestorage\documentation\projects\servers\appassure verification\export" -ea Stop
        Write-Host "Output moved to central repository. Well done you."
    } catch {
        Write-Host "No output found. Or maybe there's a credential problem."
    }
       
    # Prompt user to exit script or choose another Core #
    $verifySessionClose = Read-Host "Ready to quit verification? Y or N"
    if (($verifySessionClose -eq "y") -or ($verifySessionClose -eq "yes")){
        $tracker ++
    }
} 

#}