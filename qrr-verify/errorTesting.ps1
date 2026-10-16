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
            } catch {
                $sourceBESFileList = Get-ChildItem -path $sourceBESPathx64 -file -ea Stop
            } finally {
                Write-Host "No log files detected; proceeding with full drive scan."
            }
            try {
                $deltaBESFileList = Get-ChildItem -path $deltaBESPath -file -ea Stop
            } catch {
                $deltaBESFileList = Get-ChildItem -path $deltaBESPathx64 -file -ea Stop
            } finally {
                Write-Host "No log files detected; proceeding with full drive scan."
            }