 
$sourceItemNameList = @()
$sourceItemDirNameList = @()
$deltaItemNameList = @()
$deltaItemDirNameList =@()
$changedFiles = @()
$changedFileNames = @()
 
 <#
            ForEach ($sItem in $sourceFileList) {
                $sourceItemName = $sItem.PSChildName
                $sourceItemNameList += $sourceItemName
                $sourceItemFullName = $sItem.FullName
                $sourceItemDirName = $sourceItemFullName.Replace($sourcepath,"")
                $sourceItemDirNameList += $sourceItemDirName
            
            }
            #>
            #>

           
            #$propertyList = ($sourceFileList | Get-Member -MemberType Property | Select-Object -ExpandProperty Name)

#>
            #$differences = Compare-Object -ReferenceObject $sourceFileList -DifferenceObject $deltaFileList -Property LastWriteTime | Format-Table
            
            <#
            ForEach ($dItem in $differences) {
                $deltaItemName = $dItem.PSChildName
                $deltaItemNameList += $deltaItemName
                $deltaItemFullName = $dItem.FullName
                $deltaItemDirName = $deltaItemFullName.Replace($deltaPath,"")
                $deltaItemDirNameList += $deltaItemDirName
                $dirNameMatch = $sourceFileList.Where{($_.FullName -match $deltaItemDirName)}
                Write-Host $dItem.LastWriteTime " - " $dirNameMatch.LastWriteTime
                
                if ($dItem.LastWriteTime -gt $dirNameMatch.LastWriteTime) {
                    Write-Host "The delta version of" $deltaItemName "is newer!"
                    $changedFiles += $dItem
                }
                Write-Host "Here's one:" $deltaItemFullName
            }
            #>

 
            <#
            $dirDifferences = Compare-Object -ReferenceObject $sourceItemDirNameList -DifferenceObject $deltaItemDirNameList
            
            foreach ($item in $differences) {
               
                if ($item.Directory.StartsWith($deltaPath,1)) {
                    $deltaItems += $item
                    $deltaItemNames += $item.PSChildName
                } else {
                    $sourceItems += $item
                    $sourceItemNames += $item.PSChildName
                }
                #$itemName = if ($itemFullName.StartsWith($sourcePath,1)) {$itemFullName.Replace($sourcePath,"")} else {$itemFullName.Replace($diffPath,"")}
                #Write-Host $itemFullName
                #Write-Host $itemName
            }
            
            
            $differences | Sort-Object {$_.LastWriteTime} -Descending

            
            
            ForEach ($dItem in $deltaFileList) {
                if ($dItem.PSChildName -in $sourceItemNames) {
                    #$dItemName = $dItem.PSChildName
                    $dItemFullName = $dItem.FullName
                    $dItemDirName = $dItemFullName.Replace($deltaPath,"")
                    $match = $sourceFileList.Where{($_.FullName -match $dItemDirName)}
                    Write-Host "There is a match!" $match
                    if ($dItem.LastWriteTime -gt $match.LastWriteTime) {
                        Write-Host "The delta version is newer!"
                        $changedFiles += $dItem
                    }
                }
            }#>
            <#
            ForEach ($cItem in $changedFiles) {
                $changedFileNames += $cItem.FullName
            }

            $changedFileNames
            

            $fullDeltaList#>

# Get the list of mounted RPs, which should only contain a single volume (the one specified earlier) #
        #$PMountedRPs = Get-Mounts -ProtectedServer $PSLPHostName
        
        # Try the alternate name if the first doesn't work #
        <#If ($PMountedRPs.Exception -ne $null) {
            $PMountedRPs = Get-Mounts -ProtectedServer $PSDisplayName
        }#>

        # Get the display name of the volume and store it #
        #$PMRPoint = $PMountedRPs[0]
        #$PMRPPath = $PMRPoint.MountedVolumeDisplayName
        #$PTimeCheckResult = [string]
#$RTimeCheckResult = [string]
#$PMountCheckResult = [string]
#$RMountCheckResult = [string]
#$PSDisplayName = [string]
#$PSMasterCoreInfo = $null
#$PSCoreName = [string]
#$PSLatestPoint = $null
#$PSLPRPN = [int]
#$PSLPDateTime = [datetime]
#$PSLPHostName = [string]
#$PSLPVIS = $null
#$PSLPVDN = [string]
#$PMountRP = $null
#$PMountedRPs = @()
#$PMRPoint = $null
#$PMRPPath = [string]
#$RSDisplayName = [string]
#$RSMasterCoreInfo = $null
#$RSRemoteCoreName = [string]
#$RSLatestPoint = $null
#$RSLPRPN = [int]
#$RSLPDateTime = [datetime]
#$RSLPHostName = [string]
#$RSLPVIS = $null
#$RSLPVDN = [string]
#$RMountRP = $null
#$RMountedRPs = @()
#$RMRPoint = $null
#$RMRPPath = [string]

# Testing to see if the variables are empty or not #
#Read-Host "There should be some blanks here: $PTimeCheckResult $PSCoreName $PSLPHostName $PMRPPath. Better quit if not!"
        #$PSMasterCoreInfo = $PServer.MasterCoreInfo
        #$PSCoreName = $CoreName
        #Read-Host "The $PSLPVDN volume of $PSDisplayName (aka $PSLPHostName) from $PSLPDateTime will be mounted from recovery point number $PSLPRPN. Here we go"
        #Read-Host "Feel free to bang your head against the desk at this point"
        Write-Host "The $RSLPVDN volume of $RSLPHostName from $RSLPDateTime will be mounted from recovery point number $RSLPRPN."

        
        $RMountedRPs = Get-Mounts -ProtectedServer $RSLPHostName
        $RMRPoint = $RMountedRPs[0]
        $RMRPPath = $RMRPoint.MountedVolumeDisplayName
         
        If ($RMRPPath -eq $null) {
            Remove-Mounts
            Write-Host "Retrying mount..."
            $RMountRP = New-Mount -ProtectedServer $RSLPHostName -Rpn $RSLPRPN -Path "C:\Mounts" -Volumes $RSLPVDN
            $RMountedRPs = Get-Mounts -ProtectedServer $RSLPHostName
            $RMRPoint = $RMountedRPs[0]
            $RMRPPath = $RMRPoint.MountedVolumeDisplayName
        }
#$EmailFrom = "$PSCoreName@sewell.com"
#$EmailTo = "shamilton@sewell.com"