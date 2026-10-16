# Rapid Recovery Verification Script #
# Uses RR PS module to verify currency and mountability of RR recovery points #

# Create Variables and Arrays #
$TableEntry = @()
$ProtectedServers = @()
$ReplicatedServers = @()
$PSRPoints = @()
$PSLPSizeMatches = @()
$RSRPoints = @()
$RSLPSizeMatches = @()
$PSCoreName = "RR_Server_1"
$RCounter = 0

# Create tables that will be used in the wrap-up email #
# First create table for servers that pass the checks #
$DataTablePass = New-Object System.Data.DataTable "ResultsTable"
$PCol1 = New-Object System.Data.DataColumn MachineName,([string])
$PCol2 = New-Object System.Data.DataColumn LastSnapTime,([string])
$PCol3 = New-Object System.Data.DataColumn Current,([string])
$PCol4 = New-Object System.Data.DataColumn MountCheck,([string])
$PCol5 = New-Object System.Data.DataColumn Core,([string])
$DataTablePass.Columns.Add($PCol1)
$DataTablePass.Columns.Add($PCol2)
$DataTablePass.Columns.Add($PCol3)
$DataTablePass.Columns.Add($PCol4)
$DataTablePass.Columns.Add($PCol5)

# Next create table for objects that fail the checks #
$DataTableFail = New-Object System.Data.DataTable "ResultsTable"
$FCol1 = New-Object System.Data.DataColumn MachineName,([string])
$FCol2 = New-Object System.Data.DataColumn LastSnapTime,([string])
$FCol3 = New-Object System.Data.DataColumn Current,([string])
$FCol4 = New-Object System.Data.DataColumn MountCheck,([string])
$FCol5 = New-Object System.Data.DataColumn Core,([string])
$DataTableFail.Columns.Add($FCol1)
$DataTableFail.Columns.Add($FCol2)
$DataTableFail.Columns.Add($FCol3)
$DataTableFail.Columns.Add($FCol4)
$DataTableFail.Columns.Add($FCol5)
$ScriptStartTime = Get-Date

# Import the RR PS module #
Import-Module RapidRecoveryPowerShellModule

# Get list of protected servers and store it in an array, then ensure there are no mounted RPs #
$ProtectedServers = Get-ProtectedServers | Sort-Object -Property DisplayName
Remove-Mount -All

# Make sure that there are servers being locally protected; otherwise skip this step #
# If the Get-ProtectedServers doesn't contain anything it will return as a string #
If ($ProtectedServers -isnot [String]) {

    # Iterate through the list of servers #
    ForEach ($PServer in $ProtectedServers) {

        # Pull information about each server and store it in various variables #
        $PSDisplayName = $PServer.DisplayName
        Write-Host "Checking $PSDisplayName"
        $PSRPoints = Get-RecoveryPoints -ProtectedServer $PSDisplayName -Number All
        $PSLatestPoint = $PSRPoints[0]
        $PSLPRPN = $PSLatestPoint.Number
        Write-Host "Checking Recovery Point Number $PSLPRPN"
        $PSLPDateTime = $PSLatestPoint.DateTimestamp
        $PSLPHostName = $PSLatestPoint.AgentHostName
        $PStopWatch = New-Object -TypeName System.Diagnostics.Stopwatch

        # Check to make sure the DateTimestamp of the last RP is less than 24 hours old
        If ($PSLPDateTime -gt (Get-Date).AddHours(-24)) {
            $PTimeCheckResult = "Yes"
        } Else {
            $PTimeCheckResult = "No"
        }
        Write-Host "Is it current? $PTimeCheckResult"

        # Look at the volumes being protected and select only volumes larger than 1 GB that aren't recovery partitions #
        $PSLPVIS = $PSLatestPoint.VolumeImageSummaries
        $PSLPSizeMatches = $PSLPVIS.Where{(($_.rawCapacity -gt 1000000000) -and ($_.volumeDisplayName -ne "(Volume Labeled 'RECOVERY')"))}
        $PSLPVDN = $PSLPSizeMatches[0].VolumeDisplayName

        # Mount the latest RP according to the data pulled by the script so far #
        $PMountRP = New-Mount -ProtectedServer $PSDisplayName -Rpn $PSLPRPN -Path "C:\Mounts" -Volumes $PSLPVDN

        # If $PMountRP has a value for its Exception property, mark it as having failed the mount check, add its data to the Fail table, and exit the loop #
        # If its Exception is null, mark it as having passed and get the name of the mounted volume #
        If ($PMountRP.Exception -ne $null) {
            $PMountCheckResult = "Fail"
            Write-Host "Oops! Exception detected!"
            Write-Host $PMountRP.Exception
            
            $PSTableEntry = [System.Collections.ArrayList]@()
            $PSTableEntry.Insert(0,$PSDisplayName)
            $PSTableEntry.Insert(1,$PSLPDateTime)
            $PSTableEntry.Insert(2,$PTimeCheckResult)
            $PSTableEntry.Insert(3,$PMountCheckResult)
            $PSTableEntry.Insert(4,$PSCoreName)
            $TableEntry += $PSTableEntry

            $Row = $DataTableFail.NewRow()
            $Row.MachineName = $PSDisplayName
            $Row.LastSnapTime = $PSLPDateTime.ToString()
            $Row.Current = $PTimeCheckResult
            $Row.MountCheck = $PMountCheckResult
            $Row.Core = $PSCoreName
            $DataTableFail.Rows.Add($Row)
                       
            Write-Host "Mount Check for $PSDisplayName - $PMountCheckResult"
            Remove-Mounts
            Continue
        } Else {
            
            $PStopWatch.Start()
            $PSeconds = $PStopWatch.Elapsed.Seconds
            
            # Keep checking the name until the mount process is complete - it can't return a string or have the result "Deleted," which indicates that the mount is not yet complete #
            Do {
                $PMountedRPs = Get-Mounts -ProtectedServer $PSLPHostName
                $PMRPoint = $PMountedRPs[0]
                $PMRPPath = $PMRPoint.MountedVolumeDisplayName
            } Until ((($PMountedRPs -isnot [string]) -and ($PMRPPath -ne "Deleted")) -or ($PSeconds -gt 30))
            $PMountCheckResult = "Pass"
            Write-Host "Mount Check for $PSDisplayName - $PMountCheckResult"
            $PStopWatch.Stop()
            $PStopWatch.Reset()
        }
        
        # Perform a GCI with a Depth of 1 to verify that you can read metadata on the mounted volume, then store the results #
        $PMRPFileList = Get-ChildItem -Path $PMRPPath -Depth 1
        $PMRPItems = $PMRPFileList | Select-Object Name,Length,Directory,FullName,CreationTime,LastWriteTime | Sort-Object -Property FullName
        
        # Write to the console to show that the RP mounted properly #
        Write-Host "The MVDN is $PMRPPath"
        
        # Create an entry for the Passed table #
        $PSTableEntry = [System.Collections.ArrayList]@()
        $PSTableEntry.Insert(0,$PSDisplayName)
        $PSTableEntry.Insert(1,$PSLPDateTime)
        $PSTableEntry.Insert(2,$PTimeCheckResult)
        $PSTableEntry.Insert(3,$PMountCheckResult)
        $PSTableEntry.Insert(4,$PSCoreName)
        $TableEntry += $PSTableEntry

        $Row = $DataTablePass.NewRow()
        $Row.MachineName = $PSDisplayName
        $Row.LastSnapTime = $PSLPDateTime.ToString()
        $Row.Current = $PTimeCheckResult
        $Row.MountCheck = $PMountCheckResult
        $Row.Core = $PSCoreName
        $DataTablePass.Rows.Add($Row)
        
        # Set up the CSV for output and export the GCI data for this server #
        $CSVDateStamp = Get-Date -format "MMddyy-HHmm"
        $CSVFileName = "$PSDisplayName--$CSVDateStamp.csv"
        $CSVPath = "\\srvrdc42\filestorage\documentation\projects\servers\rapidrecovery\export\"
        $CSVFullPath = $CSVPath + $CSVFileName
        $PMRPItems | Export-Csv -Path $CSVFullPath
        
        # Remove any mounted RPs before starting the loop again #
        Remove-Mount -All
    }
}

# Get list of remote servers being replicated to this location and store it in an array, then ensure there are no mounted RPs #
$ReplicatedServers = Get-ReplicatedServers | Sort-Object -Property DisplayName
Remove-Mount -All

# Make sure that there are servers being replicated to this location; otherwise skip this step #
# If the Get-ReplicatedServers doesn't contain anything it will return as a string #
If ($ReplicatedServers -isnot [String]) {

    # Iterate through the list of servers #
    ForEach ($RServer in $ReplicatedServers) {

        # Pull information about each server and store it in various variables #
        $RSDisplayName = $RServer.DisplayName
        <#If ($RSDisplayName -eq "SRVRKBC03") {
            Continue
        }#>
        Write-Host "Checking $RSDisplayName"
        $RSRPoints = Get-RecoveryPoints -ProtectedServer $RSDisplayName -Number All
        $RSLatestPoint = $RSRPoints[0]
        $RSLPHostName = $RSLatestPoint.AgentHostName
        $RSLPRPN = $RSLatestPoint.Number
        $RStopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
        $RSWSeconds = $RStopWatch.Elapsed.Seconds
        Write-Host "Checking Recovery Point Number $RSLPRPN"
        $RSLPDateTime = $RSLatestPoint.DateTimestamp
        If ($RSLPDateTime -gt (Get-Date).AddHours(-24)) {
            $RTimeCheckResult = "Yes"
        } Else {
            $RTimeCheckResult = "No"
        }
        Write-Host "Is it current? $RTimeCheckResult"

        # Replicated servers need this information, which shows the remote core that is responsible for protecting that server #
        $RSMasterCoreInfo = $RServer.MasterCoreInfo
        $RSRemoteCoreName = $RSMasterCoreInfo.HostName
    
        # Check if replicated server is a virtual host for agentless machines; if so, skip it #
        If (($RSDisplayName -eq $RSRemoteCoreName) -and ($RSRPoints -is [string])) {
            Write-Host "$RSDisplayName is a host server and thus not backed up. Moving on."
            Continue
        }
        
        # Look at the volumes being protected and select only volumes larger than 1 GB that aren't recovery partitions #
        $RSLPVIS = $RSLatestPoint.VolumeImageSummaries
        $RSLPSizeMatches = $RSLPVIS.Where{(($_.rawCapacity -gt 1000000000) -and ($_.volumeDisplayName -ne "(Volume Labeled 'RECOVERY')"))}
        $RSLPVDN = $RSLPSizeMatches[0].VolumeDisplayName
    
        # Mount the latest RP according to the data pulled by the script so far #
        $RMountRP = New-Mount -ProtectedServer $RSDisplayName -Rpn $RSLPRPN -Path "C:\Mounts" -Volumes $RSLPVDN

        # If $RMountRP has a value for its Exception property, mark it as having failed the mount check, add its data to the Fail table, and exit the loop #
        # If its Exception is null, mark it as having passed and get the name of the mounted volume #
        If ($RMountRP.Exception -ne $null) {
            $RMountCheckResult = "Fail"
            Write-Host "Oops! Exception detected!"
            Write-Host $RMountRP.Exception
            
            # If the local copy can't mount, connect to the remote core and check the RP there #
            Invoke-Command -ComputerName $RSRemoteCoreName -ArgumentList $RSDisplayName,$RSLPVDN -ScriptBlock {
                Import-Module RapidRecoveryPowerShellModule
                $RSRPoints = Get-RecoveryPoints -ProtectedServer $args[0] -Number All
                $RSLatestPoint = $RSRPoints[0]
                $RSLPRPN = $RSLatestPoint.Number
                $RemoteStopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
                $RemoteSWSeconds = $RemoteStopWatch.Elapsed.Seconds
                New-Mount -ProtectedServer $args[0] -Path "C:\Mounts" -Rpn $RSLPRPN -Volumes $args[1]
                
                # Keep checking the name until the mount process is complete - it can't return a string or have the result "Deleted," which indicates that the mount is not yet complete #
                
                $RemoteStopWatch.Start()
                Do {
                    $RMountedRPs = Get-Mounts -ProtectedServer $args[0]
                    $RMRPoint = $RMountedRPs[0]
                    $RMRPPath = $RMRPoint.MountedVolumeDisplayName
                } Until ((($RMountedRPs -isnot [string]) -and ($RMRPPath -ne "Deleted")) -or ($RemoteSWSeconds -gt 30))
                Remove-Mounts
                $RemoteStopWatch.Stop()
                $RemoteStopWatch.Reset()
            }
            # If the remote core has a valid RP, mark the server as "Fail - Replication Problem" #
            If ($RMRPPath -ne $null) {
                $RMountCheckResult = "Fail - Replication Problem"
            }

            # Create the entry for the Fail table, adding both types of failed mount #
            $RSTableEntry = [System.Collections.ArrayList]@()
            $RSTableEntry.Insert(0,$RSDisplayName)
            $RSTableEntry.Insert(1,$RSLPDateTime)
            $RSTableEntry.Insert(2,$TimeCheckResult)
            $RSTableEntry.Insert(3,$RMountCheckResult)
            $RSTableEntry.Insert(4,$RSRemoteCoreName)
            $TableEntry += $RSTableEntry

            $Row = $DataTableFail.NewRow()
            $Row.MachineName = $RSDisplayName
            $Row.LastSnapTime = $RSLPDateTime.ToString()
            $Row.Current = $TimeCheckResult
            $Row.MountCheck = $RMountCheckResult
            $Row.Core = $RSRemoteCoreName
            $DataTableFail.Rows.Add($Row)
            
            # Write to the console saying that the mount failed, remove any mounted RPs, and exit the loop #            
            Write-Host "Mount Check for $RSDisplayName - $RMountCheckResult"
            Remove-Mount -ProtectedServer $RSDisplayName
            Continue
        } Else {
            
            # Keep checking the name until the mount process is complete - it can't return a string or have the result "Deleted," which indicates that the mount is not yet complete #
            $RStopWatch.Start()
            Do {
                $RMountedRPs = Get-Mounts -ProtectedServer $RSDisplayName
                $RMRPoint = $RMountedRPs[0]
                $RMRPPath = $RMRPoint.MountedVolumeDisplayName
            } Until ((($RMountedRPs -isnot [string]) -and ($RMRPPath -ne "Deleted")) -or ($RSWSeconds -gt 30))
            $RMountCheckResult = "Pass"
            $RStopWatch.Stop()
            $RStopWatch.Reset()
            Write-Host "Mount Check for $RSDisplayName - $RMountCheckResult"
        }

        # Perform a GCI with a Depth of 1 to verify that you can read metadata on the mounted volume, then store the results #
        $RMRPFileList = Get-ChildItem -Path $RMRPPath -Depth 1
        $RMRPItems = $RMRPFileList | Select-Object Name,Length,Directory,FullName,CreationTime,LastWriteTime | Sort-Object -Property FullName
       
        # Write to the console to show that the mount was successful #
        Write-Host "The MVDN is $RMRPPath"
        
        # Construct an entry for the Pass table #
        $RSTableEntry = [System.Collections.ArrayList]@()
        $RSTableEntry.Insert(0,$RSDisplayName)
        $RSTableEntry.Insert(1,$RSLPDateTime)
        $RSTableEntry.Insert(2,$RTimeCheckResult)
        $RSTableEntry.Insert(3,$RMountCheckResult)
        $RSTableEntry.Insert(4,$RSRemoteCoreName)
        $TableEntry += $RSTableEntry

        $Row = $DataTablePass.NewRow()
        $Row.MachineName = $RSDisplayName
        $Row.LastSnapTime = $RSLPDateTime.ToString()
        $Row.Current = $RTimeCheckResult
        $Row.MountCheck = $RMountCheckResult
        $Row.Core = $RSRemoteCoreName
        $DataTablePass.Rows.Add($Row)

        # Use the GCI data to create and export a CSV to the export directory #
        $CSVDateStamp = Get-Date -format "MMddyy-HHmm"
        $CSVFileName = "$RSDisplayName--$CSVDateStamp.csv"
        $CSVPath = "\\file_server_1\servers\rapidrecovery\export\"
        $CSVFullPath = $CSVPath + $CSVFileName
        $RMRPItems | Export-Csv -Path $CSVFullPath
        
        # Remove any mounted RPs before starting the loop again #
        Remove-Mount -ProtectedServer $RSDisplayName
    }
}

# Once the loops are done, construct the results table into HTML, which will serve as the body of an email #
$html = "<table><tr><td>Machines that failed verification are as follows:</td></tr>"
$html += "<tr><td>MachineName</td><td>LastSnapTime</td><td>Current</td><td>MountCheck</td><td>CoreServer</td></tr>"
foreach ($Row in $DataTableFail.Rows) { 
    $html += "<tr><td>" + $Row[0] + "</td><td>" + $Row[1] + "</td><td>" + $Row[2] + "</td><td>" + $Row[3] +"</td><td>" + $Row[4] + "</td></tr>"
}
$html += "<br /><tr><td>Machines that passed verification are as follows:</td></tr>"
$html += "<tr><td>MachineName</td><td>LastSnapTime</td><td>Current</td><td>MountCheck</td><td>CoreServer</td></tr>"
foreach ($Row in $DataTablePass.Rows) { 
    $html += "<tr><td>" + $Row[0] + "</td><td>" + $Row[1] + "</td><td>" + $Row[2] + "</td><td>" + $Row[3] +"</td><td>" + $Row[4] + "</td></tr>"
}
$html += "</table>"

# Note when the script is finished to keep track of how long the script is taking to run #
$ScriptFinishTime = Get-Date

# Use the HTML table and these additional variables to send the results as an email #
$SMTPServer = "mail_server_1.efaithfarm.com"
$EmailSubject = "Verification Table for $PSCoreName"
$EmailBody = "Greetings! <br />Here are the results of the verification for $PSCoreName, which began at $ScriptStartTime and finished at $ScriptFinishTime :<br /><br />" + $html 
Send-MailMessage -SmtpServer $SMTPServer -From "rr_server_1@efaithfarm.com" -To "rrverify@efaithfarm.com" -Subject $EmailSubject -Body $EmailBody -BodyAsHtml

# Write to the console to show that the script is done and when #
Write-Host "Script finished at $ScriptFinishTime"