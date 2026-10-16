# Instantiate Arrays #
$mountedRPs = @()
$differentFileList = @()
$rpoints = @()
$tracker = $NULL
$sessionTracker = $NULL

# Check status of script #
if ($tracker -eq $NULL) {

    # Connect to RR Core Server #
    $coreServer = Read-Host "Enter Core server name"
    if ($sessionTracker -eq $NULL){
        Invoke-Command -ComputerName $coreServer -ScriptBlock {
            Import-Module RapidRecoveryPowerShellModule
            # Prompt for server #
            $server = Read-Host "Enter server name"

            # Load the module for Rapid Recovery and mount the last two recovery points #
        
            $rpoints = get-recoverypoints -protectedserver $server -number all
            $latestno = $rpoints.count
            $nextlatestno = $rpoints.count -1
            new-mount -protectedserver $server -rpn $latestno -path "C:\mounts"
            new-mount -protectedserver $server -rpn $nextlatestno -path "C:\mounts\delta"

            # 
            $mountedRPs = get-mounts -protectedserver $server
            $currentRP = $mountedRPs[$mountedRPs.count -2]
            $deltaRP = $mountedRPs[$mountedRPs.count -1]
            $currentpath = $currentRP.mountedvolumedisplayname
            Write-Host $currentpath
            $deltapath = $deltaRP.mountedvolumedisplayname
            Write-Host $deltapath
            $currentFileList = Get-ChildItem -path $currentpath -recurse
            $deltaFileList = Get-ChildItem -path $deltapath -recurse
            $differences = Compare-Object -ReferenceObject $currentFileList -DifferenceObject $deltaFileList
            foreach($item in $differences){
                $path = $item.InputObject.FullName
                Write-Host $path
                #$differentFileList += $path
            }
            #Write-Host $differentFileList

            Remove-Mount -ProtectedServer $server


            <#
            $date = Get-Date
            $time = (Get-Date).AddDays(-30)
            $testFilePath = 'C:\mounts\C__\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs'

            # Build array of files under 30 days old #
            $output = Get-ChildItem $path | Where-Object {$_.LastWriteTime} -lt $time
            $array = $output -split " "

            # CHECK BES LOGS #
            Write-Host `n[*] List of BigFix Logs Recently Created
            foreach ($item in $array) {
                Write-Host $item
            }

            # GET LAST 5 MESSAGES IN THE LAST LOG #
            $mostRecentLog = $path + $array[$array.Count -1]
            $last5 = Get-Content -tail 10 $mostRecentLog
            foreach ($line in $last5) {
                Write-Host $line
            }
            #>
            if (Read-Host "Ready to close Core session? Y or N" -eq "Y"){
                Exit
                #Remove-PSSession -ComputerName $coreServer
                $sessionTracker = 1
            }
        }
    }
    if (Read-Host "Ready to quit verification? Y or N" -eq "Y"){
        $tracker = 1
    }
}