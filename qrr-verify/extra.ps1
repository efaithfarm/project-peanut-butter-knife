
enter-pssession $coreserver
$differentFileList = @()
import-module rapidrecoverypowershellmodule
$rpoints = get-recoverypoints -protectedserver $server -number all
$latestno = rpoints.count
$nextlatestno = rpoints.count -1
new-mount -protectedserver $server -rpn $latestno -path "C:\mounts"
new-mount -protectedserver $server -rpn $nextlatestno -path "C:\mounts\delta"
$mountedRPs = get-mounts -protectedserver $server
$currentRP = $mountedRPs[$mountedRPs.count -2]
$deltaRP = $mountedRPs[$mountedRPs.count -1]
$currentpath = $currentRP.mountedvolumedisplayname
$deltapath = $deltaRP.mountedvolumedisplayname
$currentFileList = Get-ChildItem -path $currentpath -recurse
$deltaFileList = Get-ChildItem -path $deltapath -recurse
$differences = Compare-Object -ReferenceObject $currentFileList -DifferenceObject $deltaFileList
foreach($item in $differences){
    $path = $item.InputObject.FullName
    $differentFileList += $path
}
