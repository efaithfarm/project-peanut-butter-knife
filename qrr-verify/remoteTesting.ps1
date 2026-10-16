$RSDisplayName = "srvrsbmwg03"
$RSRemoteCoreName = "srvrsbmwg02"
$RemoteOutput = Invoke-Command -ComputerName $RSRemoteCoreName -ArgumentList $RSDisplayName -ScriptBlock {
    New-Mount -ProtectedServer $args[0] -Path "C:\Mounts"
    $MountedRPs = Get-Mounts -ProtectedServer $args[0]
    $MRPoint = $MountedRPs[0]
    $MRPPath = $MRPoint.MountedVolumeDisplayName
    Remove-Mounts
}
$MRPPath
