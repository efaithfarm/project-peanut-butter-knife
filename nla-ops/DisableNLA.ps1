$TargetServer = Read-Host -Prompt 'Computer FQDN to remove NLA required'
$TargetUser = Read-Host -Prompt 'Username to add to RDU Group'
(Get-WmiObject -class "Win32_TerminalServiceSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer).SetAllowTSConnections(1,1)
(Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'").SetUserAuthenticationRequired(0)
(Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'").SetSecurityLayer(0)
$TargetGroup = [ADSI]"WinNT://$TargetServer/Remote Desktop Users,group"
$UserPath = [ADSI]"WinNT://******.com/$TargetUser,user"
$TargetGroup.Add($UserPath.Path)