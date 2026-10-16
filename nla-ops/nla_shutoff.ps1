$TargetServer = "*******"
(Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'").SetUserAuthenticationRequired(0)
