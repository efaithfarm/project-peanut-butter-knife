$TargetServer = Read-Host "Enter server name"
Try {
    $TermServSetting = (Get-WmiObject -class "Win32_TerminalServiceSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -ErrorAction Stop) 
} Catch {
    Write-Host "That didn't work. Try again."
}