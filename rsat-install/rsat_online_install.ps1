## Script to install Remote Server Administration Tools (RSAT) ##

# Collect all available tools with Get-WindowsCapability #
$RSATEntries = Get-WindowsCapability -Name RSAT* -Online

# Iterate through list and install each tool user selects #
$ScriptMode = Read-Host "Do you want to choose each tool to install, or just install all of them? (Select/All)"

# If user chooses to install each one, just iterate through and install them all #
If ($ScriptMode -eq "All") {
    ForEach ($Entry in $RSATEntries) {
        $PackageName = $Entry.Name
        Add-WindowsCapability -Online -Name $PackageName
        Write-Host "Installing package: $PackageName"
    }

# If user wants to pick them, then iterate through and prompt for Y/N, only installing on a "Y" #
} ElseIf ($ScriptMode -eq "Select") {
    ForEach ($Entry in $RSATEntries) {
        $PackageName = $Entry.Name
        $PackageSelect = Read-Host "Would you like to install $PackageName? (Y/N)"
        If ($PackageSelect -eq "Y") {
            Add-WindowsCapability -Online -Name $PackageName
            Write-Host "Installing package: $PackageName"
        }
    }
}