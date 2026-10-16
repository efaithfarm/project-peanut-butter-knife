# Script to manage Remote Access #

#Add-Type -AssemblyName System.Windows.Forms

# Function to enable remote access #
Function RDPEnable {


    [CmdletBinding()]
    
    Param(
        [Parameter()]
        [string]$TargetUser,

        [Parameter()]
        [string]$TargetServer,
        
        [Parameter()]
        [string]$TargetSAMAccountName,

        [Parameter()]
        [string]$TargetLocation,

        [Parameter()]
        [string]$NLASwitch
    )

    # Alter variables based on user input #
    If ($NLASwitch -eq "W") {
        $AuthRequired = 1
        $SecurityLayer = 2

        # Insert code here to check if the computername is in the security group and remove it if so #

        $ComputerNamesFEC = Get-ADGroupMember -Identity "Firewall Enabled Computers" | Select-Object -ExpandProperty SamAccountName
        If ( $ComputerNamesFEC -contains $TargetSAMAccountName ) {
            Write-Host "Firewall enabled on $TargetServer. Disabling now..."
            Remove-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName
        }

    } ElseIf ($NLASwitch -eq "M"){
        $AuthRequired = 0
        $SecurityLayer = 0

        # When setting up a Mac/HTML5 connection, add the desktop to this security group as an alternative to NLA #
        Add-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName

    } Else {

        # If the user gives a typo at the beginning, exit here #
        Write-Host "Oof. Your entry was invalid. Please try again."
        Quit

    }

    # Turn on RDP #
    Try {
        (Get-WmiObject -class "Win32_TerminalServiceSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -ErrorAction Stop).SetAllowTSConnections(1,1)
    } Catch {
        Write-Host "There was an error attempting to reach the named machine. Might want to check that it's online."
        $FailedHosts += $TargetServer
    }
    
    # Enable or disable NLA based on user input #
    If ($FailedHosts -notcontains $TargetServer) {
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetUserAuthenticationRequired($AuthRequired)
        } Catch {
            Write-Host "Oh dear. That didn't work either. Machine probably needs to be pinged."
        }
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetSecurityLayer($SecurityLayer)
        } Catch {
            Write-Host "Yep, it's not responding. Definitely ping it, and write to the user."
        }
    }

    # Add user to local Remote Desktop Users group #
    $TargetGroup = [ADSI]"WinNT://$TargetServer/Remote Desktop Users,group"
    $UserPath = [ADSI]"WinNT://Sewell.com/$TargetUser,user"
    If ($TargetGroup -contains $TargetUser) {
        Write-Host "No addition necessary - $TargetUser is already in the local group."
    } Else {
        $TargetGroup.Add($UserPath.Path)
    }

    # Add user to SSO Group for specified Pulse location #
    If ($TargetLocation -eq "DC") {
        Add-ADGroupMember -Identity "SSO Remote Access Associate" -Members $TargetUser
    } ElseIf ($TargetLocation -eq "DB") {
        Add-ADGroupMember -Identity "SSO Remote Access Data Bank Associate" -Members $TargetUser
    } Else {
        Write-Host "Sorry, that's not a valid location. Better luck next time!"
    }

    Write-Host "Thanks for playing, $TargetUser!"
}

# Function to remove remote access and revert settings #
Function RDPRevert {

    [CmdletBinding()]
    
    Param(
        [Parameter()]
        [string]$TargetUser,

        [Parameter()]
        [string]$TargetServer,
        
        [Parameter()]
        [string]$TargetSAMAccountName,

        [Parameter()]
        [string]$TargetLocation
    )
    
    
    # Search through the SSO Security groups to see if the user is there and remove them if so #

    $DCSSOGroup = Get-ADGroupMember -Identity "SSO Remote Access Associate" | Select-Object -ExpandProperty SamAccountName
    $DBSSOGroup = Get-ADGroupMember -Identity "SSO Remote Access Data Bank Associate" | Select-Object -ExpandProperty SamAccountName
    # find the user in this group - if not, check the other #
    If ($DCSSOGroup -contains $TargetUser) {
        Write-Host "User $TargetUser found in SSO Data Center group. Removing..."
        Remove-ADGroupMember -Identity "SSO Remote Access Associate" -Members $TargetUser -Confirm:$False
    } ElseIf ($DBSSOGroup -contains $TargetUser) {
        Write-Host "User $TargetUser found in SSO Data Bank group. Removing..."
        Remove-ADGroupMember -Identity "SSO Remote Access Data Bank Associate" -Members $TargetUser -Confirm:$False
    } Else {
        Write-Host "User not found! Whaddaya know?"
    }

    # Alter variables based on user input #

    $AuthRequired = 1
    $SecurityLayer = 2

    # Insert code here to check if the computername is in the security group and remove it if so #

    $ComputerNamesFEC = Get-ADGroupMember -Identity "Firewall Enabled Computers" | Select-Object -ExpandProperty SamAccountName
    If ( $ComputerNamesFEC -contains $TargetSAMAccountName ) {
        Write-Host "Firewall enabled on $TargetServer. Disabling now..."
        Remove-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName -Confirm:$False
    }

    # Turn on RDP #
    Try {
        (Get-WmiObject -class "Win32_TerminalServiceSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -ErrorAction Stop).SetAllowTSConnections(0,0)
    } Catch {
        Write-Host "There was an error attempting to reach the named machine. Might want to check that it's online."
        $FailedHosts += $TargetServer
    }
    
    

    # Enable or disable NLA based on user input #
    If ($FailedHosts -notcontains $TargetServer) {
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetUserAuthenticationRequired($AuthRequired)
        } Catch {
            Write-Host "Oh dear. That didn't work either. Machine probably needs to be pinged."
        }
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetSecurityLayer($SecurityLayer)
        } Catch {
            Write-Host "Yep, it's not responding. Definitely ping it, and write to the user."
        }
    }

    # Remove user from local Remote Desktop Users group #
    $TargetGroup = [ADSI]"WinNT://$TargetServer/Remote Desktop Users,group"
    $UserPath = [ADSI]"WinNT://Sewell.com/$TargetUser,user"
    $TargetGroup.Remove($UserPath.Path)

    Write-Host "Thanks for playing, $TargetUser!"
}
<#
$MainWindow = New-Object System.Windows.Forms.Form
$MainWindow.Text = "Remote Management SuperScript"
$MainWindow.Width = 640
$MainWindow.Height = 480
$MainWindow.AutoSize = $True
#>

# Start by asking if this is for individual users or bulk (via CSV Import) #
$ScriptMethod = Read-Host "Greetings! Type 'csv' to import, or press enter for interactive mode"

If ($ScriptMethod -eq "csv") {
    $CSVPath = Read-Host "Lovely! Please enter the path to your CSV file"
    Try {
        $CSVData = Import-Csv -Path $CSVPath
    }
    Catch {
        Throw "Looks like a bad path. Let's start over."
    }

    $ScriptFunction = Read-Host "Data imported. Will you be (1) enabling access or (2) removing it?"

    ForEach ($CSVEntry in $CSVData) {
        $User = $CSVEntry.Username
        $Machine = $CSVEntry.Desktop
        $SAMName = $Machine + "$"
        $Location = $CSVEntry.Location
        $NLA = $CSVEntry.Type
        $FullName = $CSVEntry.FullName

        If ($ScriptFunction -eq 1) {
            If ($NLA -eq "w") { $Type = "Windows" } ElseIf ($NLA -eq "m") { $Type = "Mac" }
            Write-Host "Enabling RDP for $FullName on $Machine. Username is $User and will be set up to access $Location from $Type."
            RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
        }
        ElseIf ($ScriptFunction -eq 2) {
            Write-Host "Removing access for $FullName."
            RDPRevert -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName
        }
        Else {
            Write-Host "Dang, that was a simple one too. Better luck next time."
            Quit
        }

    }

}

ElseIf ($ScriptMethod -eq "") {
    # Gather variables from user prompts #
    $MNLA = Read-Host -Prompt "Welcome! Enter W for access from Windows, or M for access from Mac"
    $MMachine = Read-Host -Prompt "Computer name"
    $MSAMName = $MMachine + "$"
    $MUser = Read-Host -Prompt "Username"
    $MLocation = Read-Host -Prompt "Location (DC or DB)"
    $MScriptFunction = Read-Host -Prompt "Will you be (1) enabling access or (2) removing it?"

    If ($MScriptFunction -eq 1) {
        RDPEnable -TargetUser $MUser -TargetServer $MMachine -TargetLocation $MLocation -TargetSAMAccountName $MSAMName -NLASwitch $MNLA
    }
    ElseIf ($MScriptFunction -eq 2) {
        RDPRevert -TargetUser $MUser -TargetServer $MMachine -TargetLocation $MLocation -TargetSAMAccountName $MSAMName
    }
    Else {
        Write-Host "Welp, I guess you're not interested in using the script after all. Give it another try!"
        Quit
    }
}

Else {
    Write-Host "Hmm. Might be a keyboard malfunction. Try again!"
    Quit
}

#$MainWindow.ShowDialog()