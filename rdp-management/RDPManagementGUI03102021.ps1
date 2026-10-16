# GUI Shell for RDPManagement.ps1 script #

# This line adds the Forms objects #
Add-Type -AssemblyName System.Windows.Forms
$RestartCounter = 0
$FailedHosts = @()

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
    If (($NLASwitch -eq "W") -or ($NLASwitch -eq "TS")) {
        $AuthRequired = 1
        $SecurityLayer = 2

        # Insert code here to check if the computername is in the security group and remove it if so #

        $ComputerNamesFEC = Get-ADGroupMember -Identity "Firewall Enabled Computers" | Select-Object -ExpandProperty SamAccountName
        If ( $ComputerNamesFEC -contains $TargetSAMAccountName ) {
            Write-Host "Firewall enabled on $TargetServer. Disabling now..."
            Remove-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName
        }

    } ElseIf (($NLASwitch -eq "M") -or ($NLASwitch -eq "HTML5")){
        $AuthRequired = 0
        $SecurityLayer = 0

        # When setting up a Mac/HTML5 connection, add the desktop to this security group as an alternative to NLA #
        Add-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName

    } Else {

        # If the user gives a typo at the beginning, exit here #
        Write-Host "This is an internal-only user, so no NLASwitching necessary. But NLA should be on anyway, so..."
        $AuthRequired = 1
        $SecurityLayer = 2

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
    Try {
        $TargetGroup = [ADSI]"WinNT://$TargetServer/Remote Desktop Users,group"
        $UserPath = [ADSI]"WinNT://Sewell.com/$TargetUser,user"
        If ($TargetGroup -contains $TargetUser) {
            Write-Host "No addition necessary - $TargetUser is already in the local group."
        } Else {
            $TargetGroup.Add($UserPath.Path)
        }
    } Catch {
        Write-Host "The local groups couldn't be edited either. You know, because the machine is offline."
    }
    

    # Add user to SSO Group for specified Pulse location #
    If (($TargetLocation -eq "DC") -or ($TargetLocation -eq "SSO Remote Access Associate")) {
        Add-ADGroupMember -Identity "SSO Remote Access Associate" -Members $TargetUser
    } ElseIf (($TargetLocation -eq "DB") -or ($TargetLocation -eq "SSO Remote Access Data Bank Associate")) {
        Add-ADGroupMember -Identity "SSO Remote Access Data Bank Associate" -Members $TargetUser
    } Else {
        Write-Host "This user is an internal-only user, so this is another unnecessary parameter. Cheers!"
    }

    Write-Host "Welcome to the world of Remote Access, $TargetUser!"
}

# Function to remove remote access and revert settings #
Function RDPRemove {

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

    # Turn off RDP #
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
    Try {
        $TargetGroup = [ADSI]"WinNT://$TargetServer/Remote Desktop Users,group"
        $UserPath = [ADSI]"WinNT://Sewell.com/$TargetUser,user"
        If ($TargetGroup -contains $TargetUser) {
            Write-Host "No addition necessary - $TargetUser is already in the local group."
        } Else {
            $TargetGroup.Remove($UserPath.Path)
        }
    } Catch {
        Write-Host "The local groups couldn't be edited either. You know, because the machine is offline."
    }

    Write-Host "Better luck next time, $TargetUser!"
}

# WINDOWS FORMS CODE TO CREATE THE GUI #

# Function to set up the form #

Function CreateForm {

    # Establish the main form window and define its dimensions #
    $MainWindow = New-Object System.Windows.Forms.Form
    $MainWindow.Text = "Remote Management SuperScript"
    $MainWindow.Width = 640
    $MainWindow.Height = 480
    $MainWindow.AutoSize = $True
    $MainWindow.Padding = 10

    # STEP ONE #
    # Start by determining whether this is bulk processing or individual processing, by means of two radio buttons and a Next button #

    $MethodGroup = New-Object System.Windows.Forms.GroupBox
    $MethodGroup.Location = '10,10'
    $MethodGroup.Text = "Is this for individual users, or batch processing with a CSV?"
    $MethodGroup.Size = '440,60'

    # Give the group a name so that it can be acted upon and removed when necessary #
    $MethodGroup.Name = "MethodGroup"

    # Radio button indicating Individual mode #
    $MethodRadio1 = New-Object System.Windows.Forms.RadioButton
    $MethodRadio1.Location = '20,20'
    $MethodRadio1.Text = "Individual"
    $MethodRadio1.Checked = $True

    # Radio button indicating CSV Import mode #
    $MethodRadio2 = New-Object System.Windows.Forms.RadioButton
    $MethodRadio2.Location = '180,20'
    $MethodRadio2.Text = "Import CSV"
    $MethodRadio2.Checked = $False

    # Add the radio buttons to the GroupBox #
    $MethodGroup.Controls.AddRange(@($MethodRadio1,$MethodRadio2))

    # The Next button moves the GUI to the next step in the process #
    $NextButton = New-Object System.Windows.Forms.Button
    $NextButton.Location = '460,20'
    $NextButton.Size = '120,23'
    $NextButton.Text = "Next"
    $NextButton.Name = "NextButton"

    # Add the initial $MethodGroup GroupBox and $NextButton to the Form #
    $MainWindow.Controls.Add($MethodGroup)
    $MainWindow.Controls.Add($NextButton)

    # Create the GroupBox for the CSV Browse dialog #
    $CSVGroup = New-Object System.Windows.Forms.GroupBox
    $CSVGroup.Location = '10,10'
    $CSVGroup.AutoSize = $True
    $CSVGroup.Name = "CSVGroup"
    $CSVGroup.Text = "Import CSV File"
    $CSVGroup.Padding = 5

    # Instruction label for the CSV Browse dialog #
    $CSVBrowseLabel = New-Object System.Windows.Forms.Label
    $CSVBrowseLabel.Location = '10,20'
    $DTCSVBrowseLabel = "Click the button to browse and find a CSV file."
    $CSVBrowseLabel.Text = $DTCSVBrowseLabel
    $CSVBrowseLabel.AutoSize = $True

    # Instruction label that will show the path chosen after the dialog is completed #
    $CSVResultPathLabel = New-Object System.Windows.Forms.Label
    $CSVResultPathLabel.Location = '10,40'
    $DTCSVResultPathLabel = "Please ensure that your CSV file has the proper fields."
    $CSVResultPathLabel.Text = $DTCSVResultPathLabel
    $CSVResultPathLabel.AutoSize = $True

    # Button that confirms user choice of CSV file and moves the script to the next step on click #
    $CSVConfirmButton = New-Object System.Windows.Forms.Button
    $CSVConfirmButton.Text = "Confirm"
    $CSVConfirmButton.Location = '450,60'
    $CSVConfirmButton.Size = '100,23'
    $CSVConfirmButton.Name = "CSVConfirmButton"

    # Button that cancels user choice of CSV file and brings the Browse button back #
    $CSVCancelButton = New-Object System.Windows.Forms.Button
    $CSVCancelButton.Text = "Cancel"
    $CSVCancelButton.Location = '570,60'
    $CSVCancelButton.Size = '100,23'
    $CSVCancelButton.Name = "CSVCancelButton"

    # Read-only, multiline textbox with scrollbars to show output of the Import-CSV process #
    $CSVOutputBox = New-Object System.Windows.Forms.Textbox
    $CSVOutputBox.Text = ""
    $CSVOutputBox.Location = '10,130'
    $CSVOutputBox.Size = '600,200'
    $CSVOutputBox.Multiline = $True
    $CSVOutputBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $CSVOutputBox.ReadOnly = $True

    # Code to bring up the file browser object #
    $BrowseDialog = New-Object System.Windows.Forms.OpenFileDialog
    $BrowseDialog.Filter = "CSV Files (*.csv)| *.csv"
    $BrowseDialog.Title = "Choose CSV File"

    # Button that brings out the file browser object when user clicks on it #
    $BrowseButton = New-Object System.Windows.Forms.Button
    $BrowseButton.Text = "Select CSV"
    $BrowseButton.Location = '460,20'
    $BrowseButton.Size = '120,23'
    $BrowseButton.Name = "BrowseButton"

    # Group to hold controls for Individual mode #
    $SingleModeGroup = New-Object System.Windows.Forms.GroupBox
    $SingleModeGroup.Text = "Enter user and desktop information"
    $SingleModeGroup.Location = '10,10'
    $SingleModeGroup.Name = "SingleModeGroup"
    $SingleModeGroup.AutoSize = $True
    #$SingleModeGroup.Size = "580,420"
    $SingleModeGroup.Padding = 10

    # Group to hold controls for choosing Direction (internal or external) #
    $SMGDirectionGroup = New-Object System.Windows.Forms.GroupBox
    $SMGDirectionGroup.Name = "SMGDirectionGroup"
    $SMGDirectionGroup.Location = '15,15'
    $SMGDirectionGroup.Size = '550,80'

    # Instruction label for choosing Direction #
    $SMGDirectionInstLabel = New-Object System.Windows.Forms.Label
    $SMGDirectionInstLabel.Location = '10,15'
    $SMGDirectionInstLabel.Text = "Will the user be connecting from home or while in the office?"
    $SMGDirectionInstLabel.AutoSize = $True

    # First checkbox/button for SMGDirection #
    $SMGDirectionButtonHome = New-Object System.Windows.Forms.CheckBox
    $SMGDirectionButtonHome.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGDirectionButtonHome.Location = '285,45'
    $SMGDirectionButtonHome.Size = '120,23'
    $SMGDirectionButtonHome.Text = "Home"
    $SMGDirectionButtonHome.Name = "SMGDirectionButtonHome"

    # Second checkbox/button for SMGDirection #
    $SMGDirectionButtonWork = New-Object System.Windows.Forms.CheckBox
    $SMGDirectionButtonWork.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGDirectionButtonWork.Location = '420,45'
    $SMGDirectionButtonWork.Size = '120,23'
    $SMGDirectionButtonWork.Text = "Office"
    $SMGDirectionButtonWork.Name = "SMGDirectionButtonWork"

    # Group to hold controls for choosing Type (Windows or Mac) #
    $SMGTypeGroup = New-Object System.Windows.Forms.GroupBox
    $SMGTypeGroup.Location = '15,15'
    #$SMGTypeGroup.AutoSize = $True
    $SMGTypeGroup.Size = '550,50'
    $SMGTypeGroup.Name = "SMGTypeGroup"

    # Instruction label for SMGType #
    $SMGTypeLabel = New-Object System.Windows.Forms.Label
    $SMGTypeLabel.Location = '10,15'
    $SMGTypeLabel.Text = "Will the user be connecting from Windows or Mac?"
    $SMGTypeLabel.AutoSize = $True

    # First checkbox/button for SMGType #
    $SMGTypeButtonWin = New-Object System.Windows.Forms.CheckBox
    $SMGTypeButtonWin.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGTypeButtonWin.Location = '285,15'
    $SMGTypeButtonWin.Size = '120,23'
    $SMGTypeButtonWin.Text = "Windows"
    $SMGTypeButtonWin.Name = "SMGTypeButtonWin"

    # Second checkbox/button for SMGType #
    $SMGTypeButtonMac = New-Object System.Windows.Forms.CheckBox
    $SMGTypeButtonMac.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGTypeButtonMac.Location = '420,15'
    $SMGTypeButtonMac.Size = '120,23'
    $SMGTypeButtonMac.Text = "Mac"
    $SMGTypeButtonMac.Name = "SMGTypeButtonMac"

    # Group to hold controls for entering Username information #
    $SMGUserGroup = New-Object System.Windows.Forms.GroupBox
    $SMGUserGroup.Location = '15,60'
    $SMGUserGroup.Size = '550,100'
    $SMGUserGroup.Name = "SMGUserGroup"

    # Instruction label for SMGUser #
    $SMGUserInstLabel = New-Object System.Windows.Forms.Label
    $SMGUserInstLabel.Location = '10,15'
    $SMGUserInstLabel.AutoSize = $True
    $SMGUserInstLabel.Text = "Enter the associate's username:"
    $SMGUserInstLabel.Name = "SMGUserInstLabel"

    # Input text box where user inputs the associate's username #
    $SMGUserTextBox = New-Object System.Windows.Forms.TextBox
    $SMGUserTextBox.Location = '185,17'
    $SMGUserTextBox.Size = '100,20'
    $SMGUserTextBox.Name = "SMGUserTextBox"

    # Button user clicks to check if the user is the correct one #
    $SMGUserButtonVerify = New-Object System.Windows.Forms.Button
    $SMGUserButtonVerify.Location = '295,15'
    $SMGUserButtonVerify.Size = '120,23'
    $SMGUserButtonVerify.Text = "Find User"
    $SMGUserButtonVerify.Name = "SMGUserButtonVerify"

    # Label for results of user search #
    $SMGUserResultLabel = New-Object System.Windows.Forms.Label
    $SMGUserResultLabel.Location = '10,45'
    $SMGUserResultLabel.AutoSize = $True
    $SMGUserResultLabel.Text = ""
    $SMGUserResultLabel.Name = "SMGUserResultLabel"

    # Button to confirm user selection #
    $SMGUserButtonConfirm = New-Object System.Windows.Forms.Button
    $SMGUserButtonConfirm.Location = '295,70'
    $SMGUserButtonConfirm.Size = '120,23'
    $SMGUserButtonConfirm.Text = "Yes"
    $SMGUserButtonConfirm.Name = "SMGUserButtonConfirm"

    # Button to cancel user selection #
    $SMGUserButtonCancel = New-Object System.Windows.Forms.Button
    $SMGUserButtonCancel.Location = '425,70'
    $SMGUserButtonCancel.Size = '120,23'
    $SMGUserButtonCancel.Text = "No"
    $SMGUserButtonCancel.Name = "SMGUserButtonCancel"

    # Group to hold Desktop selection controls #
    $SMGDesktopGroup = New-Object System.Windows.Forms.GroupBox
    $SMGDesktopGroup.Location = '15,105'
    $SMGDesktopGroup.Size = '550,150'
    $SMGDesktopGroup.Name = "SMGDesktopGroup"

    # Label giving instructions for desktop entry #
    $SMGDesktopInstLabel = New-Object System.Windows.Forms.Label
    $SMGDesktopInstLabel.Location = '10,15'
    $SMGDesktopInstLabel.Text = "Enter the associate's office PC name:"
    $SMGDesktopInstLabel.AutoSize = $True

    # TextBox for user to enter associate's desktop computername #
    $SMGDesktopTextBox = New-Object System.Windows.Forms.TextBox
    $SMGDesktopTextBox.Location = '225,17'
    $SMGDesktopTextBox.Size = '100,20'
    $SMGDesktopTextBox.Text = ""

    # Button to begin search for entered desktop name #
    $SMGDesktopButtonVerify = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonVerify.Location = '335,15'
    $SMGDesktopButtonVerify.Size = '120,23'
    $SMGDesktopButtonVerify.Text = "Find PC"

    # Label to display results of Get-ADComputer (which runs as result of the SMGDesktopButtonVerify's click) #
    $SMGDesktopResultLabel = New-Object System.Windows.Forms.Label
    $SMGDesktopResultLabel.Location = '10,45'
    $SMGDesktopResultLabel.AutoSize = $True
    $SMGDesktopResultLabel.Text = ""
    $SMGDesktopResultLabel.MaximumSize = '500,50'
    $SMGDesktopResultLabel.Name = "SMGDesktopResultLabel"

    # Button to confirm Desktop selection #
    $SMGDesktopButtonConfirm = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonConfirm.Location = '295,85'
    $SMGDesktopButtonConfirm.Size = '120,23'
    $SMGDesktopButtonConfirm.Text = "Yes"
    $SMGDesktopButtonConfirm.Name = "SMGDesktopButtonConfirm"

    # Button to cancel Desktop selection #
    $SMGDesktopButtonCancel = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonCancel.Location = '425,85'
    $SMGDesktopButtonCancel.Size = '120,23'
    $SMGDesktopButtonCancel.Text = "No"
    $SMGDesktopButtonCancel.Name = "SMGDesktopButtonCancel"

    # Group to hold Location selection controls #
    $SMGLocGroup = New-Object System.Windows.Forms.GroupBox
    $SMGLocGroup.Location = '15,150'
    $SMGLocGroup.Size = '550,80'
    $SMGLocGroup.Name = "SMGLocGroup"

    # Instruction label for SMGLoc #
    $SMGLocLabel = New-Object System.Windows.Forms.Label
    $SMGLocLabel.Location = '10,15'
    $SMGLocLabel.Text = "Will the user be connecting to the Data Center or Data Bank?"
    $SMGLocLabel.AutoSize = $True

    # First checkbox/button for SMGLoc #
    $SMGLocButtonDC = New-Object System.Windows.Forms.CheckBox
    $SMGLocButtonDC.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGLocButtonDC.Location = '285,45'
    $SMGLocButtonDC.Size = '120,23'
    $SMGLocButtonDC.Text = "Data Center"
    $SMGLocButtonDC.Name = "SMGLocButtonDC"

    # Second checkbox/button for SMGLoc #
    $SMGLocButtonDB = New-Object System.Windows.Forms.CheckBox
    $SMGLocButtonDB.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGLocButtonDB.Location = '420,45'
    $SMGLocButtonDB.Size = '120,23'
    $SMGLocButtonDB.Text = "Data Bank"
    $SMGLocButtonDB.Name = "SMGLocButtonDB"
    
    <#
    # Group to hold Email selection controls #
    $SMGEmailGroup = New-Object System.Windows.Forms.GroupBox
    $SMGEmailGroup.Location = '15,150'
    $SMGEmailGroup.Size = '550,80'
    $SMGEmailGroup.Name = "SMGEmailGroup"

    # Instruction label for SMGEmail #
    $SMGEmailLabel = New-Object System.Windows.Forms.Label
    $SMGEmailLabel.Location = '10,15'
    $SMGEmailLabel.Text = "Should an email be sent to notify the user?"
    $SMGEmailLabel.AutoSize = $True

    # First checkbox/button for SMGEmail #
    $SMGEmailButtonDC = New-Object System.Windows.Forms.CheckBox
    $SMGEmailButtonDC.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGEmailButtonDC.Location = '285,45'
    $SMGEmailButtonDC.Size = '120,23'
    $SMGEmailButtonDC.Text = "Data Center"
    $SMGEmailButtonDC.Name = "SMGEmailButtonDC"

    # Second checkbox/button for SMGEmail #
    $SMGEmailButtonDB = New-Object System.Windows.Forms.CheckBox
    $SMGEmailButtonDB.Appearance = [System.Windows.Forms.Appearance]::Button
    $SMGEmailButtonDB.Location = '420,45'
    $SMGEmailButtonDB.Size = '120,23'
    $SMGEmailButtonDB.Text = "Data Bank"
    $SMGEmailButtonDB.Name = "SMGEmailButtonDB"
    #>

    # Read-only, multiline textbox with scrollbars to show output of the SingleMode process #
    $SMGOutputBox = New-Object System.Windows.Forms.Textbox
    $SMGOutputBox.Text = ""
    $SMGOutputBox.Location = '15,235'
    $SMGOutputBox.Size = '550,75'
    $SMGOutputBox.Multiline = $True
    $SMGOutputBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $SMGOutputBox.ReadOnly = $True

    # Group to hold radio buttons for Function selection #
    $FunctionGroup = New-Object System.Windows.Forms.GroupBox
    $FunctionGroup.Text = "Choose function"
    $FunctionGroup.Location = '10,350'
    $FunctionGroup.AutoSize = $True
    $FunctionGroup.Name = "FunctionGroup"

    # Radio button to choose Enable #
    $FunctionRadioEnable = New-Object System.Windows.Forms.RadioButton
    $FunctionRadioEnable.Location = '20,20'
    $FunctionRadioEnable.Checked = $False
    $FunctionRadioEnable.Text = "Enable Access"

    # Radio button to choose Remove #
    $FunctionRadioRemove = New-Object System.Windows.Forms.RadioButton
    $FunctionRadioRemove.Location = '180,20'
    $FunctionRadioRemove.Checked = $False
    $FunctionRadioRemove.Text = "Remove Access"

    # Button to submit choice and call function #
    $FunctionSubmitButton = New-Object System.Windows.Forms.Button
    $FunctionSubmitButton.Location = '460,20'
    $FunctionSubmitButton.Size = '120,23'
    $FunctionSubmitButton.Text = "Submit"
    $FunctionSubmitButton.Name = "FunctionSubmitButton"

    # Group for controls that appear after the script has run, allowing you to restart or exit #
    $ResultGroup = New-Object System.Windows.Forms.GroupBox
    $ResultGroup.Text = "Would you like to start over or quit?"
    $ResultGroup.Location = '10,350'
    $ResultGroup.AutoSize = $True
    $ResultGroup.Name = "ResultGroup"

    # Button to restart script when results are shown #
    $ResultButtonRestart = New-Object System.Windows.Forms.Button
    $ResultButtonRestart.Location = '20,20'
    $ResultButtonRestart.Text = "Restart script"
    $ResultButtonRestart.Size = '120,23'

    # Button to exit script when results are shown #
    $ResultButtonQuit = New-Object System.Windows.Forms.Button
    $ResultButtonQuit.Location = '180,20'
    $ResultButtonQuit.Text = "Quit"
    $ResultButtonQuit.Size = '120,23'
    $ResultButtonQuit.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # BUTTON CLICKS AND SUCH #

    # Logic for the Home button in the SMGDirectionGroup #
    $SMGDirectionButtonHome.Add_CheckStateChanged(
        {
            If ($SMGDirectionButtonHome.Checked) {
                # Set the Type variable for the function to "w" #
                $SCRIPT:SMGDirection = "Home"
                Write-Host "Direction is set to $SMGDirection, so this is an external user."
                $SMGDirectionGroup.Controls.RemoveByKey("SMGDirectionButtonWork")
            
                # Clear out the DirectionGroup and add the TypeGroup #
                $SingleModeGroup.Controls.Clear()
            
                # Add label and buttons to the SMGTypeGroup #
                $SMGTypeGroup.Controls.AddRange(@($SMGTypeLabel,$SMGTypeButtonWin,$SMGTypeButtonMac))
            
                # Add SMGTypeGroup to the SingleModeGroup #
                $SingleModeGroup.Controls.AddRange(@($SMGTypeGroup))
            } Else {
                $SMGDirectionGroup.Controls.Add($SMGDirectionButtonWork)
                $SCRIPT:SMGDirection = ""
            }
        }
    )

    # Logic for the Work button in the SMGDirectionGroup #
    $SMGDirectionButtonWork.Add_CheckStateChanged(
        {
            If ($SMGDirectionButtonWork.Checked) {
                # Set the Type variable for the function to "m" #
                $SCRIPT:SMGDirection = "Work"
                Write-Host "Direction is set to $SMGDirection, so this is an internal-only user."
                $SMGDirectionGroup.Controls.RemoveByKey("SMGDirectionButtonHome")

                # Clear out the DirectionGroup and add the UserGroup (skipping TypeGroup because it's irrelevant for IO users) #
                $SingleModeGroup.Controls.Clear()
            
                # Add label and buttons to the SMGUserGroup #
                $SMGUserGroup.Controls.AddRange(@($SMGUserInstLabel,$SMGUserTextBox,$SMGUserButtonVerify,$SMGUserResultLabel))
            
                # Change the location for the User GroupBox and add SMGUserGroup to the SingleModeGroup #
                $SMGUserGroup.Location = '15,15'
                $SingleModeGroup.Controls.AddRange(@($SMGUserGroup))
            } Else {
                $SMGDirectionGroup.Controls.Add($SMGDirectionButtonHome)
                $SCRIPT:SMGDirection = ""
            }
        }
    )

    # Logic for the Windows button in the SMGTypeGroup #
    $SMGTypeButtonWin.Add_CheckStateChanged(
        {
            If ($SMGTypeButtonWin.Checked) {
                # Set the Type variable for the function to "w" #
                $SCRIPT:NLASwitch = "w"
                Write-Host "NLASwitch is set to $NLASwitch, so this is a Windows user."
                $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonMac")
                If (!($SingleModeGroup.Contains($SMGUserGroup))) {
                    $SMGUserGroup.Controls.AddRange(@($SMGUserInstLabel,$SMGUserTextBox,$SMGUserButtonVerify,$SMGUserResultLabel))
                    $SingleModeGroup.Controls.Add($SMGUserGroup)
                    $SMGUserTextBox.Focus()
                } Else {
                    $SMGUserTextBox.Focus()
                }
            } Else {
                $SMGTypeGroup.Controls.Add($SMGTypeButtonMac)
                $SCRIPT:NLASwitch = ""
            }
        }
    )

    # Logic for the Mac button in the SMGTypeGroup #
    $SMGTypeButtonMac.Add_CheckStateChanged(
        {
            If ($SMGTypeButtonMac.Checked) {
                # Set the Type variable for the function to "m" #
                $SCRIPT:NLASwitch = "m"
                Write-Host "NLASwitch is set to $NLASwitch, so this is a Mac user."
                $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonWin")
                If (!($SingleModeGroup.Contains($SMGUserGroup))) {
                    $SMGUserGroup.Controls.AddRange(@($SMGUserInstLabel,$SMGUserTextBox,$SMGUserButtonVerify,$SMGUserResultLabel))
                    $SingleModeGroup.Controls.Add($SMGUserGroup)
                    $SMGUserTextBox.Focus()
                } Else {
                    $SMGUserTextBox.Focus()
                }
            } Else {
                $SMGTypeGroup.Controls.Add($SMGTypeButtonWin)
                $SCRIPT:NLASwitch = ""
            }
        }
    )

    # Logic to find the user entered in the Single Mode Username text box #
    $SMGUserButtonVerify.Add_Click(
        {
            # First, set the text box to Read-only until further action is taken
            $SMGUserTextBox.ReadOnly = $True
        
            # Require input :P #
            If ($SMGUserTextBox.Text -ne "") {
                $SMGUserNameEntry = $SMGUserTextBox.Text
                $SMGUserSuccess = ""
                $SMGUserResultLabel.Text = "Working..."

                # Try using the inputted username as the -Identity parameter of a Get-AdUser cmdlet #
                Try {
                    $SCRIPT:SMGUser = Get-ADUser -Identity $SMGUserNameEntry -Properties *
                    $SMGUserSuccess = "YEP"
                } Catch {
                
                    # If it doesn't work, reset the form and display a message asking nicely for the user to GET IT RIGHT NEXT TIME YOU DUMMY #
                    $SMGUserResultLabel.Text = "Hmm, that entry didn't quite work. Check the username and try again!"
                    Start-Sleep 2
                    $SMGUserTextBox.ReadOnly = $False
                    $SMGUserSuccess = "NOPE"
                }

                # Make sure the Get-ADUser works before continuing #
                If ($SMGUserSuccess = "YEP") {
                
                    # Show the user information on the associate in question and ask for confirmation #
                    $SCRIPT:SMGUserName = $SMGUser.SamAccountName
                    $SCRIPT:SMGUserEmpID = $SMGUser.EmployeeID
                    $SCRIPT:SMGUserFullName = $SMGUser.DisplayName
                    $SCRIPT:SMGUserDepartment = $SMGUser.Department
                    $SMGUserResultLabel.Text = "$SMGUserFullName, employee ID $SMGUserEmpID, with $SMGUserDepartment. Is this correct?"
                    $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
                    $SMGUserGroup.Controls.AddRange(@($SMGUserButtonConfirm,$SMGUserButtonCancel))
                }

            } Else {
                $SMGUserResultLabel.Text = "Sorry, input is required here. Please try again!"
                Start-Sleep 2
                $SMGUserTextBox.ReadOnly = $False
            }
        }
    )

    # Logic for the User Group Confirm button #
    $SMGUserButtonConfirm.Add_Click(
        {
            $SMGUserGroup.Controls.Clear()
            $SMGUserTextBox.Location = '10,15'
            $SMGUserGroup.Controls.Add($SMGUserTextBox)
            $SMGUserResultLabel.Text = "$SMGUserFullName, employee ID $SMGUserEmpID, with $SMGUserDepartment."
            $SMGUserResultLabel.Location = '120,17'
            $SMGUserGroup.Controls.Add($SMGUserResultLabel)
            $SMGUserGroup.Size = '550,50'
            If ($SMGDirection -eq "Work") {
                $SMGDesktopGroup.Location = '15,60'
            }
            $SMGDesktopGroup.Controls.AddRange(@($SMGDesktopInstLabel,$SMGDesktopTextBox,$SMGDesktopButtonVerify))
            #$SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDC,$SMGLocButtonDB))
            $SingleModeGroup.Controls.Add($SMGDesktopGroup)
            $SMGDesktopTextBox.Focus()
        }
    )

    # Logic needed for the User Cancel button #
    $SMGUserButtonCancel.Add_Click(
        {
            $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonConfirm")
            $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonCancel")
            $SMGUserResultLabel.Text = ""
            $SMGUserTextBox.Text = ""
            $SMGUserTextBox.ReadOnly = $False
            $SMGUserGroup.Controls.Add($SMGUserButtonVerify)
        }
    )

    # Logic to find the user entered in the Single Mode Username text box #
    $SMGDesktopButtonVerify.Add_Click(
        {
            # First, set the text box to Read-only until further action is taken
            $SMGDesktopTextBox.ReadOnly = $True
        
            # Require input :P #
            If ($SMGDesktopTextBox.Text -ne "") {
                $SMGDesktopNameEntry = $SMGDesktopTextBox.Text
                $SMGDesktopSuccess = ""
                $SMGDesktopResultLabel.Text = "Working..."

                # Try using the inputted PC name as the -Identity parameter of a Get-ADComputer cmdlet #
                Try {
                    $SCRIPT:SMGDesktop = Get-ADComputer -Identity $SMGDesktopNameEntry -Properties *
                    $SMGDesktopSuccess = "YEP"
                } Catch {
                
                    # If it doesn't work, reset the form and display a message asking nicely for the user to GET IT RIGHT NEXT TIME YOU DUMMY #
                    $SMGDesktopResultLabel.Text = "Hmm, that entry didn't quite work. Check the computername and try again!"
                    Start-Sleep 2
                    $SMGDesktopTextBox.ReadOnly = $False
                    $SMGDesktopSuccess = "NOPE"
                }

                # Make sure the Get-ADUser works before continuing #
                If ($SMGDesktopSuccess = "YEP") {
                
                    # Show the desktop information on the PC in question and ask for confirmation #
                    $SCRIPT:SMGDesktopName = $SMGDesktop.Name
                    $SCRIPT:SMGDesktopSamAccountName = $SMGDesktop.SamAccountName
                    $SCRIPT:SMGDesktopDept = $SMGDesktop.Department
                    $SCRIPT:SMGDesktopDesc = $SMGDesktop.Description
                    $SCRIPT:SMGDesktopIP = $SMGDesktop.IPv4Address
                    $SCRIPT:SMGDesktopCo = $SMGDesktop.Company

                    # Do a ping test to make sure the computer is online - let the user know if it fails #
                    $SMGPingResult = ping $SMGDesktopName
                    If (($SMGPingResult -like "reply from*") -like "false*") {
                    
                    }

                    $SMGDesktopResultLabel.Text = "$SMGDesktopName, described as $SMGDesktopDesc, in $SMGDesktopDept at $SMGDesktopCo, with IP address $SMGDesktopIP. Is this correct?"
                    $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonVerify")
                    $SMGDesktopGroup.Controls.AddRange(@($SMGDesktopResultLabel,$SMGDesktopButtonConfirm,$SMGDesktopButtonCancel))
                }

            } Else {
                $SMGDesktopResultLabel.Text = "Sorry, input is required here. Please try again!"
                Start-Sleep 2
                $SMGDesktopTextBox.ReadOnly = $False
            }
        }
    )

    # Logic for the Desktop Group Confirm button #
    $SMGDesktopButtonConfirm.Add_Click(
        {
            $SMGDesktopGroup.Controls.Clear()
            $SMGDesktopTextBox.Location = '10,15'
            $SMGDesktopGroup.Controls.Add($SMGDesktopTextBox)
            $SMGDesktopResultLabel.Text = "$SMGDesktopName, described as $SMGDesktopDesc, in $SMGDesktopDept at $SMGDesktopCo, with IP address $SMGDesktopIP."
            $SMGDesktopResultLabel.Location = '120,17'
            $SMGDesktopResultLabel.MaximumSize = '380,70'
            $SMGDesktopGroup.Controls.Add($SMGDesktopResultLabel)
            $SMGDesktopGroup.Size = '550,50'
            If ($SMGDirection -eq "Work") {
                # Bring up the FunctionGroup to choose between adding and removing access #
                $FunctionGroup.Controls.AddRange(@($FunctionRadioEnable,$FunctionRadioRemove,$FunctionSubmitButton))
                $FunctionSubmitButton.Location = '452,20'
                $SMGOutputBox.Location = '15,115'
                $SingleModeGroup.Controls.Add($SMGOutputBox)
                $MainWindow.Controls.Add($FunctionGroup)
            } Else {
                $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDC,$SMGLocButtonDB))
                $SingleModeGroup.Controls.Add($SMGLocGroup)
            }
        }
    )

    # Logic needed for the Desktop Cancel button #
    $SMGDesktopButtonCancel.Add_Click(
        {
            $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonConfirm")
            $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonCancel")
            $SMGDesktopResultLabel.Text = ""
            $SMGDesktopTextBox.Text = ""
            $SMGDesktopTextBox.ReadOnly = $False
            $SMGDesktopGroup.Controls.Add($SMGDesktopButtonVerify)
        }
    )

    # Logic for the DC button in the SMGLocGroup #
    $SMGLocButtonDC.Add_CheckStateChanged(
        {
            If ($SMGLocButtonDC.Checked) {
                # Set the Type variable for the function to "w" #
                $SCRIPT:SMGLocation = "DC"
                Write-Host "Location is set to $SMGLocation, so this is a Data Center user."
                $SMGLocGroup.Controls.RemoveByKey("SMGLocButtonDB")
            
                # Bring up the FunctionGroup to choose between adding and removing access #
                $FunctionGroup.Controls.AddRange(@($FunctionRadioEnable,$FunctionRadioRemove,$FunctionSubmitButton))
                $FunctionSubmitButton.Location = '452,20'
                $SingleModeGroup.Controls.Add($SMGOutputBox)
                $MainWindow.Controls.Add($FunctionGroup)
            } Else {
                $SMGLocGroup.Controls.Add($SMGLocButtonDB)
                $SCRIPT:SMGLocation = ""
            }
        }
    )

    # Logic for the DB button in the SMGLocGroup #
    $SMGLocButtonDB.Add_CheckStateChanged(
        {
            If ($SMGLocButtonDB.Checked) {
                # Set the Type variable for the function to "m" #
                $SCRIPT:SMGLocation = "DB"
                Write-Host "Location is set to $SMGLocation, so this is a Data Bank user."
                $SMGLocGroup.Controls.RemoveByKey("SMGLocButtonDC")

                # Bring up the FunctionGroup to choose between adding and removing access #
                $FunctionGroup.Controls.AddRange(@($FunctionRadioEnable,$FunctionRadioRemove,$FunctionSubmitButton))
                $FunctionSubmitButton.Location = '452,20'
                $SingleModeGroup.Controls.Add($SMGOutputBox)
                $MainWindow.Controls.Add($FunctionGroup)
            } Else {
                $SMGLocGroup.Controls.Add($SMGLocButtonDC)
                $SCRIPT:SMGLocation = ""
            }
        }
    )

    # Logic to determine what happens when you click the Browse button #
    $BrowseButton.Add_Click(
        {
            # Run the OpenFileDialog first #
            $BrowseDialog.ShowDialog()

            # Store the path to the selected CSV file in a variable, called $CSVPathName #
            $SCRIPT:CSVPathName = $BrowseDialog.FileName

            # Remove the BrowseButton control after clicking it #
            $MainWindow.Controls.RemoveByKey("BrowseButton")

            # Change the Results label to ask if the file is the right one #
            $CSVBrowseLabel.Text = "Please confirm that this is the correct file!"
            $CSVResultPathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $CSVResultPathLabel.Text = $CSVPathName

            # Replace the controls in the $CSVGroup to show the Confirm and Cancel buttons #
            $CSVGroup.Controls.Clear()
            $CSVGroup.Controls.AddRange(@($CSVBrowseLabel,$CSVResultPathLabel,$CSVConfirmButton,$CSVCancelButton))
        }
    )

    # Logic for when the user confirms their choice of CSV file by clicking the Confirm button #
    $CSVConfirmButton.Add_Click(
        {
            $CSVPath = $CSVPathName
            #Write-Host "Two variables to check - CSVPath: $CSVPath, and CSVPathName: $CSVPathName"

            # Reset the CSVGroup and add $CSVOutputBox to display imported CSV data #
            $CSVGroup.Controls.Clear()
            $CSVGroup.Controls.AddRange(@($CSVBrowseLabel,$CSVResultPathLabel))
            $MainWindow.Controls.Add($CSVOutputBox)
        
            # Change $CSVBrowseLabel text, reset $CSVResultPathLabel, import the chosen CSV and show the output in $CSVOutputBox       
            $CSVBrowseLabel.Text = "Here are the results:"
            $CSVResultPathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
            $SCRIPT:CSVData = Import-Csv -Path $CSVPath 
            $CSVOutput = $CSVData | Select-Object -Property FullName | Out-String
            $CSVOutputBox.Text = $CSVOutput

            # Add the $FunctionGroup so the user can choose to Enable or Remove access; this choice is stored in a variable and that determines which Function will be called #
            # Add radio buttons and Submit button to the FunctionGroup #
            $FunctionGroup.Controls.AddRange(@($FunctionRadioEnable,$FunctionRadioRemove,$FunctionSubmitButton))
            $MainWindow.Controls.Add($FunctionGroup)
        }
    )

    # Click logic for the Cancel button - basically resets the CSVGroup and thus, the form #
    $CSVCancelButton.Add_Click(
        {
            $CSVGroup.Controls.Clear()
            $CSVResultPathLabel.Text = "Please ensure that your CSV file has the proper fields."
            $CSVResultPathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
            $CSVGroup.Controls.AddRange(@($CSVBrowseLabel,$CSVResultPathLabel,$BrowseButton))
        }
    )

    # Adding logic so that the script will take action when the Next button gets clicked #
    $NextButton.Add_Click({
        If ($MainWindow.Controls.ContainsKey("MethodGroup")) {
            If (($MethodRadio1.Checked) -and (!($MethodRadio2.Checked))) {
                # Add a variable to show that it's in Individual mode #
                $SCRIPT:ScriptMode = "Individual"
            } Else {
                $SCRIPT:ScriptMode = "ImportCSV"
            }
            Write-Host "Proceeding in $ScriptMode mode."
            $MainWindow.Controls.RemoveByKey("MethodGroup")
            If ($ScriptMode -eq "ImportCSV") {
                # Line that sets the initial controls for the CSV GroupBox #
                $CSVGroup.Controls.AddRange(@($CSVBrowseLabel,$CSVResultPathLabel,$BrowseButton))
                $MainWindow.Controls.Add($CSVGroup)
                $MainWindow.Controls.RemoveByKey("NextButton")
            } Else {
                # Add label and buttons to the SMGDirectionGroup #
                $SMGDirectionGroup.Controls.AddRange(@($SMGDirectionInstLabel,$SMGDirectionButtonHome,$SMGDirectionButtonWork))
            
                # Add SMGDirectionGroup to the SingleModeGroup #
                $SingleModeGroup.Controls.AddRange(@($SMGDirectionGroup))

                # Add SingleModeGroup to the Form #
                $MainWindow.Controls.Add($SingleModeGroup)
                $MainWindow.Controls.RemoveByKey("NextButton")
            }
        }
    })

    # Click logic for the Function Submit Button - this should fire off the functions, so it will need to process the CSV data if we're in that mode or the data from the user entry if in individual mode #
    $FunctionSubmitButton.Add_Click(
        {
            If ($ScriptMode -eq "ImportCSV") {
                $CSVOutputBox.Text = ""
                ForEach ($CSVEntry in $CSVData) {
                    $User = $CSVEntry.Username
                    $Machine = $CSVEntry.Desktop
                    $SAMName = $Machine + "$"
                    $Location = $CSVEntry.Location
                    $NLA = $CSVEntry.Type
                    $FullName = $CSVEntry.FullName

                    If (($FunctionRadioEnable.Checked) -and (!($FunctionRadioRemove.Checked))) {
                        If ($NLA -eq "w") { $Type = "Windows" } ElseIf ($NLA -eq "m") { $Type = "Mac" }
                        $CSVOutputBox.AppendText("`r`n" + "Enabling RDP for $FullName on $Machine. Username is $User and will be set up to access $Location from $Type.")
                        RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                        If ($FailedHosts -contains $Machine) {
                            $CSVOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                        } Else {
                            $CSVOutputBox.Text += "`r`n" + "$FullName has been enabled for remote access."
                        }
                    }
                    Else {
                        $CSVOutputBox.AppendText("`r`n" + "Removing access for $FullName.")
                        RDPRemove -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName
                        If ($FailedHosts -contains $Machine) {
                            $CSVOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                        } Else {
                            $CSVOutputBox.Text += "`r`n" + "The access for $FullName has been successfully removed."
                        }

                    }

                    $CSVOutputBox.Text += "`r`n" + "Processing is complete. Thanks for using the RDP Management SuperScript!"
                    $MainWindow.Controls.RemoveByKey("FunctionGroup")
                    # Add the buttons to the ResultGroup #
                    $ResultGroup.Controls.AddRange(@($ResultButtonRestart,$ResultButtonQuit))
                    $MainWindow.Controls.Add($ResultGroup)
                }
            # Need to add in code to show that the process is finished and bring the user back either to the start or to the last meaningful choice #

            } ElseIf ($ScriptMode -eq "Individual") {
                $User = $SMGUserName
                $Machine = $SMGDesktopName
                $SAMName = $SMGDesktopSamAccountName
                $Location = $SMGLocation
                $NLA = $NLASwitch
                $FullName = $SMGUserFullName

                If (($FunctionRadioEnable.Checked) -and (!($FunctionRadioRemove.Checked))) {
                    If ($NLA -eq "w") { $Type = "Windows" } ElseIf ($NLA -eq "m") { $Type = "Mac" }
                    If (($Type -eq "Windows") -or ($Type -eq "Mac")) {
                        $SMGOutputBox.Text += "`r`n" + "Enabling RDP for $FullName on $Machine. Username is $User and will be set up to access $Location from $Type."
                        RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                        If ($FailedHosts -contains $Machine) { 
                            $SMGOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                        } Else {
                            $SMGOutputBox.Text += "`r`n" + "$FullName has been enabled for remote desktop access."
                        }
                    } Else {
                        $SMGOutputBox.Text += "`r`n" + "Enabling RDP for $FullName on $Machine. Username is $User and will be set up with internal-only remote desktop access."
                        RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                        If ($FailedHosts -contains $Machine) { 
                            $SMGOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                        } Else {
                            $SMGOutputBox.Text += "`r`n" + "$FullName has been enabled for internal-only remote desktop access."
                        }
                    }
                }
                Else {
                    $SMGOutputBox.Text += "`r`n" + "Removing access for $FullName."
                    RDPRemove -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName
                    If ($FailedHosts -contains $Machine) {
                        $SMGOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                    } Else {    
                        $SMGOutputBox.Text += "`r`n" + "The access for $FullName has been successfully removed."
                    }
                }
                $SMGOutputBox.Text += "`r`n" + "Processing is complete. Thanks for using the RDP Management SuperScript!"
                $MainWindow.Controls.RemoveByKey("FunctionGroup")
                # Add the buttons to the ResultGroup #
                $ResultGroup.Controls.AddRange(@($ResultButtonRestart,$ResultButtonQuit))
                $MainWindow.Controls.Add($ResultGroup)
            }
        }
    )

    $ResultButtonRestart.Add_Click(
        {
            $RestartCounter++
            $MainWindow.Controls.Clear()
            $MainWindow.Controls.Add($MethodGroup)
            $MainWindow.Controls.Add($NextButton)
            If ($ScriptMode -eq "ImportCSV") {
                $CSVBrowseLabel.Text = $DTCSVBrowseLabel
                $CSVResultPathLabel.Text = $DTCSVResultPathLabel
            } Else {
                
                $SingleModeGroup.Controls.Clear()
                $SMGDirectionButtonHome.Checked = $False
                $SMGDirectionButtonWork.Checked = $False
                $SMGUserGroup.Size = '550,100'
                $SMGUserTextBox.Text = ""
                $SMGUserTextBox.ReadOnly = $False
                $SMGUserTextBox.Location = '185,17'
                $SMGUserResultLabel.Text = ""
                $SMGUserResultLabel.Location = '10,45'
                $SMGDesktopGroup.Size = '550,150'
                $SMGDesktopTextBox.Text = ""
                $SMGDesktopTextBox.ReadOnly = $False
                $SMGDesktopTextBox.Location = '225,17'
                $SMGDesktopResultLabel.Text = ""
                $SMGDesktopResultLabel.Location = '10,45'
                $SMGTypeButtonWin.Checked = $False
                $SMGTypeButtonMac.Checked = $False
                $SMGLocButtonDB.Checked = $False
                $SMGLocButtonDC.Checked = $False
                $FunctionRadioEnable.Checked = $False
                $FunctionRadioRemove.Checked = $False
                $FailedHosts = @()
                <#
                CreateForm
                #>
            }
        }
    )
    If ($RestartCounter -lt 1) {
        $MainWindow.ShowDialog()
    }
}


CreateForm
