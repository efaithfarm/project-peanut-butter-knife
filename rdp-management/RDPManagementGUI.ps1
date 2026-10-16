# App to handle most remote management tasks at ****** #

# This line adds the Forms objects #
Add-Type -AssemblyName System.Windows.Forms
$RestartCounter = 0
$FailedHosts = @()
$MMGCSVSearchResultList = [System.Collections.ArrayList]@()
$MMGCSVPSCOList = [System.Collections.ArrayList]@()
$UsernameList = [System.Collections.ArrayList]@()
$PSCOList = [System.Collections.ArrayList]@()
$DifferenceList = [System.Collections.ArrayList]@()
$ChangesList = [System.Collections.ArrayList]@()
$UserChoiceList = [System.Collections.ArrayList]@()
$ChoiceNameList = [System.Collections.ArrayList]@()
$LapseList = [System.Collections.ArrayList]@()

# Specify the path of the excel file #
$FilePath = "C:\Users\*********\Perpetual_Remote_Access.xlsx"

# Specify the path for the log file #
$LogFilePath = "C:\Users\*********\Perpetual_Remote_Access.log"
$LogFileArchive = "C:\Users\*********\LogArchive\"

# Function to write log entries #
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )
    # Create a variable for the timestamp #
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        # Check the logfile to see if it has gotten too big; if so, rename and move the logfile to the local Archive folder first #
        # This should create a new file once the move operation is done #
        $LogFileSize = ((Get-Item $logfile).Length/1MB)
        $FileNameDateFormat = (Get-Date).ToString("MMddyyyyHHmm")
        $ArchivePath = $LogFileArchive + "PRA_Log_" + $FileNameDateFormat + ".log"
        If ($LogFileSize -ge 1.0) {
            Move-Item -Path $logfile -Destination $ArchivePath
        }

        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

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
            Write-Log -Message "Firewall enabled on $TargetServer. Disabling now..." -Level INFO -logfile $LogFilePath
            Remove-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName
        }

    } ElseIf (($NLASwitch -eq "M") -or ($NLASwitch -eq "HTML5")){
        $AuthRequired = 0
        $SecurityLayer = 0

        # When setting up a Mac/HTML5 connection, add the desktop to this security group as an alternative to NLA #
        Add-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName

    } Else {

        # If the user gives a typo at the beginning, exit here #
        Write-Log -Message "This is an internal-only user, so no NLASwitching necessary. But NLA should be on anyway, so..." -Level INFO -logfile $LogFilePath
        $AuthRequired = 1
        $SecurityLayer = 2

    }

    # Turn on RDP #
    Try {
        (Get-WmiObject -class "Win32_TerminalServiceSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -ErrorAction Stop).SetAllowTSConnections(1,1)
    } Catch {
        Write-Log -Message "There was an error attempting to reach the named machine. Might want to check that it's online." -Level WARN -logfile $LogFilePath
        $FailedHosts += $TargetServer
    }
    
    

    # Enable or disable NLA based on user input #
    If ($FailedHosts -notcontains $TargetServer) {
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetUserAuthenticationRequired($AuthRequired)
        } Catch {
            Write-Log -Message "Oh dear. That didn't work either. Machine probably needs to be pinged." -Level WARN -logfile $LogFilePath
        }
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetSecurityLayer($SecurityLayer)
        } Catch {
            Write-Log -Message "Yep, it's not responding. Definitely ping it, and write to the user." -Level WARN -logfile $LogFilePath
        }
    }

    # Add user to local Remote Desktop Users group #
    Try {
        $TargetGroup = [ADSI]"WinNT://$TargetServer/Remote Desktop Users,group"
        $UserPath = [ADSI]"WinNT://******.com/$TargetUser,user"
        If ($TargetGroup -contains $TargetUser) {
            Write-Log -Message "No addition necessary - $TargetUser is already in the local group." -Level INFO -logfile $LogFilePath
        } Else {
            $TargetGroup.Add($UserPath.Path)
        }
    } Catch {
        Write-Log -Message "The local groups couldn't be edited. The machine could be offline, or something." -Level WARN -logfile $LogFilePath
    }
    

    # Add user to SSO Group for specified Pulse location #
    If (($TargetLocation -eq "DC") -or ($TargetLocation -eq "SSO Remote Access Associate")) {
        Add-ADGroupMember -Identity "SSO Remote Access Associate" -Members $TargetUser
    } ElseIf (($TargetLocation -eq "DB") -or ($TargetLocation -eq "SSO Remote Access Data Bank Associate")) {
        Add-ADGroupMember -Identity "SSO Remote Access Data Bank Associate" -Members $TargetUser
    } Else {
        Write-Log -Message "This user is an internal-only user, so this is another unnecessary parameter. Cheers!" -Level INFO -logfile $LogFilePath
    }

    Write-Log -Message "Welcome to the world of Remote Access, $TargetUser!" -Level INFO -logfile $LogFilePath
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
        Write-Log -Message "User $TargetUser found in SSO Data Center group. Removing..." -Level INFO -logfile $LogFilePath
        Remove-ADGroupMember -Identity "SSO Remote Access Associate" -Members $TargetUser -Confirm:$False
    } ElseIf ($DBSSOGroup -contains $TargetUser) {
        Write-Log -Message "User $TargetUser found in SSO Data Bank group. Removing..." -Level INFO -logfile $LogFilePath
        Remove-ADGroupMember -Identity "SSO Remote Access Data Bank Associate" -Members $TargetUser -Confirm:$False
    } Else {
        Write-Log -Message "User not found in either group! Whaddaya know?" -Level INFO -logfile $LogFilePath
    }

    # Alter variables based on user input #

    $AuthRequired = 1
    $SecurityLayer = 2

    # Insert code here to check if the computername is in the security group and remove it if so #

    $ComputerNamesFEC = Get-ADGroupMember -Identity "Firewall Enabled Computers" | Select-Object -ExpandProperty SamAccountName
    If ( $ComputerNamesFEC -contains $TargetSAMAccountName ) {
        Write-Log -Message "Firewall enabled on $TargetServer. Disabling now..." -Level INFO -logfile $LogFilePath
        Remove-ADGroupMember -Identity "Firewall Enabled Computers" -Members $TargetSAMAccountName -Confirm:$False
    }

    # Turn off RDP #
    Try {
        (Get-WmiObject -class "Win32_TerminalServiceSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -ErrorAction Stop).SetAllowTSConnections(0,0)
    } Catch {
        Write-Log -Message "There was an error attempting to reach the named machine. Might want to check that it's online." -Level WARN -logfile $LogFilePath
        $FailedHosts += $TargetServer
    }
    
    

    # Enable or disable NLA based on user input #
    If ($FailedHosts -notcontains $TargetServer) {
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetUserAuthenticationRequired($AuthRequired)
        } Catch {
            Write-Log -Message "Oh dear. That didn't work either. Machine probably needs to be pinged." -Level WARN -logfile $LogFilePath
        }
        Try {
            (Get-WmiObject -class "Win32_TSGeneralSetting" -Namespace root\cimv2\terminalservices -ComputerName $TargetServer -Filter "TerminalName='RDP-tcp'" -ErrorAction Stop).SetSecurityLayer($SecurityLayer)
        } Catch {
            Write-Log -Message "Yep, it's not responding. Definitely ping it, and write to the user." -Level WARN -logfile $LogFilePath
        }
    }
    # Remove user from local Remote Desktop Users group #
    Try {
        $TargetGroup = [ADSI]"WinNT://$TargetServer/Remote Desktop Users,group"
        $UserPath = [ADSI]"WinNT://******.com/$TargetUser,user"
        If ($TargetGroup -contains $TargetUser) {
            Write-Log -Message "No addition necessary - $TargetUser is already in the local group." -Level INFO -logfile $LogFilePath
        } Else {
            $TargetGroup.Remove($UserPath.Path)
        }
    } Catch {
        Write-Log -Message "The local groups couldn't be edited. Most likely because the machine is offline." -Level WARN -logfile $LogFilePath
    }

    Write-Log -Message "Better luck next time, $TargetUser!" -Level INFO -logfile $LogFilePath
}

# Function to find user information in the accompanying Excel spreadsheet and store it in a PSCO #
Function Search-Excel {
    [cmdletbinding()]
    Param (
        [parameter(Mandatory)]
        [ValidateScript({
            Try {
                If (Test-Path -Path $_) {$True}
                Else {Throw "$($_) is not a valid path."}
            }
            Catch {
                Throw $_
            }
        })]
        [string]$Source,
        [parameter(Mandatory)]
        [ValidateSet("A","I","L")]
        [string]$Mode,
        [string]$SearchText
    )

    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    # Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath)

    $WorkSheet = $WorkBook.Sheets.Item("Last 30")
    If ($Mode -eq "I") {
        $Found = $WorkSheet.Cells.Find($SearchText)

        If ($Found) {
            $BeginAddress = $Found.Address(0,0,1,1)
            $TableFound = "YEP"

            [int]$LastRow = ($WorkSheet.UsedRange.rows.count)
            [int]$NextRow = $LastRow + 1
            [int]$LastEntryNumber = $WorkSheet.Rows.Item($LastRow).Columns.Item(1).Text

            $Row = $Found.Row
            $Column = $Found.Column
            $TableEntryNumber = $WorkSheet.Rows.Item($Row).Columns.Item(1).Text
            $TableDisplayName = $WorkSheet.Rows.Item($Row).Columns.Item(2).Text
            $TableUserName = $WorkSheet.Rows.Item($Row).Columns.Item(3).Text
            $TableAssociateID = $WorkSheet.Rows.Item($Row).Columns.Item(4).Text
            $TableIDLookup = $WorkSheet.Rows.Item($Row).Columns.Item(5).Text
            $TableEarningsGroup = $WorkSheet.Rows.Item($Row).Columns.Item(6).Text
            $TableDealership = $WorkSheet.Rows.Item($Row).Columns.Item(7).Text
            $TableDepartment = $WorkSheet.Rows.Item($Row).Columns.Item(8).Text
            $TableTitle = $WorkSheet.Rows.Item($Row).Columns.Item(9).Text
            $TableLastLoginCurrent = $WorkSheet.Rows.Item($Row).Columns.Item(10).Text
            $TableLastLoginFormula = $WorkSheet.Rows.Item($Row).Columns.Item(11).Formula
            $TableLastLoginHistorical = $WorkSheet.Rows.Item($Row).Columns.Item(12).Text
            $TableLast30Formula = $WorkSheet.Rows.Item($Row).Columns.Item(13).Formula
            $TableLast30Current = $WorkSheet.Rows.Item($Row).Columns.Item(14).Text
            $TableRDPToggle = $WorkSheet.Rows.Item($Row).Columns.Item(15).Text
            $TableVPNToggle = $WorkSheet.Rows.Item($Row).Columns.Item(16).Text
            $TableGroups = $WorkSheet.Rows.Item($Row).Columns.Item(17).Text
            $TableAccessType = $WorkSheet.Rows.Item($Row).Columns.Item(18).Text
            $TableDesktopName = $WorkSheet.Rows.Item($Row).Columns.Item(19).Text
            $TableDateAdded = $WorkSheet.Rows.Item($Row).Columns.Item(20).Text
            $TableDateLastUpdated = $WorkSheet.Rows.Item($Row).Columns.Item(21).Text
            $TableStatus = $WorkSheet.Rows.Item($Row).Columns.Item(22).Text
            $TableNotes = $WorkSheet.Rows.Item($Row).Columns.Item(23).Text
            $NextEntryNumber = $LastEntryNumber + 1
            $TableEnabled = ""
            $TableEmailAddress = ""
            $TableInfo = ""

            Write-Log -Level INFO -logfile $LogFilePath -Message "User $TableUserName found on row $Row. Proceeding with processing."

            # Displays it the first time it is found
            [pscustomobject]@{
                #WorkSheet = $WorkSheet.Name
                1 = $TableEntryNumber
                2 = $TableDisplayName
                3 = $TableUserName
                4 = $TableAssociateID
                5 = $TableIDLookup
                6 = $TableEarningsGroup
                7 = $TableDealership
                8 = $TableDepartment
                9 = $TableTitle
                10 = $TableLastLoginCurrent
                11 = $TableLastLoginFormula
                12 = $TableLastLoginHistorical
                13 = $TableLast30Formula
                14 = $TableLast30Current
                15 = $TableRDPToggle
                16 = $TableVPNToggle
                17 = $TableGroups
                18 = $TableAccessType
                19 = $TableDesktopName
                20 = $TableDateAdded
                21 = $TableDateLastUpdated
                22 = $TableStatus
                23 = $TableNotes
                24 = $Row # In the event you want to display Row
                25 = $Column # In the event you wanted to display column
                26 = $NextRow
                27 = $NextEntryNumber
                28 = $TableFound
                29 = $TableEnabled
                30 = $TableEmailAddress
                31 = $TableInfo
            }

        }
        Else {

            Write-Log -Message "No matches found in the $($Worksheet.Name) tab for $SearchText." -Level INFO -logfile $LogFilePath
            $TableFound = "NOPE"
            [int]$LastRow = ($WorkSheet.UsedRange.rows.count)
            [int]$NextRow = $LastRow + 1
            [int]$LastEntryNumber = $WorkSheet.Rows.Item($LastRow).Columns.Item(1).Text
            [int]$NextEntryNumber = $LastEntryNumber + 1

            [pscustomobject]@{
                26 = $NextRow
                27 = $NextEntryNumber
                28 = $TableFound
            }

        }

    } ElseIf ($Mode -eq "A") {

        $LastRow = ($WorkSheet.UsedRange.rows.count)
        $NextRow = $LastRow + 1
        [int]$Row = 4

        Do
            {
                $UsernameList.Add($WorkSheet.Cells.Item($Row,3).Text) | Out-Null
                $Row ++
            }
        Until ($Row -eq ($NextRow))

        [int]$LastEntryNumber = $WorkSheet.Rows.Item($LastRow).Columns.Item(1).Text
        $NextEntryNumber = $LastEntryNumber + 1


        ForEach ($Username in $UsernameList) {
            $Found = $WorkSheet.Cells.Find($Username)
            $BeginAddress = $Found.Address(0,0,1,1)

            $Row = $Found.Row
            $TableDisplayName = $WorkSheet.Rows.Item($Row).Columns.Item(2).Text
            $TableUserName = $WorkSheet.Rows.Item($Row).Columns.Item(3).Text
            $TableAssociateID = $WorkSheet.Cells.Item($Row,4).Text
            Write-Host "Associate ID for $TableDisplayName is recorded as $TableAssociateID."
            $TableDealership = $WorkSheet.Rows.Item($Row).Columns.Item(7).Text
            $TableDepartment = $WorkSheet.Rows.Item($Row).Columns.Item(8).Text
            $TableTitle = $WorkSheet.Rows.Item($Row).Columns.Item(9).Text
            $TableNotes = $WorkSheet.Rows.Item($Row).Columns.Item(23).Text

            $PSCO = [pscustomobject]@{
                2 = $TableDisplayName
                3 = $TableUserName
                4 = $TableAssociateID
                7 = $TableDealership
                8 = $TableDepartment
                9 = $TableTitle
                23 = $TableNotes
                24 = $Row
            }
            $PSCOList.Add($PSCO) | Out-Null
        }

    } Else {

        $LastRow = ($WorkSheet.UsedRange.rows.count)
        $NextRow = $LastRow + 1
        [int]$Row = 4

        Do
            {
                $UsernameList.Add($WorkSheet.Cells.Item($Row,3).Text) | Out-Null
                $Row ++
            }
        Until ($Row -eq ($NextRow))

        [int]$LastEntryNumber = $WorkSheet.Rows.Item($LastRow).Columns.Item(1).Text
        $NextEntryNumber = $LastEntryNumber + 1


        ForEach ($Username in $UsernameList) {
            $Found = $WorkSheet.Cells.Find($Username)
            $BeginAddress = $Found.Address(0,0,1,1)

            $Row = $Found.Row
            $TableDisplayName = $WorkSheet.Rows.Item($Row).Columns.Item(2).Text
            $TableUserName = $WorkSheet.Rows.Item($Row).Columns.Item(3).Text
            $TableAssociateID = $WorkSheet.Cells.Item($Row,4).Text
            Write-Host "Associate ID for $TableDisplayName is recorded as $TableAssociateID."
            $TableDealership = $WorkSheet.Rows.Item($Row).Columns.Item(7).Text
            $TableDepartment = $WorkSheet.Rows.Item($Row).Columns.Item(8).Text
            $TableTitle = $WorkSheet.Rows.Item($Row).Columns.Item(9).Text
            $TableLastLoginCurrent = $WorkSheet.Rows.Item($Row).Columns.Item(10).Text
            $TableVPNToggle = $WorkSheet.Rows.Item($Row).Columns.Item(16).Text
            $TableGroups = $WorkSheet.Rows.Item($Row).Columns.Item(17).Text
            $TableAccessType = $WorkSheet.Rows.Item($Row).Columns.Item(18).Text
            $TableDesktopName = $WorkSheet.Rows.Item($Row).Columns.Item(19).Text
            $TableStatus = $WorkSheet.Rows.Item($Row).Columns.Item(22).Text
            $TableNotes = $WorkSheet.Rows.Item($Row).Columns.Item(23).Text

            $PSCO = [pscustomobject]@{
                2 = $TableDisplayName
                3 = $TableUserName
                4 = $TableAssociateID
                7 = $TableDealership
                8 = $TableDepartment
                9 = $TableTitle
                10 = $TableLastLoginCurrent
                16 = $TableVPNToggle
                17 = $TableGroups
                18 = $TableAccessType
                19 = $TableDesktopName
                22 = $TableStatus
                23 = $TableNotes
                24 = $Row
            }
            $PSCOList.Add($PSCO) | Out-Null
        }

    }
    # Get rid of the Excel object #

    Write-Log -Message "..." -Level INFO -logfile $LogFilePath
    $objExcel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    Remove-Variable excel -ErrorAction SilentlyContinue
}

# Function to make changes to the spreadsheet based on a PSCO created elsewhere in the script #
Function Write-Excel {
    [cmdletbinding()]
    Param (
        
        # The $Source is required, and should be checked to make sure the file is there #
        [parameter(Mandatory)]
        [ValidateScript({
        Try{
            If (Test-Path -Path $_) {$True}
            Else {Throw "$($_) is not a valid path."}
            }
        Catch {
            Throw $_
        }
        })]
        [string]$Source,

        # This function is dependent on the output of the Search-Excel function; it creates the PSCO and this one uses it #
        [parameter(Mandatory)]
        [PSCustomObject]$SearchExcelOutput
    )

    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    # Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath,0,$false,5)

    $WorkSheet = $WorkBook.Sheets.Item("Last 30")

    [int]$TargetRow = $SearchExcelOutput.24
    Write-Log -Message "The Target Row is $TargetRow" -Level INFO -logfile $LogFilePath
    $SearchExcelOutput.PSObject.Properties | ForEach-Object {
        [int]$OutputItemColumnNumber = $_.Name
        Write-Log -Message "In column number $OutputItemColumnNumber..." -Level INFO -logfile $LogFilePath
        $OutputItemValue = $_.Value
        Write-Log -Message "We find the value of...$OutputItemValue" -Level INFO -logfile $LogFilePath
        If ($OutputItemColumnNumber -lt 24) {
            If (($OutputItemColumnNumber -eq "22") -and ($OutputItemValue -eq "I")) {
                Write-Log -Message "Changing the style at ($TargetRow,$OutputItemColumnNumber). Status is now $OutputItemValue." -Level INFO -logfile $LogFilePath
                $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Style = 'Bad'
                $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber) = $OutputItemValue
            } ElseIf (($OutputItemColumnNumber -eq "22") -and ($OutputItemValue -eq "A")) {
                Write-Log -Message "Changing the style at ($TargetRow,$OutputItemColumnNumber). Status is now $OutputItemValue." -Level INFO -logfile $LogFilePath
                $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Style = 'Good'
                $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber) = $OutputItemValue
            } ElseIf ($OutputItemColumnNumber -eq "20") {
                Write-Log -Message "Changing formatting for DateAdded column." -Level INFO -logfile $LogFilePath
                $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).HorizontalAlignment = -4152
                $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber) = $OutputItemValue
            } Else {
                $CurrentValue = $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Text
                If ((($OutputItemValue -eq "") -and ($CurrentValue -ne "")) -or ($OutputItemValue -eq $CurrentValue)) {
                    Write-Log -Message "No overwriting necessary, we've already got a valid value here." -Level INFO -logfile $LogFilePath
                } Else {
                    Write-Log -Message "Entering $OutputItemValue at ($TargetRow,$OutputItemColumnNumber) to replace $CurrentValue." -Level INFO -logfile $LogFilePath
                    $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber) = $OutputItemValue
                }
            }
        } Else {
            Write-Log -Message "Skipping this column - it should not be valued. That kinda sounds mean." -Level INFO -logfile $LogFilePath
        }
    }

    Write-Log -Message "..." -Level INFO -logfile $LogFilePath
    $objExcel.DisplayAlerts = $false
    $WorkBook.SaveAs($FilePath)
    #Read-Host "Ready to quit? Just hit enter."
    $objExcel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    Remove-Variable excel -ErrorAction SilentlyContinue

}

# Function to massage the output of Get-ADUser into a PSCO like the Search-Excel function provides #
Function Fix-GetADUser {
    [cmdletbinding()]
    Param (
        
        # Two mandatory parameters here - first, this function needs to be fed the results of a Get-ADUser command #
        [parameter(Mandatory)]
        [Microsoft.ActiveDirectory.Management.ADUser]$ADInfo,
        # Second, we need to know how many fields to use; when comparing this output with the spreadsheet, we need a lot less #
        [bool]$Compare
    )

    # Step through the properties of the PSCO-to-be and set them either as blank if GADU doesn't provide it or pull the info out #
    $Row = ""
    $Column = ""
    $TableEntryNumber = ""
    $TableDisplayName = $ADInfo.DisplayName
    $TableUserName = $ADInfo.SamAccountName
    $TableAssociateID = $ADInfo.EmployeeID
    $TableIDLookup = ""
    $TableEarningsGroup = ""
    If ($ADInfo.Title -eq "CustomerCareRepresentativ") {
        $TableTitle = "Customer Care Representative"
    } Else {
        $TableTitle = $ADInfo.Title
    }
    #Write-Log -Message "Title set to $TableTitle, via AD" -Level INFO -logfile $LogFilePath

    # Switch statement, because payroll/HR is weird and uses dumb codes for all of the dealerships. This corrects that #
    Switch ($ADInfo.Company) {
        "BMWG" { $TableDealership = "SBMWG"; Break }
        "JLRNA" { $TableDealership = "SJLRNA"; Break }
        "KBCSH" { $TableDealership = "KBC"; Break }
        "MBWH" { $TableDealership = "SMBWH"; Break }
        "SAH" { $TableDealership = "SANH"; Break }

        # Exhibit 1A on how weird PHR is; "SCSD" means a lot of things that have nothing direct to do with Cadillac Dallas #
        "SCSD" { Switch ($ADInfo.Department) {
                "SCSD-Accounting" { $TableDealership = "CORP"; Break }
                "SCSD-Corporate" { $TableDealership = "CORP"; Break }
                "SCSD-General & Administrative" { 
                    If (($TableTitle -eq "Customer Care Representative") -or ($TableTitle -eq "Director") -or ($TableTitle -eq "Manager")) {
                        $TableDealership = "CORP"
                        $TableDepartment = "CXC"
                        $DepartmentSet = "YEP"
                        
                    } Else {
                        $TableDealership = "SCD"
                        $TableDepartment = "General/Admin"
                        $DepartmentSet = "YEP"
                    }; Break
                }
                "SCSD-HR/Payroll" { $TableDealership = "CORP"; Break }
                "SCSD-Information Services" { $TableDealership = "CORP"; Break }
                "SCSD-Marketing" { $TableDealership = "CORP"; Break }
                Default { $TableDealership = "SCD"; Break }
            }
        }
        "SCSG" { $TableDealership = "SCG"; Break }
        "SHD" { $TableDealership = "SSD"; Break }
        "SBPG" { $TableDealership = "SBGD"; Break }
        "SI" { $TableDealership = "SID"; Break }
        Default { $TableDealership = $ADInfo.Company }
    }
    #Write-Log -Message "Dealership set to $TableDealership, via AD" -Level INFO -logfile $LogFilePath
    
    # Pull apart the "Department" attribute and ditch the dumb company code; you then (for the most part) get the actual Department for each user #
    If ($DepartmentSet -ne "YEP") {
        $SplitDepartment = $ADInfo.Department -split '-'
        Switch ($SplitDepartment[1]) {
            "Finance" { $TableDepartment = "F&I"; Break }
            "Information Services" { $TableDepartment = "IS"; Break }
            "Corporate" { $TableDepartment = "General/Admin"; Break }
            "General & Administrative" { $TableDepartment = "General/Admin"; Break }
            Default { $TableDepartment = $SplitDepartment[1] }
        }
    }
    #Write-Log -Message "Department set to $TableDepartment, via AD" -Level INFO -logfile $LogFilePath

    # None of these are found in AD; they need to be set later in this script #
    $TableLastLoginCurrent = ""
    $TableLastLoginFormula = ""
    $TableLastLoginHistorical = ""
    $TableLast30Formula = ""
    $TableLast30Current = ""
    $TableRDPToggle = ""
    $TableVPNToggle = ""
    $TableGroups = ""
    $TableAccessType = ""
    $TableDesktopName = ""
    $TableDateAdded = ""
    $TableDateLastUpdated = ""
    $TableStatus = ""
    $TableNotes = ""
    $NextEntryNumber = ""
    $TableFound = ""
    $TableEnabled = $ADInfo.Enabled
    $TableEmailAddress = $ADInfo.EmailAddress
    $TableInfo = $ADInfo.Info
    If (!($Compare)) {
        # Create the PSCO here so that the results of this can be passed around and set up by the script #
        [pscustomobject]@{
            1 = $TableEntryNumber
            2 = $TableDisplayName
            3 = $TableUserName
            4 = $TableAssociateID
            5 = $TableIDLookup
            6 = $TableEarningsGroup
            7 = $TableDealership
            8 = $TableDepartment
            9 = $TableTitle
            10 = $TableLastLoginCurrent
            11 = $TableLastLoginFormula
            12 = $TableLastLoginHistorical
            13 = $TableLast30Formula
            14 = $TableLast30Current
            15 = $TableRDPToggle
            16 = $TableVPNToggle
            17 = $TableGroups
            18 = $TableAccessType
            19 = $TableDesktopName
            20 = $TableDateAdded
            21 = $TableDateLastUpdated
            22 = $TableStatus
            23 = $TableNotes
            24 = $Row
            25 = $Column
            26 = $NextRow
            27 = $NextEntryNumber
            28 = $TableFound
            29 = $TableEnabled
            30 = $TableEmailAddress
            31 = $TableInfo
        }
    } Else {
        # Create the PSCO here but only with select properties to focus on what could actually be different #
        [pscustomobject]@{
            2 = $TableDisplayName
            3 = $TableUserName
            4 = $TableAssociateID
            7 = $TableDealership
            8 = $TableDepartment
            9 = $TableTitle
            29 = $TableEnabled
        }
    }
}

# Function to compare two PSCOs, in this case; the code for this was written by someone else, I don't remember who #
Function Compare-ObjectProperties {
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject 
    )
    $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops = $objprops | Sort | Select -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff) {            
            $diffprops = @{
                PropertyName=$objprop
                RefValue=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                DiffValue=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }        
    }
    if ($diffs) {return ($diffs | Select PropertyName,RefValue,DiffValue)}     
}

# WINDOWS FORMS CODE TO CREATE THE GUI #

# Function to set up the form #

Function CreateForm {

    # Establish the main form window and define its dimensions #
    $MainWindow = New-Object System.Windows.Forms.Form
    $MainWindow.Text = "****** Remote Access Management App"
    $MainWindow.Width = 640
    $MainWindow.Height = 480
    $MainWindow.AutoSize = $True
    $MainWindow.Padding = 10

    # STEP ONE #
    # Start by determining whether this is bulk processing or individual processing, by means of two radio buttons and a Next button #

    $MethodGroup = New-Object System.Windows.Forms.GroupBox
    $MethodGroup.Location = '10,10'
    $MethodGroup.Text = "Please select a processing mode:"
    $MethodGroup.Size = '440,180'

    # Give the group a name so that it can be acted upon and removed when necessary #
    $MethodGroup.Name = "MethodGroup"

    # Radio button indicating Individual mode #
    $MethodRadio1 = New-Object System.Windows.Forms.RadioButton
    $MethodRadio1.Location = '20,20'
    $MethodRadio1.Text = "Individual"
    $MethodRadio1.Checked = $True

    # Radio button indicating MMGCSV Import mode #
    $MethodRadio2 = New-Object System.Windows.Forms.RadioButton
    $MethodRadio2.Location = '180,20'
    $MethodRadio2.Text = "Import CSV"
    $MethodRadio2.Checked = $False

    # Radio button indicating Update/Find Lapsed Users mode #
    $MethodRadio3 = New-Object System.Windows.Forms.RadioButton
    $MethodRadio3.Location = '20,80'
    $MethodRadio3.Text = "Update/Find Lapsed Users"
    $MethodRadio3.Checked = $False
    $MethodRadio3.Size = '120,30'

    # Radio button indicating Check Status mode #
    $MethodRadio4 = New-Object System.Windows.Forms.RadioButton
    $MethodRadio4.Location = '180,80'
    $MethodRadio4.Text = "Check User Status"
    $MethodRadio4.Checked = $False
    $MethodRadio4.Size = '120,30'

    # Add the radio buttons to the GroupBox #
    $MethodGroup.Controls.AddRange(@($MethodRadio1,$MethodRadio2,$MethodRadio3,$MethodRadio4))

    # The Next button moves the GUI to the next step in the process #
    $NextButton = New-Object System.Windows.Forms.Button
    $NextButton.Location = '460,20'
    $NextButton.Size = '120,23'
    $NextButton.Text = "Next"
    $NextButton.Name = "NextButton"

    # Add the initial $MethodGroup GroupBox and $NextButton to the Form #
    $MainWindow.Controls.Add($MethodGroup)
    $MainWindow.Controls.Add($NextButton)

    # Group to hold controls for CSV and UaL modes #
    $MultiModeGroup = New-Object System.Windows.Forms.GroupBox
    $MultiModeGroup.Location = '10,10'
    $MultiModeGroup.AutoSize = $True
    $MultiModeGroup.Name = "MultiModeGroup"
    $MultiModeGroup.Text = ""
    $MultiModeGroup.Padding = 5

    # Instruction label for the CSV Browse dialog #
    $MMGCSVBrowseLabel = New-Object System.Windows.Forms.Label
    $MMGCSVBrowseLabel.Location = '10,20'
    # This is to set the default text of the label as a variable, so it can be reset later #
    $DTCSVBrowseLabel = "Click the button to browse and find a CSV file."
    $MMGCSVBrowseLabel.Text = $DTCSVBrowseLabel
    $MMGCSVBrowseLabel.AutoSize = $True

    # Instruction label that will show the path chosen after the dialog is completed #
    $MMGCSVResultPathLabel = New-Object System.Windows.Forms.Label
    $MMGCSVResultPathLabel.Location = '10,40'
    # This is to set the default text of the label as a variable, so it can be reset later #
    $DTCSVResultPathLabel = "Please ensure that your CSV file has the proper fields."
    $MMGCSVResultPathLabel.Text = $DTCSVResultPathLabel
    $MMGCSVResultPathLabel.AutoSize = $True

    # Button that confirms user choice of CSV file and moves the script to the next step on click #
    $MMGCSVConfirmButton = New-Object System.Windows.Forms.Button
    $MMGCSVConfirmButton.Text = "Confirm"
    $MMGCSVConfirmButton.Location = '450,60'
    $MMGCSVConfirmButton.Size = '100,23'
    $MMGCSVConfirmButton.Name = "MMGCSVConfirmButton"

    # Button that cancels user choice of CSV file and brings the Browse button back #
    $MMGCSVCancelButton = New-Object System.Windows.Forms.Button
    $MMGCSVCancelButton.Text = "Cancel"
    $MMGCSVCancelButton.Location = '570,60'
    $MMGCSVCancelButton.Size = '100,23'
    $MMGCSVCancelButton.Name = "MMGCSVCancelButton"

    # Read-only, multiline textbox with scrollbars to show output of the Import-CSV process #
    $MMGCSVOutputBox = New-Object System.Windows.Forms.Textbox
    $MMGCSVOutputBox.Text = ""
    $MMGCSVOutputBox.Location = '10,130'
    $MMGCSVOutputBox.Size = '600,200'
    $MMGCSVOutputBox.Multiline = $True
    $MMGCSVOutputBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $MMGCSVOutputBox.ReadOnly = $True

    # Code to bring up the file browser object #
    $MMGCSVBrowseDialog = New-Object System.Windows.Forms.OpenFileDialog
    $MMGCSVBrowseDialog.Filter = "CSV Files (*.csv)| *.csv"
    $MMGCSVBrowseDialog.Title = "Choose CSV File"

    # Button that brings out the file browser object when user clicks on it #
    $MMGCSVBrowseButton = New-Object System.Windows.Forms.Button
    $MMGCSVBrowseButton.Text = "Select CSV"
    $MMGCSVBrowseButton.Location = '460,20'
    $MMGCSVBrowseButton.Size = '120,23'
    $MMGCSVBrowseButton.Name = "MMGCSVBrowseButton"

    # Instruction label for the UaL StatusCheck Button #
    $MMGUaLUpdateLabel = New-Object System.Windows.Forms.Label
    $MMGUaLUpdateLabel.Location = '10,20'
    # This is to set the default text of the label as a variable, so it can be reset later #
    $DTUaLUpdateLabel = "Click here to check Active Directory for updated information."
    $MMGUaLUpdateLabel.Text = $DTUaLUpdateLabel
    $MMGUaLUpdateLabel.AutoSize = $True
    $MMGUaLUpdateLabel.Name = "MMGUaLUpdateLabel"

    # Instruction label for the UaL LapseCheck Button #
    $MMGUaLLapseCheckLabel = New-Object System.Windows.Forms.Label
    $MMGUaLLapseCheckLabel.Location = '10,70'
    # This is to set the default text of the label as a variable, so it can be reset later #
    $DTUaLLapseCheckLabel = "Click here to check for updates and for lapsed users."
    $MMGUaLLapseCheckLabel.Text = $DTUaLLapseCheckLabel
    $MMGUaLLapseCheckLabel.AutoSize = $True
    $MMGUaLLapseCheckLabel.Name = "MMGUaLLapseCheckLabel"

    # Button to proceed to StatusCheck #
    $MMGUaLUpdateButton = New-Object System.Windows.Forms.Button
    $MMGUaLUpdateButton.Text = "Check for Updates"
    $MMGUaLUpdateButton.Location = '420,20'
    $MMGUaLUpdateButton.Size = '160,23'
    $MMGUaLUpdateButton.Name = "MMGUaLUpdateButton"

    # Button to proceed to StatusCheck #
    $MMGUaLLapseCheckButton = New-Object System.Windows.Forms.Button
    $MMGUaLLapseCheckButton.Text = "Check for Lapsed Users"
    $MMGUaLLapseCheckButton.Location = '420,70'
    $MMGUaLLapseCheckButton.Size = '160,23'
    $MMGUaLLapseCheckButton.Name = "MMGUaLLapseCheckButton"

    # Checkbox to select all changes or deselect them #
    $MMGUaLCheckAllCheckbox = New-Object System.Windows.Forms.CheckBox
    $MMGUaLCheckAllCheckbox.Location = '15,60'
    $MMGUaLCheckAllCheckbox.Name = "MMGUaLCheckAllCheckbox"

    # CheckedListBox to select users to update after all checks are complete #
    $MMGUaLCheckedList = New-Object System.Windows.Forms.CheckedListBox
    $MMGUaLCheckedList.Location = '10,70'
    $MMGUaLCheckedList.CheckOnClick = $true
    $MMGUaLCheckedList.Size = '600,300'

    # Instruction label for the UaL Results buttons #
    $MMGUaLResultLabel = New-Object System.Windows.Forms.Label
    $MMGUaLResultLabel.Location = '10,400'
    # This is to set the default text of the label as a variable, so it can be reset later #
    $DTUaLResultLabel = "Would you like to accept the changes, or cancel?"
    $MMGUaLResultLabel.Text = $DTUaLResultLabel
    $MMGUaLResultLabel.AutoSize = $True
    $MMGUaLResultLabel.Name = "MMGUaLResultLabel"

    # Button to submit list of changes #
    $MMGUaLSubmitButton = New-Object System.Windows.Forms.Button
    $MMGUaLSubmitButton.Text = "Submit Changes"
    $MMGUaLSubmitButton.Location = '300,400'
    $MMGUaLSubmitButton.Size = '120,23'
    $MMGUaLSubmitButton.Name = "MMGUaLSubmitButton"

    # Button to cancel proposed changes #
    $MMGUaLCancelButton = New-Object System.Windows.Forms.Button
    $MMGUaLCancelButton.Text = "Cancel Changes"
    $MMGUaLCancelButton.Location = '430,400'
    $MMGUaLCancelButton.Size = '120,23'
    $MMGUaLCancelButton.Name = "MMGUaLCancelButton"

    # Group to hold controls for Individual and StatusCheck modes #
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

    # Group to hold controls for entering Username information #
    $SMGUserGroup = New-Object System.Windows.Forms.GroupBox
    $SMGUserGroup.Location = '15,15'
    $SMGUserGroup.Size = '550,120'
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
    $SMGUserButtonConfirm.Location = '295,90'
    $SMGUserButtonConfirm.Size = '120,23'
    $SMGUserButtonConfirm.Text = "Yes"
    $SMGUserButtonConfirm.Name = "SMGUserButtonConfirm"

    # Button to cancel user selection #
    $SMGUserButtonCancel = New-Object System.Windows.Forms.Button
    $SMGUserButtonCancel.Location = '425,90'
    $SMGUserButtonCancel.Size = '120,23'
    $SMGUserButtonCancel.Text = "No"
    $SMGUserButtonCancel.Name = "SMGUserButtonCancel"

    # Button user clicks when user isn't in the sheet and is being added for the first time #
    $SMGUserButtonNewUser = New-Object System.Windows.Forms.Button
    $SMGUserButtonNewUser.Location = '295,90'
    $SMGUserButtonNewUser.Size = '120,23'
    $SMGUserButtonNewUser.Text = "New User"
    $SMGUserButtonNewUser.Name = "SMGUserButtonNewUser"

    # Group to hold Desktop selection controls #
    $SMGDesktopGroup = New-Object System.Windows.Forms.GroupBox
    $SMGDesktopGroup.Location = '15,60'
    $SMGDesktopGroup.Size = '550,150'
    $SMGDesktopGroup.Name = "SMGDesktopGroup"

    # Label giving instructions for desktop entry #
    $SMGDesktopInstLabel = New-Object System.Windows.Forms.Label
    $SMGDesktopInstLabel.Location = '10,15'
    $SMGDesktopInstLabel.Text = "Enter the associate's office PC name:"
    $SMGDesktopInstLabel.AutoSize = $True
    $SMGDesktopInstLabel.Name = "SMGDesktopInstLabel"

    # TextBox for user to enter associate's desktop computername #
    $SMGDesktopTextBox = New-Object System.Windows.Forms.TextBox
    $SMGDesktopTextBox.Location = '225,17'
    $SMGDesktopTextBox.Size = '100,20'
    $SMGDesktopTextBox.Text = ""
    $SMGDesktopTextBox.Name = "SMGDesktopTextBox"

    # Button to accept desktop name from sheet #
    $SMGDesktopButtonAccept = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonAccept.Location = '295,105'
    $SMGDesktopButtonAccept.Size = '120,23'
    $SMGDesktopButtonAccept.Text = "Yes"
    $SMGDesktopButtonAccept.Name = "SMGDesktopButtonAccept"

    # Button to reject desktop name from sheet and begin search for entered desktop name #
    $SMGDesktopButtonReject = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonReject.Location = '425,105'
    $SMGDesktopButtonReject.Size = '120,23'
    $SMGDesktopButtonReject.Text = "No"
    $SMGDesktopButtonReject.Name = "SMGDesktopButtonReject"

    # Button to abort script in case of ping failure #
    $SMGDesktopButtonAbort = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonAbort.Location = '295,105'
    $SMGDesktopButtonAbort.Size = '120,23'
    $SMGDesktopButtonAbort.Text = "Quit"
    $SMGDesktopButtonAbort.Name = "SMGDesktopButtonAbort"

    # Button to begin search for entered desktop name #
    $SMGDesktopButtonVerify = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonVerify.Location = '335,15'
    $SMGDesktopButtonVerify.Size = '120,23'
    $SMGDesktopButtonVerify.Text = "Find PC"
    $SMGDesktopButtonVerify.Name = "SMGDesktopButtonVerify"

    # Label to display results of Get-ADComputer (which runs as result of the SMGDesktopButtonVerify's click) #
    $SMGDesktopResultLabel = New-Object System.Windows.Forms.Label
    $SMGDesktopResultLabel.Location = '10,75'
    $SMGDesktopResultLabel.AutoSize = $True
    $SMGDesktopResultLabel.Text = ""
    $SMGDesktopResultLabel.MaximumSize = '500,50'
    $SMGDesktopResultLabel.Name = "SMGDesktopResultLabel"

    # Button to confirm Desktop selection #
    $SMGDesktopButtonConfirm = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonConfirm.Location = '295,115'
    $SMGDesktopButtonConfirm.Size = '120,23'
    $SMGDesktopButtonConfirm.Text = "Yes"
    $SMGDesktopButtonConfirm.Name = "SMGDesktopButtonConfirm"

    # Button to cancel Desktop selection #
    $SMGDesktopButtonCancel = New-Object System.Windows.Forms.Button
    $SMGDesktopButtonCancel.Location = '425,115'
    $SMGDesktopButtonCancel.Size = '120,23'
    $SMGDesktopButtonCancel.Text = "No"
    $SMGDesktopButtonCancel.Name = "SMGDesktopButtonCancel"

    # Group to hold controls for choosing Type (Windows or Mac) #
    $SMGTypeGroup = New-Object System.Windows.Forms.GroupBox
    $SMGTypeGroup.Location = '15,105'
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

    # Button to confirm preselected option from sheet #
    $SMGTypeButtonConfirm = New-Object System.Windows.Forms.Button
    $SMGTypeButtonConfirm.Location = '465,15'
    $SMGTypeButtonConfirm.Size = '75,23'
    $SMGTypeButtonConfirm.Text = "Confirm"
    $SMGTypeButtonConfirm.Name = "SMGTypeButtonConfirm"

    # Button to indicate a transfer of settings from one Type to another #
    $SMGTypeButtonSwitch = New-Object System.Windows.Forms.Button
    $SMGTypeButtonSwitch.Location = '375,15'
    $SMGTypeButtonSwitch.Size = '75,23'
    $SMGTypeButtonSwitch.Text = "Switch"
    $SMGTypeButtonSwitch.Name = "SMGTypeButtonSwitch"

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

    # Button for SMGLoc to accept the information pulled from the sheet #
    $SMGLocButtonConfirm = New-Object System.Windows.Forms.Button
    $SMGLocButtonConfirm.Location = '420,45'
    $SMGLocButtonConfirm.Size = '120,23'
    $SMGLocButtonConfirm.Text = "Confirm"
    $SMGLocButtonConfirm.Name = "SMGLocButtonConfirm"
    
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

    <#
    # Label showing the current Status attribute in the sheet and prompting for confirmation - This is SMG-only #
    $FunctionLabel = New-Object System.Windows.Forms.Label
    $FunctionLabel.Location = '10,15'
    $FunctionLabel.Text = ""
    $FunctionLabel.AutoSize = $True
    $FunctionLabel.Name = "FunctionLabel"
    #>

    # Radio button to choose Enable #
    $FunctionRadioEnable = New-Object System.Windows.Forms.RadioButton
    $FunctionRadioEnable.Location = '20,20'
    $FunctionRadioEnable.Checked = $False
    $FunctionRadioEnable.Text = "Enable Access"
    $FunctionRadioEnable.Name = "FunctionRadioEnable"

    # Radio button to choose Remove #
    $FunctionRadioRemove = New-Object System.Windows.Forms.RadioButton
    $FunctionRadioRemove.Location = '180,20'
    $FunctionRadioRemove.Checked = $False
    $FunctionRadioRemove.Text = "Remove Access"
    $FunctionRadioRemove.Name = "FunctionRadioRemove"

    <#
    # Button to submit choice and call function #
    $FunctionChangeButton = New-Object System.Windows.Forms.Button
    $FunctionChangeButton.Location = '330,20'
    $FunctionChangeButton.Size = '120,23'
    $FunctionChangeButton.Text = "Change"
    $FunctionChangeButton.Name = "FunctionChangeButton"
    #>

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

    # Adding logic so that the script will take action when the Next button gets clicked #
    $NextButton.Add_Click({
        If ($MainWindow.Controls.ContainsKey("MethodGroup")) {
            If ($MethodRadio1.Checked) {
                # Add a variable to show that it's in Individual mode #
                $SCRIPT:ScriptMode = "Individual"
            } ElseIf ($MethodRadio2.Checked) {
                $SCRIPT:ScriptMode = "ImportCSV"
            } ElseIf ($MethodRadio3.Checked){
                $SCRIPT:ScriptMode = "UpdateAndLapse"
            } Else {
                $SCRIPT:ScriptMode = "StatusCheck"
            }
            Write-Log -Message "Proceeding in $ScriptMode mode." -Level INFO -logfile $LogFilePath
            Write-Host "Proceeding in $ScriptMode mode."
            $MainWindow.Controls.RemoveByKey("MethodGroup")
            If ($ScriptMode -eq "ImportCSV") {
                # Line that sets the initial controls for the CSV GroupBox #
                $MultiModeGroup.Text = "Import CSV File"
                $MultiModeGroup.Controls.AddRange(@($MMGCSVBrowseLabel,$MMGCSVResultPathLabel,$MMGCSVBrowseButton))
                $MainWindow.Controls.Add($MultiModeGroup)
                $MainWindow.Controls.RemoveByKey("NextButton")
            } ElseIf ($SCRIPT:ScriptMode -eq "Individual") {
                # Add label and buttons to the SMGDirectionGroup #
                $SMGDirectionGroup.Controls.AddRange(@($SMGDirectionInstLabel,$SMGDirectionButtonHome,$SMGDirectionButtonWork))
            
                # Add SMGDirectionGroup to the SingleModeGroup #
                $SingleModeGroup.Controls.Add($SMGDirectionGroup)

                # Add SingleModeGroup to the Form #
                $MainWindow.Controls.Add($SingleModeGroup)
                $MainWindow.Controls.RemoveByKey("NextButton")
            } ElseIf ($SCRIPT:ScriptMode -eq "UpdateAndLapse") {
                
                $MultiModeGroup.Text = "Updates/Find Lapsed Users"
                # Create the controls for UaL mode and add the range of them to the MMG at this point #
                $MultiModeGroup.Controls.AddRange(@($MMGUaLUpdateLabel,$MMGUaLLapseCheckLabel,$MMGUaLUpdateButton,$MMGUaLLapseCheckButton))
                $MainWindow.Controls.Add($MultiModeGroup)
                $MainWindow.Controls.RemoveByKey("NextButton")

                ## Set up the form for UaL mode ##
                # Actually should probably run the functions here, so that the next thing to appear could be the list of changes #
                # That information is already in the Write-Excel function (the changes being made); it just needs to be pulled out before it makes the change #
                # Not sure if that would make for a slightly different version of the Write-Excel function where it basically runs in "List Changes" mode #
                # Like a List-ExcelChanges or something; it does the same thing, but instead of doing any writing it just finds anything that needs to change and puts that into a PSCO #
                # The Form here needs to have a Proposed Changes box, and hopefully there is a way to tick a box to confirm that it should be updated #
                # Then at the bottom there's a Confirm button to carry out the changes - you'd also need to have a routine that updates your PSCO based on the input in Proposed Changes #
                # But there should also be a button to Export, in case you don't want to make the changes (particularly useful for lapsed users; those usually need approval #
                # The Export button might need some extra help to turn the PSCO into a useable CSV; at this point I'm not 100% sure what the CSVs are going to need to look like #
                # At any rate, once you click Confirm it hands it off to Write-Excel #

            } Else {
                
                ## Set up the form for StatusCheck mode ##
                # This is kind of just for my own benefit; I think it's useful to be able to pull information from the sheet without making any changes and having to open the sheet and dig for something #
                # It would be pretty simple; just call the Search-Excel and the Fix-GetADUser functions for a particular user and organize what you find into a table #
                # Then output the table in a big box, like the OutputBox #
                # Should be relatively simple to execute #

                # Add the UserGroup controls - the textbox and label - along with a new button called SMGUserButtonCheck
                $SMGUserGroup.Controls.AddRange(@($SMGUserInstLabel,$SMGUserTextBox,$SMGUserButtonVerify,$SMGUserResultLabel))
                $SCRIPT:CheckMode = "YEP"
                
                # Add SMGUserGroup to the SingleModeGroup #
                $SingleModeGroup.Controls.Add($SMGUserGroup)

                # Add SingleModeGroup to the Form #
                $MainWindow.Controls.Add($SingleModeGroup)
                $MainWindow.Controls.RemoveByKey("NextButton")
            }
        }
    })

    # Logic for the Home button in the SMGDirectionGroup #
    $SMGDirectionButtonHome.Add_CheckStateChanged(
        {
            If ($SMGDirectionButtonHome.Checked) {
                # Set the Type variable for the function to "w" #
                $SCRIPT:SMGDirection = "Home"
                Write-Log -Message "Direction is set to $SMGDirection, so this is an external user." -Level INFO -logfile $LogFilePath
                $SMGDirectionGroup.Controls.RemoveByKey("SMGDirectionButtonWork")
            
                # Clear out the DirectionGroup and add the UserGroup #
                $SingleModeGroup.Controls.Clear()
            
                # Add label and buttons to the SMGUserGroup #
                $SMGUserGroup.Controls.AddRange(@($SMGUserInstLabel,$SMGUserTextBox,$SMGUserButtonVerify,$SMGUserResultLabel))
            
                # Add SMGUserGroup to the SingleModeGroup #
                #$SMGUserGroup.Location = '15,15'
                $SingleModeGroup.Controls.AddRange(@($SMGUserGroup))
            } <#Else {
                $SMGDirectionGroup.Controls.Add($SMGDirectionButtonWork)
                $SCRIPT:SMGDirection = ""
            }#>
        }
    )

    # Logic for the Work button in the SMGDirectionGroup #
    $SMGDirectionButtonWork.Add_CheckStateChanged(
        {
            If ($SMGDirectionButtonWork.Checked) {
                # Set the Type variable for the function to "m" #
                $SCRIPT:SMGDirection = "Work"
                Write-Log -Message "Direction is set to $SMGDirection, so this is an internal-only user." -Level INFO -logfile $LogFilePath
                $SMGDirectionGroup.Controls.RemoveByKey("SMGDirectionButtonHome")

                # Clear out the DirectionGroup and add the UserGroup (skipping TypeGroup because it's irrelevant for IO users) #
                $SingleModeGroup.Controls.Clear()
            
                # Add label and buttons to the SMGUserGroup #
                $SMGUserGroup.Controls.AddRange(@($SMGUserInstLabel,$SMGUserTextBox,$SMGUserButtonVerify,$SMGUserResultLabel))
            
                # Change the location for the User GroupBox and add SMGUserGroup to the SingleModeGroup #
                #$SMGUserGroup.Location = '15,15'
                $SingleModeGroup.Controls.AddRange(@($SMGUserGroup))
            } <#Else {
                $SMGDirectionGroup.Controls.Add($SMGDirectionButtonHome)
                $SCRIPT:SMGDirection = ""
            }#>
        }
    )


    # Logic to find the user entered in the Single Mode Username text box #
    $SMGUserButtonVerify.Add_Click(
        {
            # First, set the text box to Read-only until further action is taken #
            $SMGUserTextBox.ReadOnly = $True
        
            # Next, check the Perp sheet to see if the user is already present; if not, use Get-AdUser #
            # Require input :P #
            If ($SMGUserTextBox.Text -ne "") {
                $SCRIPT:SMGUserNameEntry = $SMGUserTextBox.Text
                $SCRIPT:SMGUserFoundOnSheet = ""
                $SMGUserResultLabel.Text = "Working..."

                # Run Fix-GetADUser to pull some preliminary data before you even check the sheet #
                $SCRIPT:FixedADInfo = Fix-GetADUser -ADInfo (Get-ADUser -Identity $SCRIPT:SMGUserNameEntry -Properties *) -Compare $false

                # Use the Search-Excel function defined above to get information about the user based on their username, assuming they have a record in the Perp sheet #
                $SCRIPT:SMGUser = Search-Excel -SearchText $SCRIPT:SMGUserNameEntry -Source $FilePath -Mode I
                
                # SMGUser.28 should be whether or not the user exists in the sheet ($TableFound) #
                If ($SCRIPT:SMGUser.28 -eq "NOPE") {   
                    If ($SCRIPT:CheckMode -ne "YEP") {
                        # If it doesn't work, leave the name displayed and ask if the user wants to check the spelling - or if this is just a new user who wouldn't have a record in the sheet #
                        $SMGUserResultLabel.Text = "Well, that record wasn't found. Is there a spelling error, or is this a new remote access user?"
                        #Start-Sleep 2
                        $SCRIPT:SMGUserFoundOnSheet = "NOPE"
                        $SCRIPT:SMGUserRow = $SCRIPT:SMGUser.26
                        $SCRIPT:SMGUserEntryNumber = $SCRIPT:SMGUser.27
                        $SCRIPT:SMGUserLastUpdated = (Get-Date -Format "MM/dd/yyyy")
                        $SCRIPT:SMGUserDateAdded = (Get-Date -Format "MM/dd/yyyy")
                        $SCRIPT:SMGUserEnabled = $SCRIPT:FixedADInfo.29
                        $SCRIPT:SMGUserEmailAddress = $SCRIPT:FixedADInfo.30
                        $SCRIPT:SMGUserInfo = $SCRIPT:FixedADInfo.31
                        # Populate the box with buttons labeled 'New User' and 'Try Again' #
                        $SMGUserButtonCancel.Text = "Try Again"
                        $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
                        $SMGUserGroup.Controls.AddRange(@($SMGUserButtonNewUser,$SMGUserButtonCancel))
                    } Else {
                        $SCRIPT:SMGUserFullName = $SCRIPT:FixedADInfo.2
                        $SCRIPT:SMGUserName = $SCRIPT:FixedADInfo.3
                        $SCRIPT:SMGUserEmpID = $SCRIPT:FixedADInfo.4
                        $SCRIPT:SMGUserDealership = $SCRIPT:FixedADInfo.7
                        $SCRIPT:SMGUserDepartment = $SCRIPT:FixedADInfo.8
                        $SCRIPT:SMGUserTitle = $SCRIPT:FixedADInfo.9
                        $SCRIPT:SMGUserEnabled = $SCRIPT:FixedADInfo.29
                        $SCRIPT:SMGUserEmailAddress = $SCRIPT:FixedADInfo.30
                        $SCRIPT:SMGUserInfo = $SCRIPT:FixedADInfo.31
                        Write-Log -Message "$SCRIPT:SMGUserFullName was not found in the spreadsheet. Proceeding with Active Directory lookup." -Level INFO -logfile $LogFilePath
                        $SMGUserResultLabel.Text = "No record found in the spreadsheet for that user. Here's what could be found in Active Directory:" + "`r`n" + "$SMGUserFullName, employee ID $SMGUserEmpID, title $SMGUserTitle, with $SMGUserDepartment at $SMGUserDealership."
                        Write-Log -Message "Active Directory lookup results: $SMGUserFullName, employee ID $SMGUserEmpID, title $SMGUserTitle, with $SMGUserDepartment at $SMGUserDealership." -Level INFO -logfile $LogFilePath
                        $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
                        $ResultGroup.Controls.AddRange(@($ResultButtonRestart,$ResultButtonQuit))
                        $MainWindow.Controls.Add($ResultGroup)
                    }
                } Else {
                    $SCRIPT:SMGUserFoundOnSheet = "YEP"
                }
                ## Pulling out code that used the Get-AdUser cmdlet to find parameters and putting it in the code for the SMGUserButtonNewUser button ##

                # Add code to extrapolate and display the parameters in the sheet for the targeted user #
                If ($SCRIPT:SMGUserFoundOnSheet -eq "YEP") {
                    
                    # Assign the information from the search to variables #
                    $SCRIPT:SMGUserRow = $SCRIPT:SMGUser.24   # Doesn't change
                    $SCRIPT:SMGUserNextRow = $SCRIPT:SMGUser.26   # Doesn't change
                    $SCRIPT:SMGUserEntryNumber = $SCRIPT:SMGUser.1   # Doesn't change
                    $SCRIPT:SMGUserFullName = $SCRIPT:FixedADInfo.2   # Doesn't change
                    $SCRIPT:SMGUserName = $SCRIPT:FixedADInfo.3   # Doesn't change
                    $SCRIPT:SMGUserEmpID = $SCRIPT:FixedADInfo.4   # Doesn't change
                    $SCRIPT:SMGUserIDLookup = $SCRIPT:SMGUser.5   # Doesn't change
                    $SCRIPT:SMGUserEarningsGroup = $SCRIPT:SMGUser.6   # Doesn't change
                    $SCRIPT:SMGUserDealership = $SCRIPT:FixedADInfo.7   # Doesn't change
                    $SCRIPT:SMGUserDepartment = $SCRIPT:FixedADInfo.8   # Doesn't change
                    $SCRIPT:SMGUserTitle = $SCRIPT:FixedADInfo.9   # Doesn't change
                    $SCRIPT:SMGUserLastLoginCurrent = $SCRIPT:SMGUser.10   # Doesn't change
                    $SCRIPT:SMGUserLastLoginFormula = $SCRIPT:SMGUser.11   # Doesn't change
                    $SCRIPT:SMGUserLastLoginHistorical = $SCRIPT:SMGUser.12   # Doesn't change
                    $SCRIPT:SMGUserLast30Formula = $SCRIPT:SMGUser.13   # Doesn't change
                    $SCRIPT:SMGUserLast30Current = $SCRIPT:SMGUser.14   # Doesn't change
                    $SCRIPT:SMGUserRDP = "Y"   # Should always be "Y" for anyone coming through this script, since we're always dealing with RDP users
                    $SCRIPT:SMGUserVPN = $SCRIPT:SMGUser.16   # Doesn't change
                    $SCRIPT:SMGUserMemberships = $SCRIPT:SMGUser.17   # Changed by LocationGroup
                    $SCRIPT:SMGUserAccessType = $SCRIPT:SMGUser.18   # Changed by TypeGroup
                    $SCRIPT:SMGUserDesktopName = $SCRIPT:SMGUser.19   # Changed by DesktopGroup
                    $SCRIPT:SMGUserDateAdded = $SCRIPT:SMGUser.20   # Doesn't change
                    $SMGUserPreviousUpdate = $SCRIPT:SMGUser.21   # Used by CheckMode
                    $SCRIPT:SMGUserLastUpdated = (Get-Date -Format "MM/dd/yyyy")
                    $SCRIPT:SMGUserStatus = $SCRIPT:SMGUser.22   # Changed by FunctionGroup, actually in the SubmitButton's .AddClick()
                    $SCRIPT:SMGUserNotes = $SCRIPT:SMGUser.23   # Potentially changed by FunctionGroup - NOT YET IMPLEMENTED, might just leave it at the moment
                    $SCRIPT:SMGNextEntryNumber = $SCRIPT:SMGUser.27
                    $SCRIPT:SMGUserEnabled = $SCRIPT:FixedADInfo.29
                    $SCRIPT:SMGUserEmailAddress = $SCRIPT:FixedADInfo.30
                    $SCRIPT:SMGUserInfo = $SCRIPT:FixedADInfo.31
                    If ($SCRIPT:CheckMode -ne "YEP") {
                        Write-Log -Message "$SCRIPT:SMGUserFullName was found in the spreadsheet on row $SCRIPT:SMGUserRow." -Level INFO -logfile $LogFilePath
                        $SMGUserResultLabel.Text = "$SCRIPT:SMGUserFullName was found on row $SCRIPT:SMGUserRow." + "`r`n" + "Their title is $SCRIPT:SMGUserTitle, employee ID $SCRIPT:SMGUserEmpID, with " + $SCRIPT:SMGUserDepartment + " at $SCRIPT:SMGUserDealership." + "`r`n" + "Is this correct?"
                        Write-Log -Message "Their title is $SCRIPT:SMGUserTitle, employee ID $SCRIPT:SMGUserEmpID, with $SCRIPT:SMGUserDepartment at $SCRIPT:SMGUserDealership." -Level INFO -logfile $LogFilePath
                        $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
                        $SMGUserGroup.Controls.AddRange(@($SMGUserButtonConfirm,$SMGUserButtonCancel))
                    } Else {
                        Write-Log -Message "$SCRIPT:SMGUserFullName was found in the spreadsheet on row $SCRIPT:SMGUserRow." -Level INFO -logfile $LogFilePath
                        $SMGUserResultLabel.Text = "$SCRIPT:SMGUserFullName was found on row $SCRIPT:SMGUserRow." + "`r`n" + "Their title is $SCRIPT:SMGUserTitle, employee ID $SCRIPT:SMGUserEmpID, with " + $SCRIPT:SMGUserDepartment + " at $SCRIPT:SMGUserDealership." + "`r`n" + "They log in via $SCRIPT:SMGUserAccessType to $SCRIPT:SMGUserDesktopName."
                        Write-Log -Message "Their title is $SCRIPT:SMGUserTitle, employee ID $SCRIPT:SMGUserEmpID, with $SCRIPT:SMGUserDepartment at $SCRIPT:SMGUserDealership." -Level INFO -logfile $LogFilePath
                        If (($SCRIPT:SMGUserDateAdded -ne "") -and ($SCRIPT:SMGUserStatus -ne "I")) {
                            $SMGUserResultLabel.Text += "`r`n" + "$SCRIPT:SMGUserFullName has had access since $SCRIPT:SMGUserDateAdded, last logging via $SCRIPT:SMGUserAccessType in $SCRIPT:SMGUserLastLoginCurrent."
                            Write-Log -Message "$SCRIPT:SMGUserFullName has had access since $SCRIPT:SMGUserDateAdded, last logging in $SCRIPT:SMGUserLastLoginCurrent." -Level INFO -logfile $LogFilePath
                        } ElseIf (($SCRIPT:SMGUserDateAdded -ne "") -and ($SCRIPT:SMGUserStatus -eq "I")) {
                            $SMGUserResultLabel.Text += "`r`n" + "$SCRIPT:SMGUserFullName previously had access from $SCRIPT:SMGUserDateAdded to $SMGUserPreviousUpdate;" + "`r`n" + "Their last login was via $SCRIPT:SMGUserAccessType on $SCRIPT:SMGUserLastLoginCurrent."
                            Write-Log -Message "$SCRIPT:SMGUserFullName previously had access from $SCRIPT:SMGUserDateAdded to $SMGUserPreviousUpdate, last logging in $SCRIPT:SMGUserLastLoginCurrent." -Level INFO -logfile $LogFilePath
                        } ElseIf (($SCRIPT:SMGUserDateAdded -eq "") -and ($SCRIPT:SMGUserStatus -eq "I")) {
                            $SMGUserResultLabel.Text += "`r`n" + "$SCRIPT:SMGUserFullName previously had access until $SMGUserPreviousUpdate, last logging in $SCRIPT:SMGUserLastLoginCurrent."
                            Write-Log -Message "$SCRIPT:SMGUserFullName previously had access until $SMGUserPreviousUpdate;" + "`r`n" + "Their last login was via $SCRIPT:SMGUserAccessType on $SCRIPT:SMGUserLastLoginCurrent." -Level INFO -logfile $LogFilePath
                        } Else {
                            $SMGUserResultLabel.Text += "`r`n" + "$SCRIPT:SMGUserFullName last logged in via $SCRIPT:SMGUserAccessType on $SCRIPT:SMGUserLastLoginCurrent."
                            Write-Log -Message "$SCRIPT:SMGUserFullName last logged in $SCRIPT:SMGUserLastLoginCurrent." -Level INFO -logfile $LogFilePath
                        }
                        $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
                        $ResultGroup.Controls.AddRange(@($ResultButtonRestart,$ResultButtonQuit))
                        $MainWindow.Controls.Add($ResultGroup)  
                    }
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
            $SMGUserResultLabel.Text = "$SMGUserFullName, employee ID $SMGUserEmpID, with $SMGUserDepartment at $SMGUserDealership."
            $SMGUserResultLabel.Location = '120,17'
            $SMGUserGroup.Controls.Add($SMGUserResultLabel)
            $SMGUserGroup.Size = '550,50'
            If ($SCRIPT:SMGDirection -eq "Work") {
                $SMGDesktopGroup.Location = '15,10'
            }
            If ($SCRIPT:SMGUserFoundOnSheet -eq "YEP") {
                # Add code to place UI elements for choosing the desktop already in the sheet #
                $SMGDesktopTextBox.Text = "$SMGUserDesktopName"
                $SMGDesktopTextBox.ReadOnly = $True
                $SMGDesktopInstLabel.Text = "Is this the correct Desktop Name?"
                $SMGDesktopGroup.Controls.AddRange(@($SMGDesktopInstLabel,$SMGDesktopTextBox,$SMGDesktopButtonAccept,$SMGDesktopButtonReject))
                $SingleModeGroup.Controls.Add($SMGDesktopGroup)
            } Else {
                $SMGDesktopGroup.Controls.AddRange(@($SMGDesktopInstLabel,$SMGDesktopTextBox,$SMGDesktopButtonVerify))
                #$SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDC,$SMGLocButtonDB))
                $SingleModeGroup.Controls.Add($SMGDesktopGroup)
                $SMGDesktopTextBox.Focus() 
            }
        }
    )

    # Logic needed for the User Cancel button #
    $SMGUserButtonCancel.Add_Click(
        {
            If ($SCRIPT:SMGUserFoundOnSheet -eq "NOPE") {
                $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonNewUser")
            }
            $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonConfirm")
            $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonCancel")
            $SMGUserResultLabel.Text = ""
            $SMGUserTextBox.Text = ""
            $SMGUserTextBox.ReadOnly = $False
            $SMGUserGroup.Controls.Add($SMGUserButtonVerify)
        }
    )

    # Logic for the New User button in the User Group #
    $SMGUserButtonNewUser.Add_Click(
        {
            
            $SCRIPT:SMGNewUser = "YEP"
            # Try using the inputted username as the -Identity parameter of a Get-AdUser cmdlet #
            Try {
                $SCRIPT:FixedADInfo = Fix-GetADUser -ADInfo (Get-ADUser -Identity $SCRIPT:SMGUserNameEntry -Properties *)
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
                
                # Need to include the line of the sheet where this will be added #
                # Get enough information from the Get-AdUser to fill in the whole row of the sheet, including the date #
                # At least as much as possible here #

                # Show the user information on the associate in question and ask for confirmation #
                $SCRIPT:SMGUserFullName = $SCRIPT:FixedADInfo.2
                $SCRIPT:SMGUserName = $SCRIPT:FixedADInfo.3
                $SCRIPT:SMGUserEmpID = $SCRIPT:FixedADInfo.4
                $SCRIPT:SMGUserDealership = $SCRIPT:FixedADInfo.7
                $SCRIPT:SMGUserDepartment = $SCRIPT:FixedADInfo.8
                $SCRIPT:SMGUserTitle = $SCRIPT:FixedADInfo.9
                $SCRIPT:SMGUserLastLoginFormula = "=IFNA(VLOOKUP(C$SCRIPT:SMGUserRow,current_data,3,FALSE),`"`")"
                $SCRIPT:SMGUserLast30Formula = "=IFNA(VLOOKUP(C$SCRIPT:SMGUserRow,current_data,5,FALSE),`"0`")"
                $SCRIPT:SMGUserLast30Current = "0"
                $SCRIPT:SMGUserRDP = "Y"
                $SCRIPT:SMGUserDateAdded = (Get-Date -Format "MM/dd/yyyy")
                $SCRIPT:SMGTypeSwitch = ""
                $SMGUserResultLabel.Text = "$SMGUserFullName, employee ID $SMGUserEmpID, title $SMGUserTitle, with $SMGUserDepartment at $SMGUserDealership." + "`r`n" + "They will be placed on row $SCRIPT:SMGUserRow, labeled as number $SCRIPT:SMGUserEntryNumber. Is this correct?"
                $SMGUserButtonConfirm.Text = "Yes"
                $SMGUserButtonCancel.Text = "No"
                $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
                $SMGUserGroup.Controls.RemoveByKey("SMGUserButtonNewUser")
                $SMGUserGroup.Controls.AddRange(@($SMGUserButtonConfirm,$SMGUserButtonCancel))
            }
        }
    )

    # Logic to reject the proposed desktop offered by the sheet #
    $SMGDesktopButtonReject.Add_Click(
        {
            
            # Basically just remove the text from the TextBox and make it writeable and repopulate the Group with a different label text and the Verify button #
            $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonReject")
            $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonAccept")
            $SMGDesktopGroup.Controls.Add($SMGDesktopButtonVerify)
            $SMGDesktopInstLabel.Text = "Enter the associate's office PC name:"
            $SMGDesktopTextBox.ReadOnly = $False
            $SMGDesktopTextBox.Text = ""
            $SMGDesktopTextBox.Focus()

        }
    )

    # Logic to pull information about the desktop offered by the sheet and store it, then move on to the next step #
    $SMGDesktopButtonAccept.Add_Click(
        {
            # First ping the box and show the result #
            $SMGDesktopInstLabel.Text += "`r`n" + "Testing connectivity to the target PC:"
            $SMGDesktopSuccess = ""
            $SMGPingResult = ping $SMGDesktopName
            #Write-Host $SMGPingResult
            If (($SMGPingResult -like "reply from*") -like "false*") {
                # Do something with the ping result #
                $SMGDesktopInstLabel.Text += "`r`n" + "Ping failed! Sucks!"
                $SMGDesktopSuccess = "NOPE"
            } Else {
                $SMGDesktopInstLabel.Text += "`r`n" + "Connectivity established! Sweet!"
                $SMGDesktopSuccess = "YEP"
            } 
            
            

            # Next pull in the desktop from the text and pass it to Get-ADComputer #
            $SMGDesktopNameEntry = $SMGDesktopTextBox.Text
            $SCRIPT:SMGDesktop = Get-ADComputer -Identity $SMGDesktopNameEntry -Properties *

            # Show the desktop information on the PC in question and ask for confirmation #
            $SCRIPT:SMGDesktopName = $SMGDesktop.Name
            $SCRIPT:SMGDesktopSamAccountName = $SMGDesktop.SamAccountName
            $SCRIPT:SMGDesktopDept = $SMGDesktop.Department
            $SCRIPT:SMGDesktopDesc = $SMGDesktop.Description
            $SCRIPT:SMGDesktopIP = $SMGDesktop.IPv4Address
            $SCRIPT:SMGDesktopCo = $SMGDesktop.Company

            #Write-Host "DesktopSuccess is set to $SMGDesktopSuccess."

            $SMGDesktopGroup.Controls.Add($SMGDesktopResultLabel)
            # Show results of Get-ADComputer and prompt for confirmation, moving on to next step from there #
            If ($SMGDesktopSuccess -eq "NOPE") {
                $SMGDesktopResultLabel.Text = "Would you like to try a different desktop, or quit?"
                $SMGDesktopButtonReject.Text = "Try Again"
                $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonAccept")
                $SMGDesktopGroup.Controls.Add($SMGDesktopButtonAbort)
            } ElseIf ($SMGDesktopSuccess -eq "YEP") {
                $SMGDesktopResultLabel.Text = "$SMGDesktopName, described as $SMGDesktopDesc, in $SMGDesktopDept at $SMGDesktopCo, with IP address $SMGDesktopIP. Is this correct?"
                $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonAccept")
                $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonReject")
                $SMGDesktopGroup.Controls.AddRange(@($SMGDesktopResultLabel,$SMGDesktopButtonConfirm,$SMGDesktopButtonCancel))
            }
        }
    )

    # Not a ton of Logic here, just quitting #
    $SMGDesktopButtonAbort.Add_Click(
        {
            Exit(1)
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
                $SMGDesktopInstLabel.Text = "Working..."

                # Try using the inputted PC name as the -Identity parameter of a Get-ADComputer cmdlet #
                Try {
                    $SCRIPT:SMGDesktop = Get-ADComputer -Identity $SMGDesktopNameEntry -Properties *
                    $SMGDesktopSuccess = "YEP"
                } Catch {
                
                    # If it doesn't work, reset the form and display a message asking nicely for the user to GET IT RIGHT NEXT TIME YOU DUMMY #
                    $SMGDesktopInstLabel.Text += "`r`n" + "Hmm, that entry didn't quite work. Check the computername and try again!"
                    Start-Sleep 2
                    $SMGDesktopTextBox.ReadOnly = $False
                    $SMGDesktopSuccess = "NOPE"
                }

                # Make sure the Get-ADUser works before continuing #
                If ($SMGDesktopSuccess -eq "YEP") {
                
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
                        # Do something with the ping result #
                        $SMGDesktopInstLabel.Text += "`r`n" + "Ping failed! Sucks!"
                        $SMGDesktopPing = "NOPE"
                    } Else {
                        $SMGDesktopInstLabel.Text += "`r`n" + "Connectivity established! Sweet!"
                        $SMGDesktopPing = "YEP"
                    }

                    $SMGDesktopResultLabel.Text = "$SMGDesktopName, described as $SMGDesktopDesc, in $SMGDesktopDept at $SMGDesktopCo, with IP address $SMGDesktopIP. Is this correct?"
                    $SMGDesktopGroup.Controls.RemoveByKey("SMGDesktopButtonVerify")
                    $SMGDesktopGroup.Controls.AddRange(@($SMGDesktopResultLabel,$SMGDesktopButtonConfirm,$SMGDesktopButtonCancel))
                }

            } Else {
                $SMGDesktopInstLabel.Text = "Sorry, input is required here. Please try again!"
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
            $SCRIPT:SMGUserDesktopName = $SMGDesktopName
            If ($SCRIPT:SMGDirection -eq "Work") {
                # Bring up the FunctionGroup to choose between adding and removing access #
                $FunctionGroup.Controls.AddRange(@($FunctionRadioEnable,$FunctionRadioRemove,$FunctionSubmitButton))
                $FunctionSubmitButton.Location = '452,20'
                $SMGOutputBox.Location = '15,115'
                $SingleModeGroup.Controls.Add($SMGOutputBox)
                $MainWindow.Controls.Add($FunctionGroup)
            } Else {
                
                # If the user is on the sheet, present the current setting to the user and give them a confirm/switch option. If not, just proceed as before with Windows/Mac option #
                If ($SMGUserFoundOnSheet -eq "YEP") {
                    If ($SMGUserAccessType -eq "TS") {
                        $SMGTypeSet = $True
                        #$SMGTypeButtonWin.Checked = $True
                        $SMGTypeLabel.Text = "Records show that this is a Windows (TS) user. Confirm or Switch?"
                        $SCRIPT:NLASwitch = "w"
                        $SMGTypeGroup.Controls.AddRange(@($SMGTypeLabel,$SMGTypeButtonConfirm,$SMGTypeButtonSwitch))
                    } ElseIf ($SMGUserAccessType -eq "HTML5") {
                        $SMGTypeSet = $True
                        #$SMGTypeButtonMac.Checked = $True
                        $SMGTypeLabel.Text = "Records show that this is a Mac (HTML5) user. Confirm or Switch?"
                        $SCRIPT:NLASwitch = "m"
                        $SMGTypeGroup.Controls.AddRange(@($SMGTypeLabel,$SMGTypeButtonConfirm,$SMGTypeButtonSwitch))
                    } Else {
                        
                        # If it's a VPN user, provide this information. If the user is not TS, HTML5, or VPN, remark on this. Either way, the script should just proceed #
                        If ($SMGUserVPN -eq "Y") {
                            $SMGTypeLabel.Text = "This user appears to be a VPN-only user. Interesting."
                        } Else {
                            $SMGTypeLabel.Text = "Something unusual is happening. Might want to restart."
                        }
                        $SMGTypeGroup.Controls.AddRange(@($SMGTypeLabel,$SMGTypeButtonConfirm))
                    }
                } Else {
                    Write-Log -Message "Desktop established; moving on to Type setting for this first-time user." -Level INFO -logfile $LogFilePath
                    $SMGTypeSet = $False
                    # Proceed with first-time setup options of Windows or Mac #
                    $SMGTypeGroup.Controls.AddRange(@($SMGTypeLabel,$SMGTypeButtonWin,$SMGTypeButtonMac))
                }

                $SingleModeGroup.Controls.Add($SMGTypeGroup)
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


    
    # Logic for the Windows button in the SMGTypeGroup #
    $SMGTypeButtonWin.Add_CheckStateChanged(
        {
            If ($SMGTypeButtonWin.Checked) {
                # Set the Type variable for the function to "w" #
                $SCRIPT:NLASwitch = "w"
                Write-Log -Message "NLASwitch is set to $NLASwitch, so this is a Windows user." -Level INFO -logfile $LogFilePath
                $SCRIPT:SMGUserAccessType = "TS"
                $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonMac")
                <#
                If (!($SingleModeGroup.Contains($SMGLocGroup))) {
                    $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDB,$SMGLocButtonDC))
                    $SingleModeGroup.Controls.Add($SMGLocGroup)
                }#>
            } Else {
                $SMGTypeGroup.Controls.Add($SMGTypeButtonMac)
                $SCRIPT:NLASwitch = ""
            }
            If ($SMGTypeSet -eq $True) {
                $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonConfirm))
            } Else {
                $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDB,$SMGLocButtonDC))
                
            }
            $SingleModeGroup.Controls.Add($SMGLocGroup)
        }
    )

    # Logic for the Mac button in the SMGTypeGroup #
    $SMGTypeButtonMac.Add_CheckStateChanged(
        {
            If ($SMGTypeButtonMac.Checked) {
                # Set the Type variable for the function to "m" #
                $SCRIPT:NLASwitch = "m"
                Write-Log -Message "NLASwitch is set to $NLASwitch, so this is a Mac user." -Level INFO -logfile $LogFilePath
                $SCRIPT:SMGUserAccessType = "HTML5"
                $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonWin")
                <#
                If (!($SingleModeGroup.Contains($SMGLocGroup))) {
                }#>
            } Else {
                $SMGTypeGroup.Controls.Add($SMGTypeButtonWin)
                $SCRIPT:NLASwitch = ""
            }
            If ($SMGTypeSet -eq $True) {
                $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonConfirm))
            } Else {
                $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDB,$SMGLocButtonDC))
                
            }
            $SingleModeGroup.Controls.Add($SMGLocGroup)
        }
    )

    # Logic needed for the Type Confirm button #
    $SMGTypeButtonConfirm.Add_Click(
        {
            $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonConfirm")
            $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonSwitch")
            $SCRIPT:SMGTypeSwitch = ""
            If ($SCRIPT:NLASwitch -eq "w") {
                $SMGTypeLabel.Text = "Proceeding with setup for $SCRIPT:SMGUserFullName on Windows (TS)."
                $SCRIPT:SMGUserAccessType = "TS"
            } ElseIf ($SCRIPT:NLASwitch -eq "m") {
                $SMGTypeLabel.Text = "Proceeding with setup for $SCRIPT:SMGUserFullName on Mac (HTML5)."
                $SCRIPT:SMGUserAccessType = "HTML5"
            }
            If ($SMGUserFoundOnSheet -eq "YEP") {
                If ($SCRIPT:SMGUserMemberships -like "*Data Bank*") {
                    $SCRIPT:SMGLocation = "DB"
                    $SMGLocLabel.Text = "User $SCRIPT:SMGUserFullName is connecting to Data Bank. Please confirm."
                    $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonConfirm))
                } ElseIf ($SCRIPT:SMGUserMemberships -notlike "*VPN*") {
                    $SCRIPT:SMGLocation = "DC"
                    $SMGLocLabel.Text = "User $SCRIPT:SMGUserFullName is connecting to the Data Center. Please confirm."
                    $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonConfirm))
                }
            } Else {
                $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDB,$SMGLocButtonDC))
            }
            $SingleModeGroup.Controls.Add($SMGLocGroup)
        }
    )

    # Logic needed for the Type Switch button #
    $SMGTypeButtonSwitch.Add_Click(
        {
            $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonConfirm")
            $SMGTypeGroup.Controls.RemoveByKey("SMGTypeButtonSwitch")
            If ($SCRIPT:NLASwitch -eq "w") {
                $SCRIPT:SMGTypeSwitch = "WinToMac"
                $SMGTypeLabel.Text = "Switching access for $SCRIPT:SMGUserFullName from Windows (TS) to Mac (HTML5)."
                $SCRIPT:SMGUserAccessType = "HTML5"
            } ElseIf ($SCRIPT:NLASwitch -eq "m") {
                $SCRIPT:SMGTypeSwitch = "MacToWin"
                $SMGTypeLabel.Text = "Switching access for $SCRIPT:SMGUserFullName from Mac (HTML5) to Windows (TS)."
                $SCRIPT:SMGUserAccessType = "TS"
            }
            If ($SMGUserFoundOnSheet -eq "YEP") {
                If ($SCRIPT:SMGUserMemberships -like "*Data Bank*") {
                    $SCRIPT:SMGLocation = "DB"
                    $SMGLocLabel.Text = "User $SCRIPT:SMGUserFullName is connecting to Data Bank. Please confirm."
                    $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonConfirm))
                } ElseIf ($SCRIPT:SMGUserMemberships -notlike "*VPN*") {
                    $SCRIPT:SMGLocation = "DC"
                    $SMGLocLabel.Text = "User $SCRIPT:SMGUserFullName is connecting to the Data Center. Please confirm."
                    $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonConfirm))
                }
            } Else {
                $SMGLocGroup.Controls.AddRange(@($SMGLocLabel,$SMGLocButtonDB,$SMGLocButtonDC))
            }
            $SingleModeGroup.Controls.Add($SMGLocGroup)
        }
    )


    # Logic for the DC button in the SMGLocGroup #
    $SMGLocButtonDC.Add_CheckStateChanged(
        {
            If ($SMGLocButtonDC.Checked) {
                # Set the Type variable for the function to "w" #
                $SCRIPT:SMGLocation = "DC"
                Write-Log -Message "Location is set to $SMGLocation, so this is a Data Center user." -Level INFO -logfile $LogFilePath
                $SMGLocGroup.Controls.RemoveByKey("SMGLocButtonDB")
                $SCRIPT:SMGUserMemberships = "SSO Remote Access Associate"
            
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
                Write-Log -Message "Location is set to $SMGLocation, so this is a Data Bank user." -Level INFO -logfile $LogFilePath
                $SMGLocGroup.Controls.RemoveByKey("SMGLocButtonDC")
                $SCRIPT:SMGUserMemberships = "SSO Remote Access Data Bank Associate"

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

    # Logic needed for the Location Confirm button #
    $SMGLocButtonConfirm.Add_Click(
        {
            $SMGLocGroup.Controls.RemoveByKey("SMGLocButtonConfirm")
            If ($SCRIPT:SMGLocation -eq "DB") {
                $SMGLocLabel.Text = "Proceeding with setup for $SCRIPT:SMGUserFullName on Data Bank Pulse."
                If ((!($SCRIPT:SMGUserVPN -eq "Y")) -or (!($SCRIPT:SMGUserDepartment -eq "IS"))) {
                    $SCRIPT:SMGUserMemberships = "SSO Remote Access Data Bank Associate"
                } ElseIf ((!($SCRIPT:SMGUserVPN -eq "Y")) -and ($SCRIPT:SMGUserDepartment -eq "IS")) {
                    $SCRIPT:SMGUserMemberships = "SSO Remote Access IS"
                }
            } ElseIf ($SCRIPT:SMGLocation -eq "DC") {
                $SMGLocLabel.Text = "Proceeding with setup for $SCRIPT:SMGUserFullName on Data Center Pulse."
                If ((!($SCRIPT:SMGUserVPN -eq "Y")) -or (!($SCRIPT:SMGUserDepartment -eq "IS"))) {
                    $SCRIPT:SMGUserMemberships = "SSO Remote Access Associate"
                } ElseIf ((!($SCRIPT:SMGUserVPN -eq "Y")) -and ($SCRIPT:SMGUserDepartment -eq "IS")) {
                    $SCRIPT:SMGUserMemberships = "SSO Remote Access IS"
                }
            }
            If ($SMGUserFoundOnSheet -eq "YEP") {
                If ($SCRIPT:SMGUserStatus -eq "I") {
                    $FunctionRadioEnable.Checked = $True
                } ElseIf ($SCRIPT:SMGUserStatus -eq "A") {
                    $FunctionRadioRemove.Checked = $True
                }
            } 
            # Bring up the FunctionGroup to choose between adding and removing access #
            $FunctionGroup.Controls.AddRange(@($FunctionRadioEnable,$FunctionRadioRemove,$FunctionSubmitButton))
            $FunctionSubmitButton.Location = '452,20'
            $SingleModeGroup.Controls.Add($SMGOutputBox)
            $MainWindow.Controls.Add($FunctionGroup)
        }
    )

    ###
    ### MMGCSV SECTION - NEEDS TO BE UPDATED TO USE THE PERP SHEET, BOTH TO SEARCH RECORDS AND TO RECORD THEM ###
    ###

    # Logic to determine what happens when you click the Browse button #
    $MMGCSVBrowseButton.Add_Click(
        {
            # Run the OpenFileDialog first #
            $MMGCSVBrowseDialog.ShowDialog()

            # Store the path to the selected CSV file in a variable, called $MMGCSVPathName #
            $SCRIPT:MMGCSVPathName = $MMGCSVBrowseDialog.FileName

            # Remove the BrowseButton control after clicking it #
            $MainWindow.Controls.RemoveByKey("BrowseButton")

            # Change the Results label to ask if the file is the right one #
            $MMGCSVBrowseLabel.Text = "Please confirm that this is the correct file!"
            $MMGCSVResultPathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $MMGCSVResultPathLabel.Text = $MMGCSVPathName

            # Replace the controls in the $MultiModeGroup to show the Confirm and Cancel buttons #
            $MultiModeGroup.Controls.Clear()
            $MultiModeGroup.Controls.AddRange(@($MMGCSVBrowseLabel,$MMGCSVResultPathLabel,$MMGCSVConfirmButton,$MMGCSVCancelButton))
        }
    )

    # Logic for when the user confirms their choice of CSV file by clicking the Confirm button #
    $MMGCSVConfirmButton.Add_Click(
        {
            $MMGCSVPath = $MMGCSVPathName
            #Write-Host "Two variables to check - CSVPath: $MMGCSVPath, and CSVPathName: $MMGCSVPathName"

            # Reset the CSVGroup and add $MMGCSVOutputBox to display imported CSV data #
            $MultiModeGroup.Controls.Clear()
            $MultiModeGroup.Controls.AddRange(@($MMGCSVBrowseLabel,$MMGCSVResultPathLabel))
            $MainWindow.Controls.Add($MMGCSVOutputBox)
        
            # Change $MMGCSVBrowseLabel text, reset $MMGCSVResultPathLabel, import the chosen CSV and show the output in $MMGCSVOutputBox # 
            $MMGCSVBrowseLabel.Text = "Here are the results:"
            $MMGCSVResultPathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
            $SCRIPT:MMGCSVData = Import-Csv -Path $MMGCSVPath 

            # Take each username and run F-GAU and S-E against it, storing the results in variables and storing the variables in an ArrayList #

            #$MMGCSVUsernames = $SCRIPT:MMGCSVData | Select-Object -Property Username
            ForEach ($CSVRecord in $MMGCSVData) {
                $CSVUsername = $CSVRecord.Username
                $CSVDesktop = $CSVRecord.Desktop
                $CSVLocation = $CSVRecord.Location
                $CSVType = $CSVRecord.Type
                $SCRIPT:CSVSEPSCO = Search-Excel -Source $FilePath -SearchText $CSVUsername -Mode I
                $SCRIPT:CSVFGAUPSCO = Fix-GetADUser -ADInfo (Get-ADUser -Identity $CSVUsername -Properties *) -Compare $False

                ### Basically do the same thing as SMGUserVerifyButton here, build up a PSCO from a combination of S-E and F-GAU ###
                # CSVSEPSCO.28 should be whether or not the user exists in the sheet ($TableFound) #
                If ($SCRIPT:CSVSEPSCO.28 -eq "NOPE") {
                    $SCRIPT:CSVFoundOnSheet = "NOPE"
                    $SCRIPT:CSVRow = $SCRIPT:CSVSEPSCO.26
                    $SCRIPT:CSVEntryNumber = $SCRIPT:CSVSEPSCO.27
                    $SCRIPT:CSVFullName = $SCRIPT:CSVFGAUPSCO.2
                    $SCRIPT:CSVUserName = $SCRIPT:CSVFGAUPSCO.3
                    $SCRIPT:CSVEmpID = $SCRIPT:CSVFGAUPSCO.4
                    $SCRIPT:CSVDealership = $SCRIPT:CSVFGAUPSCO.7
                    $SCRIPT:CSVDepartment = $SCRIPT:CSVFGAUPSCO.8
                    $SCRIPT:CSVTitle = $SCRIPT:CSVFGAUPSCO.9
                    $SCRIPT:CSVRDP = "Y"
                    $SCRIPT:CSVMemberships = $CSVLocation
                    $SCRIPT:CSVAccessType = $CSVType
                    $SCRIPT:CSVDesktopName = $CSVDesktop
                    $SCRIPT:CSVLastUpdated = (Get-Date -Format "MM/dd/yyyy")
                    $SCRIPT:CSVDateAdded = (Get-Date -Format "MM/dd/yyyy")
                    $SCRIPT:CSVEnabled = $SCRIPT:CSVFGAUPSCO.29
                    $SCRIPT:CSVEmailAddress = $SCRIPT:CSVFGAUPSCO.30
                    $SCRIPT:CSVInfo = $SCRIPT:CSVFGAUPSCO.31
                } Else {
                    $SCRIPT:CSVFoundOnSheet = "YEP"
                }
                ## Pulling out code that used the Get-AdUser cmdlet to find parameters and putting it in the code for the SMGUserButtonNewUser button ##

                # Add code to extrapolate and display the parameters in the sheet for the targeted user #
                If ($SCRIPT:CSVFoundOnSheet -eq "YEP") {
                    
                    # Assign the information from the search to variables #
                    $SCRIPT:CSVRow = $SCRIPT:CSVSEPSCO.24
                    $SCRIPT:CSVNextRow = $SCRIPT:CSVSEPSCO.26
                    $SCRIPT:CSVEntryNumber = $SCRIPT:CSVSEPSCO.1
                    $SCRIPT:CSVFullName = $SCRIPT:CSVFGAUPSCO.2
                    $SCRIPT:CSVUserName = $SCRIPT:CSVFGAUPSCO.3
                    $SCRIPT:CSVEmpID = $SCRIPT:CSVFGAUPSCO.4
                    $SCRIPT:CSVIDLookup = $SCRIPT:CSVSEPSCO.5
                    $SCRIPT:CSVEarningsGroup = $SCRIPT:CSVSEPSCO.6
                    $SCRIPT:CSVDealership = $SCRIPT:CSVFGAUPSCO.7
                    $SCRIPT:CSVDepartment = $SCRIPT:CSVFGAUPSCO.8
                    $SCRIPT:CSVTitle = $SCRIPT:CSVFGAUPSCO.9
                    $SCRIPT:CSVLastLoginCurrent = $SCRIPT:CSVSEPSCO.10
                    $SCRIPT:CSVLastLoginFormula = $SCRIPT:CSVSEPSCO.11
                    $SCRIPT:CSVLastLoginHistorical = $SCRIPT:CSVSEPSCO.12
                    $SCRIPT:CSVLast30Formula = $SCRIPT:CSVSEPSCO.13
                    $SCRIPT:CSVLast30Current = $SCRIPT:CSVSEPSCO.14
                    $SCRIPT:CSVRDP = "Y"   # Should always be "Y" for anyone coming through this script, since we're always dealing with RDP users
                    $SCRIPT:CSVVPN = $SCRIPT:CSVSEPSCO.16
                    $SCRIPT:CSVMemberships = $SCRIPT:CSVSEPSCO.17
                    $SCRIPT:CSVAccessType = $SCRIPT:CSVSEPSCO.18
                    $SCRIPT:CSVDesktopName = $SCRIPT:CSVSEPSCO.19
                    $SCRIPT:CSVDateAdded = $SCRIPT:CSVSEPSCO.20
                    $SCRIPT:CSVLastUpdated = (Get-Date -Format "MM/dd/yyyy")
                    $SCRIPT:CSVStatus = $SCRIPT:CSVSEPSCO.22
                    $SCRIPT:CSVNotes = $SCRIPT:CSVSEPSCO.23
                    $SCRIPT:CSVNextEntryNumber = $SCRIPT:CSVSEPSCO.27
                    $SCRIPT:CSVEnabled = $SCRIPT:CSVFGAUPSCO.29
                    $SCRIPT:CSVEmailAddress = $SCRIPT:CSVFGAUPSCO.30
                    $SCRIPT:CSVInfo = $SCRIPT:CSVFGAUPSCO.31
                }

                # Create the PSCO now, with the accumulated info from S-E and F-GAU #
                $CSVWriteObject = [PSCustomObject]@{
                    1 = $SCRIPT:CSVEntryNumber
                    2 = $SCRIPT:CSVFullName
                    3 = $SCRIPT:CSVUserName
                    4 = $SCRIPT:CSVEmpID
                    5 = $SCRIPT:CSVIDLookup
                    6 = $SCRIPT:CSVEarningsGroup
                    7 = $SCRIPT:CSVDealership
                    8 = $SCRIPT:CSVDepartment
                    9 = $SCRIPT:CSVTitle
                    10 = $SCRIPT:CSVLastLoginCurrent
                    11 = $SCRIPT:CSVLastLoginFormula
                    12 = $SCRIPT:CSVLastLoginHistorical
                    13 = $SCRIPT:CSVLast30Formula
                    14 = $SCRIPT:CSVLast30Current
                    15 = $SCRIPT:CSVRDP
                    16 = $SCRIPT:CSVVPN
                    17 = $SCRIPT:CSVMemberships
                    18 = $SCRIPT:CSVAccessType
                    19 = $SCRIPT:CSVDesktopName
                    20 = $SCRIPT:CSVDateAdded
                    21 = $SCRIPT:CSVLastUpdated
                    22 = $SCRIPT:CSVStatus
                    23 = $SCRIPT:CSVNotes
                    24 = $SCRIPT:CSVRow
                }

                # Add each PSCO as it is created to the ArrayList to feed to FunctionSubmit #
                $MMGCSVPSCOList.Add($CSVWriteObject) | Out-Null

            }
            $MMGCSVOutput = $MMGCSVData | Select-Object -Property FullName | Out-String
            $MMGCSVOutputBox.Text = $MMGCSVOutput

            # Add the $FunctionGroup so the user can choose to Enable or Remove access #
            # Add radio buttons and Submit button to the FunctionGroup #
            $FunctionGroup.Controls.AddRange(@($FunctionRadioEnable,$FunctionRadioRemove,$FunctionSubmitButton))
            $MainWindow.Controls.Add($FunctionGroup)
        }
    )

    # Click logic for the Cancel button - basically resets the MultiModeGroup and thus, the form #
    $MMGCSVCancelButton.Add_Click(
        {
            $MultiModeGroup.Controls.Clear()
            $MMGCSVResultPathLabel.Text = "Please ensure that your CSV file has the proper fields."
            $MMGCSVResultPathLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
            $MultiModeGroup.Controls.AddRange(@($MMGCSVBrowseLabel,$MMGCSVResultPathLabel,$MMGCSVBrowseButton))
        }
    )

    ###
    ### END MMGCSV BUTTON SECTION ###
    ###

    ###
    ### MMGUaL BUTTON SECTION ###
    ###

    # Click logic for MMG UaL Check for Updates Button, which fires off F-GAU and S-E functions and compares them, then collects the differential data and stores it for use by the CheckedListBox #
    $MMGUaLUpdateButton.Add_Click(
        {
            # Rearrange the elements of the form when the button is clicked
            $MMGUaLUpdateLabel.Text = "Fetching records..."
            $MultiModeGroup.Controls.RemoveByKey("MMGUalUpdateButton")
            $MultiModeGroup.Controls.RemoveByKey("MMGUaLLapseCheckLabel")
            $MultiModeGroup.Controls.RemoveByKey("MMGUaLLapseCheckButton")
            $SCRIPT:MMGMode = "Update"
            
            # Call the Search-Excel function in A (all) mode; this should populate the $PSCOList ArrayList so that we can iterate against it
            Search-Excel -Source $FilePath -Mode A
            $MMGUaLDateStamp = (Get-Date -Format "MM/dd/yyyy")
            $MMGUaLSelectAllItem = "Select All"
            $MMGUaLCheckedList.Items.Add($MMGUaLSelectAllItem)
            ForEach ($Record in $PSCOList) {               
                $Identity = $Record.3
                $Name = $Record.2
                # Get the Notes property, store it, and then remove the property to avoid confusing the comparison function #
                $TempStatus = $Record.22
                $TempNotes = $Record.23
                $TempRow = $Record.24
                Write-Log -Level INFO -Message "$Identity - $Name - Status: $TempStatus - Notes: $TempNotes" -logfile $LogFilePath
                $Record.PSObject.Properties.Remove(22)
                $Record.PSObject.Properties.Remove(23)
                $Record.PSObject.Properties.Remove(24)
                $MultiModeGroup.Text = "Updates"
                #$MMGUaLUpdateLabel.Text = "Fetching records..." + "`r`n" + "Record found for $Name ($Identity)."
                Try {
                    $ADCheck = Fix-GetADUser -ADInfo (Get-ADUser -Identity $Identity -Properties * -ErrorAction Stop) -Compare $true
                    Start-Sleep 0.25
                    $SCRIPT:UserTermedAndDeleted = $false
                } Catch {
                    Write-Log -Message "User $Name ($Identity) no longer has a valid AD account object. Marking this record accordingly." -Level WARN -logfile $LogFilePath
                    $SCRIPT:UserTermedAndDeleted = $true
                    $ADCheck = $Record
                }
                $ADIdentity = $ADCheck.3
                # Same here - check, store, and remove the Enabled property #
                $ADEnabled = $ADCheck.29
                $ADCheck.PSObject.Properties.Remove(29)
                Write-Log -Message "Record for $ADIdentity shows that Enabled is $ADEnabled." -Level INFO -logfile $LogFilePath
                        
                # If the account is disabled, write a note in the log and mark the account #
                If ($SCRIPT:UserTermedAndDeleted) {
                    Write-Host "Change found for $Name. Record marked as Account Object DELETED."
                    $ChangeItemDel = "Change for $Name - Record marked as Term"
                    $ChangesList.Add($ChangeItemDel) | Out-Null
                    $TempStatus = "I"
                    $TempNotes = "TERM"

                } ElseIf ($ADEnabled -eq $False) {
                    Write-Log -Message "User $Name ($Identity) still has a valid AD account object, but the account is disabled. Proceeding with Update function." -Level WARN -logfile $LogFilePath
                    $SCRIPT:UserTermed = $true
                    Write-Host "Change found for $Name. Record marked as Account Disabled."
                    $ChangeItemTerm = "Change for $Name - Record marked as Account Disabled"
                    $ChangesList.Add($ChangeItemTerm) | Out-Null
                    $TempStatus = "I"
                    $TempNotes = "Account Disabled as of $MMGUaLDateStamp"

                } Else {
                    Write-Log -Message "S-E says this is for $Name ($Identity); F-GAU says $ADIdentity. Running comparison function now." -Level INFO -logfile $LogFilePath
                    $SCRIPT:UserTermed = $false
                    #$TempStatus = "A"
                }
                $Comparison = Compare-ObjectProperties -ReferenceObject $Record -DifferenceObject $ADCheck
                $Difference = $Comparison.DiffValue
                $Reference = $Comparison.RefValue
                If ($Difference -ne $null) {
                    ForEach ($CompItem in $Comparison) {
                        $CompValue = $CompItem.PropertyName
                        $CompRef = $CompItem.RefValue
                        $CompDiff = $CompItem.DiffValue
                        Write-Host "Change for $Name - Category $CompValue = $CompRef --> $CompDiff"
                        $ChangeItem = "$Identity`: Change for $Name - Category $CompValue = $CompRef --> $CompDiff"
                        $ChangesList.Add($ChangeItem) | Out-Null
                    }
                    Write-Log -Message "For $Name (aka $Identity) (aaka $ADIdentity), the difference is: `r`n $Difference --- $Reference" -Level INFO -logfile $LogFilePath
                    If ($TempStatus -eq "I") {
                        $ADCheck | Add-Member -MemberType NoteProperty -Name 22 -Value $TempStatus
                        Write-Host "$Name`: Adding $TempStatus to the Status field, which is category 22."
                    }
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 23 -Value $TempNotes
                    Write-Host "$Name`: Adding $TempNotes to the Notes field, which is category 23."
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 24 -Value $TempRow
                    Write-Host "$Name`: Adding $TempRow to the Row field, which is category 24."
                    $DifferenceList.Add($ADCheck) | Out-Null
                } ElseIf ($SCRIPT:UserTermedAndDeleted) {
                    Write-Log -Message "For $Name (aka $Identity), the difference is that they are Termed, and will be marked as such." -Level INFO -logfile $LogFilePath
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 22 -Value $TempStatus
                    Write-Host "$Name`: Adding $TempStatus to the Status field, which is category 22."
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 23 -Value $TempNotes
                    Write-Host "$Name`: Adding $TempNotes to the Notes field, which is category 23."
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 24 -Value $TempRow
                    Write-Host "$Name`: Adding $TempRow to the Row field, which is category 24."
                    $DifferenceList.Add($ADCheck) | Out-Null
                } Else {
                    Write-Log -Message "Nothing to report for $Name (aka $ADIdentity)." -Level INFO -logfile $LogFilePath
                }
                $UserChangeList = $ChangesList -match $Identity
                $MMGUaLCheckedList.Items.AddRange($UserChangeList)
                #$UserChangeList | Format-Table

            }
            [int]$CLCount = $MMGUaLCheckedList.Items.Count
            $CLCountFix = $CLCount - 1
            $MMGUaLUpdateLabel.Text = "Fetching records..." + "`r`n" + "$CLCountFix record(s) found. Please select the changes you would like to make."
            For ($i = 0; $i -lt $CLCount; $i++) {
                $MMGUaLCheckedList.SetItemChecked($i, $true)
            }
            
            
            $MultiModeGroup.Controls.AddRange(@($MMGUaLCheckedList,$MMGUaLResultLabel,$MMGUaLSubmitButton,$MMGUaLCancelButton))

            #$MMGUaLCheckedList.CheckedItems

             
        }
    )

    $MMGUaLLapseCheckButton.Add_Click(
        {
            # Rearrange the elements of the form when the button is clicked
            $MMGUaLUpdateLabel.Text = "Fetching records..."
            $MultiModeGroup.Controls.RemoveByKey("MMGUalUpdateButton")
            $MultiModeGroup.Controls.RemoveByKey("MMGUaLLapseCheckLabel")
            $MultiModeGroup.Controls.RemoveByKey("MMGUaLLapseCheckButton")
            $SCRIPT:MMGMode = "Lapse"
            
            # Call the Search-Excel function in L (lapse) mode; this should populate the $PSCOList ArrayList so that we can iterate against it
            Search-Excel -Source $FilePath -Mode L
            $MMGDate = Get-Date
            $MMGUaLDateStamp = (Get-Date -Format "MM/dd/yyyy")
            $MMGUaLFirstOfMonth = (Get-Date -Year $MMGDate.Year -Month $MMGDate.Month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0)
            $MMGUaLLapseDate = $MMGUaLFirstOfMonth.AddDays(-60)
            $MMGUaLSelectAllItem = "Select All"
            $MMGUaLCheckedList.Items.Add($MMGUaLSelectAllItem)
            ForEach ($Record in $PSCOList) {               
                
                
                ### This all needs to change; the biggest thing is that we're not doing an update here. This section should only be used to compare the LastLoginCurrent of each PSCO ###
                ### with the first day of the current month and determine if it's 60 days old or not. Once the list is pruned to the 60+ group, there will be further prunings: ###
                ### First, make sure to only include users whose Status is set to A; we don't care about ones that are already inactive. ###
                ### Second, need to make sure the AccessType isn't IO, because those are in the spreadsheet but won't ever have a LastLoginCurrent value. ###
                ### Last, anyone who has a "Y" for VPNToggle is out as well. VPN users aren't going to be lapsed in this way. ###
                
                
                $Identity = $Record.3
                $Name = $Record.2
                $LastLoginRaw = $Record.10
                If ($LastLoginRaw -eq "") {
                    $LastLogin = ""
                } Else {
                    $LastLogin = [DateTime]::ParseExact($LastLoginRaw,"ddd MMM dd HH:mm:ss yyyy",$null)
                }
                $VPN = $Record.16
                $Type = $Record.18
                $Status = $Record.22
                $Notes = $Record.23
                
                #Write-Log -Level INFO -Message "$Identity - $Name - Status: $Status - Notes: $Notes" -logfile $LogFilePath
                $MultiModeGroup.Text = "Lapsed Users"
                #$MMGUaLUpdateLabel.Text = "Fetching records..." + "`r`n" + "Record found for $Name ($Identity)."
                
                If (((($LastLogin -lt $MMGUaLLapseDate) -and ($Status -ne "I") -and ($Type -ne "IO") -and ($VPN -ne "Y")))) {
                    If ($LastLogin -ne "") {
                        Write-Log -Level INFO -logfile $LogFilePath -Message "User $Name ($Identity) - last login was $LastLogin - Notes: $Notes"
                        $ChangeItem = "User $Name ($Identity) - last login was $LastLogin - Notes: $Notes"
                    } Else {
                        Write-Log -Level INFO -logfile $LogFilePath -Message "User $Name ($Identity) - no logins recorded - Notes: $Notes"
                        $ChangeItem = "User $Name ($Identity) - no logins recorded - Notes: $Notes"
                    }
                    $ChangesList.Add($ChangeItem) | Out-Null
                    $LapseList.Add($Record) | Out-Null
                }
                
                $UserChangeList = $ChangesList -match $Identity
                $MMGUaLCheckedList.Items.AddRange($UserChangeList)
                #$UserChangeList | Format-Table

            }
            [int]$CLCount = $MMGUaLCheckedList.Items.Count
            $CLCountFix = $CLCount - 1
            $MMGUaLUpdateLabel.Text = "Fetching records..." + "`r`n" + "$CLCountFix record(s) found. Please select the changes you would like to make."
            For ($i = 0; $i -lt $CLCount; $i++) {
                $MMGUaLCheckedList.SetItemChecked($i, $true)
            }
            
            
            $MultiModeGroup.Controls.AddRange(@($MMGUaLCheckedList,$MMGUaLResultLabel,$MMGUaLSubmitButton,$MMGUaLCancelButton))

            #$MMGUaLCheckedList.CheckedItems

             
        }
    )

    # Logic for Select All button selection inside the CheckedListBox #
    $MMGUaLCheckedList.Add_Click(
        {
            If ($This.SelectedItem -eq 'Select All') {
                For ($i=1;$i -lt $MMGUaLCheckedList.Items.count;$i++) {
                    If ($MMGUaLCheckedList.GetItemChecked($i)) {
                        $MMGUaLCheckedList.SetItemchecked($i,$False)
                    } Else {
                        $MMGUaLCheckedList.SetItemChecked($i,$True)
                    }
                }
            }
        }
    )

    # Logic for Cancel changes button in MMGUaL #
    $MMGUaLCancelButton.Add_Click(
        {
            Exit(1)
        }
    )

    # Logic for Submit changes button in MMGUaL #
    $MMGUaLSubmitButton.Add_Click(
        {
            ForEach ($CItem in $MMGUaLCheckedList.CheckedItems) {
                Write-Host "Here is the change: $CItem"
                $UserChoiceList.Add($CItem) | Out-Null
            }

            #Read-Host "MMGMode is set to $SCRIPT:MMGMode."

            If ($SCRIPT:MMGMode -eq "Update") {
                ForEach ($DItem in $DifferenceList) {
                
                    ## Need to make sure we are accounting for UserTermed and UserTermedAndDeleted here; probably have to pass some kind of variable switch to Write-Excel ##
                    ## So it will know what changes to make to the record ##
                
                    $DiffUser = $DItem.2
                    $DiffName = $DItem.3
                    $ChoiceNameList = $UserChoiceList -match $DiffUser
                    $CNLCount = $ChoiceNameList.Count
                    Write-Host "The target name is $DiffName ($DiffUser). There are $CNLCount entries for this user. If there is no name, we have a problem"
                    ForEach ($CNItem in $ChoiceNameList) {
                        Switch ($CNItem) {
                            {($_ -match "Category 2") -eq $false} {$DItem.PSObject.Properties.Remove(2); Write-Log -Message "$DiffUser($DiffName) - DisplayName field removed - no changes needed." -Level INFO -logfile $LogFilePath}
                            {($_ -match "Category 3") -eq $false} {$DItem.PSObject.Properties.Remove(3); Write-Log -Message "$DiffUser($DiffName) - UserName field removed - no changes needed." -Level INFO -logfile $LogFilePath}
                            {($_ -match "Category 4") -eq $false} {$DItem.PSObject.Properties.Remove(4); Write-Log -Message "$DiffUser($DiffName) - AssociateID field removed - no changes needed." -Level INFO -logfile $LogFilePath}
                            {($_ -match "Category 7") -eq $false} {$DItem.PSObject.Properties.Remove(7); Write-Log -Message "$DiffUser($DiffName) - Dealership field removed - no changes needed." -Level INFO -logfile $LogFilePath}
                            {($_ -match "Category 8") -eq $false} {$DItem.PSObject.Properties.Remove(8); Write-Log -Message "$DiffUser($DiffName) - Department field removed - no changes needed." -Level INFO -logfile $LogFilePath}
                            {($_ -match "Category 9") -eq $false} {$DItem.PSObject.Properties.Remove(9); Write-Log -Message "$DiffUser($DiffName) - Title field removed - no changes needed." -Level INFO -logfile $LogFilePath}
                        }
                    }
                    If ($CNLCount -eq 0) {
                        Write-Log -Message "No changes written for $DiffUser ($DiffName), since they were not selected for processing." -Level INFO -logfile $LogFilePath
                    } Else {
                        Write-Excel -Source $FilePath -SearchExcelOutput $DItem
                        Write-Log -Message "Changes written for $DiffUser ($DiffName)." -Level INFO -logfile $LogFilePath
                    }

                }
            } ElseIf ($SCRIPT:MMGMode -eq "Lapse") {
                
                $LCurrentDate = (Get-Date -Format "MM/dd/yyyy")
                #Read-Host "Just making sure we're set to lapse"
                ForEach ($LItem in $LapseList) {
                    $LapseUser = $LItem.2
                    $LapseName = $LItem.3
                    $LapseMachine = $LItem.19
                    $LapseSAMName = $LapseMachine + "$"
                    $LapseGroups = $LItem.17
                    If ($LapseGroups -like "*Data Bank*") {
                        $LapseLocation = "DB"
                    } ElseIf ($LapseGroups -notlike "*VPN*") {
                        $LapseLocation = "DC"
                    }
                    $LItem.21 = $LCurrentDate
                    $LItem.22 = "I"
                    $LItem.23 = "Access Lapsed $LCurrentDate"
                    $ChoiceNameList = $UserChoiceList -match $LapseUser
                    $CNLCount = $ChoiceNameList.Count

                    Write-Host "CNLCount for $LapseName is $CNLCount"

                    If ($CNLCount -eq 0) {
                        Write-Log -Message "No changes written for $LapseUser ($LapseName), even though they qualify." -Level INFO -logfile $LogFilePath
                    } Else {
                        Read-Host "Preparing to write! Press enter to continue"
                        Write-Excel -Source $FilePath -SearchExcelOutput $LItem
                        Write-Log -Message "User $LapseUser ($LapseName) has been marked as lapsed. Proceeding to disable..." -Level INFO -logfile $LogFilePath

                        RDPRemove -TargetUser $LapseUser -TargetServer $LapseMachine -TargetSAMAccountName $LapseSAMName -TargetLocation $LapseLocation
                        Write-Log -Message "Access disabled for $LapseUser ($LapseName)." -Level INFO -logfile $LogFilePath
                    }
                }
            }
            $MMGUaLResultLabel.Text = "Processing is complete. Thanks for using the RDP Management SuperScript!"
            #$MainWindow.Controls.RemoveByKey("FunctionGroup")
            # Add the buttons to the ResultGroup #
            $ResultGroup.Controls.AddRange(@($ResultButtonRestart,$ResultButtonQuit))
            $MainWindow.Controls.Add($ResultGroup)
        }
    )



    ###
    ###
    ### FUNCTION-SUBMIT BUTTON, WHERE THINGS ARE ACTUALLY HAPPENING ###
    ###
    ###


    # Click logic for the Function Submit Button - this should fire off the functions, so it will need to process the CSV data if we're in that mode or the data from the user entry if in individual mode #
    $FunctionSubmitButton.Add_Click(
        {
            If ($ScriptMode -eq "ImportCSV") {
                $MMGCSVOutputBox.Text = ""
                ForEach ($MMGCSVEntry in $MMGCSVPSCOList) {
                    $User = $MMGCSVEntry.3
                    $Machine = $MMGCSVEntry.19
                    $SAMName = $Machine + "$"
                    $Location = $MMGCSVEntry.17
                    $NLA = $MMGCSVEntry.18
                    $FullName = $MMGCSVEntry.2
                    
                    Write-Log -Message "Preparing to set up access for $User on $Machine. NLA set to $NLA and pointing to $Location. Direction set to $SCRIPT:SMGDirection." -Level INFO -logfile $LogFilePath
                    
                    If (($FunctionRadioEnable.Checked) -and (!($FunctionRadioRemove.Checked))) {
                        If ($NLA -eq "w") { $Type = "Windows" } ElseIf ($NLA -eq "m") { $Type = "Mac" }
                        $MMGCSVOutputBox.AppendText("`r`n" + "Enabling RDP for $FullName on $Machine. Username is $User and will be set up to access $Location from $Type.")
                        RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                        $MMGCSVEntry.23 = "Access enabled $SCRIPT:CSVLastUpdated"
                        Write-Excel -Source $FilePath -SearchExcelOutput $MMGCSVEntry
                        
                        If ($FailedHosts -contains $Machine) {
                            $MMGCSVOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                        } Else {
                            $MMGCSVOutputBox.Text += "`r`n" + "$FullName has been enabled for remote access."
                        }
                    }
                    Else {
                        $MMGCSVOutputBox.AppendText("`r`n" + "Removing access for $FullName.")
                        RDPRemove -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName
                        $MMGCSVEntry.23 = "Access removed $SCRIPT:CSVLastUpdated"
                        Write-Excel -Source $FilePath -SearchExcelOutput $MMGCSVEntry
                        
                        If ($FailedHosts -contains $Machine) {
                            $MMGCSVOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                        } Else {
                            $MMGCSVOutputBox.Text += "`r`n" + "The access for $FullName has been successfully removed."
                        }

                    }

                    $MMGCSVOutputBox.Text += "`r`n" + "Processing is complete. Thanks for using the RDP Management SuperScript!"
                    $MainWindow.Controls.RemoveByKey("FunctionGroup")
                    # Add the buttons to the ResultGroup #
                    $ResultGroup.Controls.AddRange(@($ResultButtonRestart,$ResultButtonQuit))
                    $MainWindow.Controls.Add($ResultGroup)
                }
            # Need to add in code to show that the process is finished and bring the user back either to the start or to the last meaningful choice #

            } ElseIf ($ScriptMode -eq "Individual") {
                
                # Proceeding now in Individual (SMG) mode #

                # Gather variables for use with RDPEnable/RDPRemove functions #
                
                $User = $SCRIPT:SMGUserName
                $Machine = $SCRIPT:SMGDesktopName
                $SAMName = $SCRIPT:SMGDesktopSamAccountName
                $Location = $SCRIPT:SMGLocation
                If ($SCRIPT:SMGTypeSwitch -eq "") {
                    $NLA = $SCRIPT:NLASwitch
                } Else {
                    $NLA = $SCRIPT:SMGTypeSwitch
                }
                $FullName = $SCRIPT:SMGUserFullName
                
                Write-Log -Message "Preparing to set up access for $User on $Machine. NLA set to $NLA and pointing to $Location. Direction set to $SCRIPT:SMGDirection." -Level INFO -logfile $LogFilePath

                $SMGWriteObject = [PSCustomObject]@{
                    1 = $SCRIPT:SMGUserEntryNumber
                    2 = $SCRIPT:SMGUserFullName
                    3 = $SCRIPT:SMGUserName
                    4 = $SCRIPT:SMGUserEmpID
                    5 = $SCRIPT:SMGUserIDLookup
                    6 = $SCRIPT:SMGUserEarningsGroup
                    7 = $SCRIPT:SMGUserDealership
                    8 = $SCRIPT:SMGUserDepartment
                    9 = $SCRIPT:SMGUserTitle
                    10 = $SCRIPT:SMGUserLastLoginCurrent
                    11 = $SCRIPT:SMGUserLastLoginFormula
                    12 = $SCRIPT:SMGUserLastLoginHistorical
                    13 = $SCRIPT:SMGUserLast30Formula
                    14 = $SCRIPT:SMGUserLast30Current
                    15 = $SCRIPT:SMGUserRDP
                    16 = $SCRIPT:SMGUserVPN
                    17 = $SCRIPT:SMGUserMemberships
                    18 = $SCRIPT:SMGUserAccessType
                    19 = $SCRIPT:SMGUserDesktopName
                    20 = $SCRIPT:SMGUserDateAdded
                    21 = $SCRIPT:SMGUserLastUpdated
                    22 = $SCRIPT:SMGUserStatus
                    23 = $SCRIPT:SMGUserNotes
                    24 = $SCRIPT:SMGUserRow
                }

                #Read-Host "Press enter to proceed"

                If (($FunctionRadioEnable.Checked) -and (!($FunctionRadioRemove.Checked))) {
                    $SMGWriteObject.22 = "A"
                    If ($SCRIPT:SMGDirection -eq "Home") {
                        If ($NLA -eq "w") { $Type = "Windows" } ElseIf ($NLA -eq "m") { $Type = "Mac" } ElseIf ($SCRIPT:SMGTypeSwitch -ne "") { $Type = $SCRIPT:SMGTypeSwitch }
                        If (($Type -eq "Windows") -or ($Type -eq "Mac")) {
                            $SMGOutputBox.Text += "`r`n" + "Enabling RDP for $FullName on $Machine. Username is $User and will be set up to access $Location from $Type."
                            RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                            $SMGWriteObject.23 = "Access enabled $SCRIPT:SMGUserLastUpdated"
                            Write-Excel -Source $FilePath -SearchExcelOutput $SMGWriteObject

                        
                            If ($FailedHosts -contains $Machine) { 
                                $SMGOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                            } Else {
                                $SMGOutputBox.Text += "`r`n" + "$FullName has been enabled for remote desktop access."
                            }
                        } ElseIf ($Type -eq "MacToWin") {
                            $SMGOutputBox.Text += "`r`n" + "Switching access for $FullName on $Machine from Mac to Windows."
                            $NLA = "w"
                            RDPRemove -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName
                            Write-Log -Message "Mac access disabled on $Machine. Re-enabling with Windows access..." -Level INFO -logfile $LogFilePath
                            RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                            $SMGWriteObject.23 = "Moved to Windows $SCRIPT:SMGUserLastUpdated"
                            Write-Excel -Source $FilePath -SearchExcelOutput $SMGWriteObject

                        } ElseIf ($Type -eq "WinToMac") {
                            $SMGOutputBox.Text += "`r`n" + "Switching access for $FullName on $Machine from Windows to Mac."
                            $NLA = "m"
                            RDPRemove -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName
                            Write-Log -Message "Windows access disabled on $Machine. Re-enabling with Mac access..." -Level INFO -logfile $LogFilePath
                            RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                            $SMGWriteObject.23 = "Moved to Mac $SCRIPT:SMGUserLastUpdated"
                            Write-Excel -Source $FilePath -SearchExcelOutput $SMGWriteObject
                        }
                    } Else {
                        $SMGOutputBox.Text += "`r`n" + "Enabling RDP for $FullName on $Machine. Username is $User and will be set up with internal-only remote desktop access."
                        RDPEnable -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName -NLASwitch $NLA
                        $SMGWriteObject.23 = "IO Access enabled $SCRIPT:SMGUserLastUpdated"
                        Write-Excel -Source $FilePath -SearchExcelOutput $SMGWriteObject
                        
                        If ($FailedHosts -contains $Machine) { 
                            $SMGOutputBox.Text += "`r`n" + "No response was given from $Machine. Need to follow up with $FullName."
                        } Else {
                            $SMGOutputBox.Text += "`r`n" + "$FullName has been enabled for internal-only remote desktop access."
                        }
                    }
                }
                Else {
                    $SMGWriteObject.22 = "I"
                    $SMGOutputBox.Text += "`r`n" + "Removing access for $FullName."
                    RDPRemove -TargetUser $User -TargetServer $Machine -TargetLocation $Location -TargetSAMAccountName $SAMName
                    $SMGWriteObject.23 = "Access removed $SCRIPT:SMGUserLastUpdated"
                    
                    Write-Excel -Source $FilePath -SearchExcelOutput $SMGWriteObject
                    
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
                $MMGCSVBrowseLabel.Text = $DTCSVBrowseLabel
                $MMGCSVResultPathLabel.Text = $DTCSVResultPathLabel
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
