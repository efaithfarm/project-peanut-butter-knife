
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
$FilePath = "C:\Users\scohamil\Sewell Automotive Companies\IT Department - RA_Sheet\Perpetual_Remote_Access.xlsx"

# Specify the path for the log file #
$LogFilePath = "C:\Users\scohamil\Sewell Automotive Companies\IT Department - RA_Sheet\Perpetual_Remote_Access.log"
$LogFileArchive = "C:\Users\scohamil\Sewell Automotive Companies\IT Department - RA_Sheet\LogArchive\"

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
        $UserPath = [ADSI]"WinNT://Sewell.com/$TargetUser,user"
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
        $UserPath = [ADSI]"WinNT://Sewell.com/$TargetUser,user"
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
    <#
    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    # Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath)
    #>
    $objExcel = Open-ExcelPackage -Path $FilePath
    #$WorkSheet = $WorkBook.Sheets.Item("Last 30")
    $WorkSheet = $objExcel.Workbook.Worksheets['Last 30']
    
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
                #$UsernameList.Add($WorkSheet.Cells["C"+$Row].Value) | Out-Null
                $TableDisplayName = $WorkSheet.Cells["B"+$Row].Value
                $TableUserName = $WorkSheet.Cells["C"+$Row].Value
                $TableAssociateID = $WorkSheet.Cells["D"+$Row].Value
                $TableDealership = $WorkSheet.Cells["G"+$Row].Value
                $TableDepartment = $WorkSheet.Cells["H"+$Row].Value
                $TableTitle = $WorkSheet.Cells["I"+$Row].Value
                $TableNotes = $WorkSheet.Cells["W"+$Row].Value
                Write-Host "$TableDisplayName, $TableuserName, $TableAssociateID, $TableDealership, $TableDepartment, $TableTitle, $TableNotes."

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
                $Row ++
            }
        Until ($TableDisplayName -eq $null)

        [int]$LastEntryNumber = $WorkSheet.Cells["A"+$LastRow].Value
        $NextEntryNumber = $LastEntryNumber + 1

        <#
        ForEach ($Username in $UsernameList) {
            $Found = $WorkSheet.Cells.Find($Username)
            $BeginAddress = $Found.Address(0,0,1,1)

            $Row = $Found.Row
        }
        #>
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
    Close-ExcelPackage $objExcel
    <#
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    Remove-Variable excel -ErrorAction SilentlyContinue
    #>
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
    <#
    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    # Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath)
    #>
    $objExcel = Open-ExcelPackage -Path $FilePath
    #$WorkSheet = $WorkBook.Sheets.Item("Last 30")
    $WorkSheet = $objExcel.Workbook.Worksheets['Last 30']
    
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

Function CreateForm {

   # Establish the main form window and define its dimensions #
    $MainWindow = New-Object System.Windows.Forms.Form
    $MainWindow.Text = "Sewell Remote Access Management App"
    $MainWindow.Width = 640
    $MainWindow.Height = 480
    $MainWindow.AutoSize = $True
    $MainWindow.Padding = 10

    # Group to hold controls for CSV and UaL modes #
    $MultiModeGroup = New-Object System.Windows.Forms.GroupBox
    $MultiModeGroup.Location = '10,10'
    $MultiModeGroup.AutoSize = $True
    $MultiModeGroup.Name = "MultiModeGroup"
    $MultiModeGroup.Text = ""
    $MultiModeGroup.Padding = 5

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

    $MultiModeGroup.Controls.AddRange(@($MMGUaLUpdateLabel,$MMGUaLLapseCheckLabel,$MMGUaLUpdateButton,$MMGUaLLapseCheckButton))
    $MainWindow.Controls.Add($MultiModeGroup)

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
                    $TempStatus = "I"
                    $TempNotes = "TERM"

                } ElseIf ($ADEnabled -eq $False) {
                    Write-Log -Message "User $Name ($Identity) still has a valid AD account object, but the account is disabled. Proceeding with Update function." -Level WARN -logfile $LogFilePath
                    $SCRIPT:UserTermed = $true
                    Write-Host "Change found for $Name. Record marked as Account Disabled."
                    $ChangeItemTerm = "Change for $Name (aka $Identity) - Record marked as Account Disabled"
                    $ChangesList.Add($ChangeItemTerm) | Out-Null
                    $TempStatus = "I"
                    $TempNotes = "Account Disabled as of $MMGUaLDateStamp"

                } Else {
                    Write-Log -Message "S-E says this is for $Name ($Identity); F-GAU says $ADIdentity. Running comparison function now." -Level INFO -logfile $LogFilePath
                    $SCRIPT:UserTermed = $false
                    #$TempStatus = "A"
                }
                $Difference = $null
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
                    $StatusCheck = $ADCheck.22
                    $NotesCheck = $ADCheck.23
                    $RowCheck = $ADCheck.24
                    Write-Host "ADCheck currently shows: 22 ($StatusCheck), 23 ($NotesCheck), and 24 ($RowCheck)."
                    $ChangeItemDel = "Change for $Name (aka $Identity) - Record marked as Term"
                    $ChangesList.Add($ChangeItemDel) #| Out-Null
                    $DifferenceList.Add($ADCheck) | Out-Null
                } Else {
                    Write-Log -Message "Nothing to report for $Name (aka $ADIdentity)." -Level INFO -logfile $LogFilePath
                }
                $UserChangeList = $ChangesList -match $Identity
                #Write-Host "$UserChangeList"
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


    If ($RestartCounter -lt 1) {
        $MainWindow.ShowDialog()
    }

}

CreateForm
