# Specify the path of the excel file
$FilePath = "C:\Users\scohamil\Sewell Automotive Companies\IT Department - RA_Sheet\Perpetual_Remote_Access.xlsx"

# Specify the path for the log file #
$LogFilePath = "C:\Users\scohamil\Sewell Automotive Companies\IT Department - RA_Sheet\Perpetual_Remote_Access_testing.log"
$LogFileArchive = "C:\Users\scohamil\Sewell Automotive Companies\IT Department - RA_Sheet\LogArchive\"

# What you are searching for
#$SearchText = Read-Host "Please enter a username"

$UsernameList = [System.Collections.ArrayList]@()
$PSCOList = [System.Collections.ArrayList]@()
$DifferenceList = [System.Collections.ArrayList]@()
$ChangesList = [System.Collections.ArrayList]@()
[bool]$SCRIPT:UserTermed = $false
[bool]$SCRIPT:UserTermedAndDeleted = $false 

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
        $ArchivePath = $LogFileArchive + "PRA_Log_Test_" + $FileNameDateFormat + ".log"
        If ($LogFileSize -ge 1.0) {
            Move-Item -Path $logfile -Destination $ArchivePath
        }

        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

# Function to find user information in the accompanying Excel spreadsheet #
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
        [bool]$All,
        [string]$SearchText
    )

    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    # Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath)

    $WorkSheet = $WorkBook.Sheets.Item("Last 30")
    If (!($All)) {
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
            $TableStatus = $WorkSheet.Rows.Item($Row).Columns.Item(22).Text
            $TableNotes = $WorkSheet.Rows.Item($Row).Columns.Item(23).Text

            $PSCO = [pscustomobject]@{
                2 = $TableDisplayName
                3 = $TableUserName
                4 = $TableAssociateID
                7 = $TableDealership
                8 = $TableDepartment
                9 = $TableTitle
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
Write-Host "The Target Row is $TargetRow"
$SearchExcelOutput.PSObject.Properties | ForEach-Object {
    [int]$OutputItemColumnNumber = $_.Name
    Write-Host "In column number $OutputItemColumnNumber..."
    $OutputItemValue = $_.Value
    Write-Host "We find the value of...$OutputItemValue"
    If ($OutputItemColumnNumber -lt 24) {
        If (($OutputItemColumnNumber -eq "22") -and ($OutputItemValue -eq "I")) {
            Write-Host "Changing the style at ($TargetRow,$OutputItemColumnNumber)."
            $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Style = 'Bad'
        } ElseIf (($OutputItemColumnNumber -eq "22") -and ($OutputItemValue -eq "A")) {
            Write-Host "Changing the style at ($TargetRow,$OutputItemColumnNumber)."
            $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Style = 'Good'
        } ElseIf ($OutputItemValue -eq "") {
            #Continue
            $CurrentValue = $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Text
            Write-Host "Nothing to do here at ($TargetRow,$OutputItemColumnNumber)."
        } Else {
            $CurrentValue = $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Text
            If (($OutputItemValue -eq "") -and ($CurrentValue -ne "")) {
                Write-Host "No overwriting necessary, we've already got a value here."
            } Else {
                Write-Host "Entering $OutputItemValue at ($TargetRow,$OutputItemColumnNumber) to replace $CurrentValue."
                $WorkSheet.Cells.Item($TargetRow,$OutputItemColumnNumber).Value = $OutputItemValue
            }
        }
    } Else {
        Write-Host "Skipping this column - it should not be valued. That kinda sounds mean."
    }
}

Write-Host ""
$objExcel.DisplayAlerts = $false
$WorkBook.SaveAs($FilePath)
Read-Host "Ready to quit? Just hit enter."
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
        [parameter(Mandatory)]
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

# Call the Search-Excel function in $All mode; this should populate the $PSCOList ArrayList so that we can iterate against it
            Search-Excel -Source $FilePath -All $true
            $MMGUaLDateStamp = (Get-Date -Format "MM/dd/yyyy")
            #$MMGUaLSelectAllItem = "Select All"
            #$MMGUaLCheckedList.Items.Add($MMGUaLSelectAllItem)
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
                #$MultiModeGroup.Text = "Updates"
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
                        $ChangeItem = "Change for $Name - Category $CompValue = $CompRef --> $CompDiff"
                        $ChangesList.Add($ChangeItem) | Out-Null
                    }
                    Write-Log -Message "For $Name (aka $Identity) (aaka $ADIdentity), the difference is: `r`n $Difference --- $Reference" -Level INFO -logfile $LogFilePath
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 22 -Value $TempStatus
                    Write-Host "$Name`: Adding $TempStatus to the Status field, which is category 22."
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
                $UserChangeList = $ChangesList -match $Name
                #$MMGUaLCheckedList.Items.AddRange($UserChangeList)
                #$UserChangeList | Format-Table

            }
            $DifferenceList | Format-Table
            
Kill -Processname EXCEL