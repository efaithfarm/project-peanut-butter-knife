

            # First thing after the UserConfirm button is to know if $SMGUserFoundOnSheet is YEP, and if so, change the way this Group looks #
            # It could use the TextBox and populate it with a PC name it pulls from the sheet #
            # And it would give the user two buttons, a Yes or No, asking if this is the PC the user wants to proceed with #
            # If No is chosen, the TextBox becomes cleared and writable and the script picks up at this point #
            # Otherwise, the script marks DesktopSuccess as YEP and moves forward from there #
        

            ## Need to pull in information from the sheet if we're doing a reinstatement, and ask if it is still current. Unfortunately that requires an outside-the-script tool (FindUser) ##
            ## If the desktop is current (reinstatement), it's probably still important to run the Get-ADComputer so the function can run against it ##
            ## If this is for a new user, need to store the desktop name and update the sheet with that information ##
            ## Would be nice to be able to use the script to update the sheet even if you're not making changes to anyone's access, like if you just wanted to enter someone's desktop information or Associate ID or something ##

                                    # Need to add the code here to write to Excel #
                        # Use variables collected through the script to create a new PSCustomObject to pass as a parameter to Write-Excel #
                        # Check to see if SMGNewUser is set to YEP; if it is, use the SMGNextRow parameter to know where to write #

                        <#
# First, set the text box to Read-only until further action is taken #
#$SMGUserTextBox.ReadOnly = $True
        
# Next, check the Perp sheet to see if the user is already present; if not, use Get-AdUser #
# Require input :P #

$SCRIPT:SMGUserNameEntry = $SearchText
$SCRIPT:SMGUserFoundOnSheet = ""
#$SMGUserResultLabel.Text = "Working..."

# Run Fix-GetADUser to pull some preliminary data before you even check the sheet #
$SCRIPT:FixedADInfo = Fix-GetADUser -ADInfo (Get-ADUser -Identity $SCRIPT:SMGUserNameEntry -Properties *)

# Use the Search-Excel function defined above to get information about the user based on their username, assuming they have a record in the Perp sheet #
Try {
    $SCRIPT:SMGUser = Search-Excel -SearchText $SCRIPT:SMGUserNameEntry -Source $FilePath
    $SCRIPT:SMGUserFoundOnSheet = "YEP"
} Catch {
                    
    # If it doesn't work, leave the name displayed and ask if the user wants to check the spelling - or if this is just a new user who wouldn't have a record in the sheet #
    <#$SMGUserResultLabel.Text =#> <#Write-Host "Well, that record wasn't found. Is there a spelling error, or is this a new remote access user?"
    Start-Sleep 2
    $SCRIPT:SMGUserFoundOnSheet = "NOPE"

    # Populate the box with buttons labeled 'New User' and 'Try Again' #
    $SMGUserButtonCancel.Text = "Try Again"
    <#$SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
    $SMGUserGroup.Controls.AddRange(@($SMGUserButtonNewUser,$SMGUserButtonCancel))
}#>

## Pulling out code that used the Get-AdUser cmdlet to find parameters and putting it in the code for the SMGUserButtonNewUser button ##
<#
# Add code to extrapolate and display the parameters in the sheet for the targeted user #
If ($SCRIPT:SMGUserFoundOnSheet -eq "YEP") {
                    
    # Assign the information from the search to variables #
    $SCRIPT:SMGUserRow = $SCRIPT:SMGUser.24
    $SCRIPT:SMGUserNextRow = $SCRIPT:SMGUser.26
    $SCRIPT:SMGUserEntryNumber = $SCRIPT:SMGUser.1
    $SCRIPT:SMGUserFullName = $SCRIPT:FixedADInfo.2
    $SCRIPT:SMGUserEmpID = $SCRIPT:FixedADInfo.4
    $SCRIPT:SMGUserIDLookup = $SCRIPT:SMGUser.5
    $SCRIPT:SMGUserEarningsGroup = $SCRIPT:SMGUser.6
    $SCRIPT:SMGUserDealership = $SCRIPT:FixedADInfo.7
    $SCRIPT:SMGUserDepartment = $SCRIPT:FixedADInfo.8
    $SCRIPT:SMGUserTitle = $SCRIPT:FixedADInfo.9
    $SCRIPT:SMGUserLastLoginCurrent = $SCRIPT:SMGUser.10
    $SCRIPT:SMGUserLastLoginFormula = $SCRIPT:SMGUser.11
    $SCRIPT:SMGUserLastLoginHistorical = $SCRIPT:SMGUser.12
    $SCRIPT:SMGUserLast30Formula = $SCRIPT:SMGUser.13
    $SCRIPT:SMGUserLast30Current = $SCRIPT:SMGUser.14
    $SCRIPT:SMGUserRDP = $SCRIPT:SMGUser.15
    $SCRIPT:SMGUserVPN = $SCRIPT:SMGUser.16
    $SCRIPT:SMGUserMemberships = $SCRIPT:SMGUser.17
    $SCRIPT:SMGUserAccessType = $SCRIPT:SMGUser.18
    $SCRIPT:SMGUserDesktopName = $SCRIPT:SMGUser.19
    $SCRIPT:SMGUserDateAdded = $SCRIPT:SMGUser.20
    $SCRIPT:SMGUserLastUpdated = (Get-Date -Format "MM/DD/YYYY")
    $SCRIPT:SMGUserStatus = $SCRIPT:SMGUser.22
    $SCRIPT:SMGUserNotes = $SCRIPT:SMGUser.23
    <#$SMGUserResultLabel.Text =#> Write-Host "$SCRIPT:SMGUserFullName was found on row $SCRIPT:SMGUserRow." + "`r`n" + "Their title is $SCRIPT:SMGUserTitle, employee ID $SCRIPT:SMGUserEmpID, with $SCRIPT:SMGUserDepartment at $SCRIPT:SMGUserDealership." + "`r`n" + "Is this correct?"
    <#$SMGUserGroup.Controls.RemoveByKey("SMGUserButtonVerify")
    $SMGUserGroup.Controls.AddRange(@($SMGUserButtonConfirm,$SMGUserButtonCancel))
}
#>
<#} Else {
    $SMGUserResultLabel.Text = "Sorry, input is required here. Please try again!"
    Start-Sleep 2
    $SMGUserTextBox.ReadOnly = $False
}#>



 ElseIf ($OutputItemColumnNumber -eq "26") {
            Continue
        } 



        (($ADInfo.Title -eq "CustomerCareRepresentativ") -or ($ADInfo.Title -eq "Manager") -or ($ADInfo.Title -eq "Director"))




# Establish the column numbers for use when writing back #
$ColumnEntryNumber = 1
$ColumnDisplayName = 2
$ColumnUserName = 3
$ColumnAssociateID = 4
$ColumnIDLookupValue = 5
$ColumnEarningsGroup = 6
$ColumnDealership = 7
$ColumnDepartment = 8
$ColumnTitle = 9
$ColumnLastLoginCurrent = 10
$ColumnLastLoginFormula = 11
$ColumnLastLoginHistorical = 12
$ColumnLast30Formula = 13
$ColumnLast30Current = 14
$ColumnRDPToggle = 15
$ColumnVPNToggle = 16
$ColumnGroups = 17
$ColumnAccessType = 18
$ColumnDesktopName = 19
$ColumnDateAdded = 20
$ColumnDateLastUpdated = 21
$ColumnStatus = 22
$ColumnNotes = 23



        Do {

            # Looks for subsequent items after the first
            $Found = $WorkSheet.Cells.FindNext($Found)
            $Address = $Found.Address(0,0,1,1)
            If ($Address -eq $BeginAddress) {
                Break
            }

            #Must match up with listing above
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

            #Displays it the second to infinity times it is found. This section must match the section above
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
            }

        } Until ($false)




# Function to get a list of usernames from the accompanying Excel spreadsheet #
Function Pull-ExcelUsernames {
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
        [string]$Source
    )

    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    # Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath)

    $WorkSheet = $WorkBook.Sheets.Item("Last 30")


    <#
    Write-Log -Message "No matches found in the $($Worksheet.Name) tab for $SearchText." -Level INFO -logfile $LogFilePath
    $TableFound = "NOPE" #>

    # Get rid of the Excel object #

    #Write-Log -Message "..." -Level INFO -logfile $LogFilePath
    $objExcel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    Remove-Variable excel -ErrorAction SilentlyContinue
}


# Function to get a list of usernames from the accompanying Excel spreadsheet #
Function Pull-ExcelUsernames {
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
        [string]$Source
    )

    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    # Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath)

    $WorkSheet = $WorkBook.Sheets.Item("Last 30")


    <#
    Write-Log -Message "No matches found in the $($Worksheet.Name) tab for $SearchText." -Level INFO -logfile $LogFilePath
    $TableFound = "NOPE" #>

    $SCRIPT:UsernameList = [System.Collections.ArrayList]@()

    $NextRow = ($WorkSheet.UsedRange.rows.count) + 1
    [int]$Row = 4

    Do
        {
            $UsernameList.Add($WorkSheet.Cells.Item($Row,3).Text)
            $Row ++
        }
    Until ($Row -eq ($NextRow))

    # Get rid of the Excel object #

    #Write-Log -Message "..." -Level INFO -logfile $LogFilePath
    $objExcel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    Remove-Variable excel -ErrorAction SilentlyContinue
}



#$UserInfo = Search-Excel -Source $FilePath -SearchText $SearchText #| Select-Object -Property DisplayName, Username, LastLoginCurrent, Last30Current, Status, DateAdded, LastUpdated, Notes | Format-Table

#$GetADUserOutput = Get-ADUser -Identity $SearchText -Properties * #| Select-Object -Property DisplayName, SamAccountName, EmployeeID, Company, Department, Title

#$FixedADInfo = Fix-GetADUser -ADInfo (Get-ADUser -Identity $SearchText -Properties *)

#Write-Excel -Source $FilePath -SearchExcelOutput $UserInfo

#Pull-ExcelUsernames -Source $FilePath
<#
$SCRIPT:UsernameList | ForEach-Object {
    Write-Host "The current username is $_."
}
#>
#Start-Sleep -Seconds 2
<#Search-Excel -Source $FilePath -All $true
ForEach ($Record in $PSCOList) {
    $Identity = $Record.3
    $Notes = $Record.23
    If ($Notes -eq "TERM") {
        Write-Host "$Identity is no longer with Sewell - the Notes field says $Notes."
        Try {
            $ADCheck = Fix-GetADUser -ADInfo (Get-ADUser -Identity $Identity -Properties * -ErrorAction Stop)
            $ADIdentity = $ADCheck.3
            Write-Host "Well, something strange is happening - $ADIdentity is the current user, according to F-GAU."
        } Catch {
            Write-Host "As expected - $Identity's AD record is missing."
        }
    }
    
    Write-Host "$Identity is the current user, according to Search-Excel."
    $ADCheck = Fix-GetADUser -ADInfo (Get-ADUser -Identity $Identity -Properties *)
    $ADIdentity = $ADCheck.3
    Write-Host "$ADIdentity is the current user, according to F-GAU."
    $Comparison = Compare-ObjectProperties -ReferenceObject $Record -DifferenceObject $ADCheck
    $Difference = $Comparison.DiffValue
    Write-Host "For $Identity (aka $ADIdentity), the difference (if any) is: $Difference"
    #Write-Host "$Differences.SideIndicator is what resulted for $DifferencesObjectUsername"
    
}#>


                    <#
                        Try {$ADCheck | Add-Member -MemberType NoteProperty -Name 22 -Value $TempStatus} Catch {Read-Host "$Name`: $Record.2 is still showing as the Name. This is jacked up"}
                        Write-Host "$Name`: Adding $TempStatus to the Status field, which is category 22."
                    
                        
                    
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 23 -Value $TempNotes
                    Write-Host "$Name`: Adding $TempNotes to the Notes field, which is category 23."
                    $ADCheck | Add-Member -MemberType NoteProperty -Name 24 -Value $TempRow
                    Write-Host "$Name`: Adding $TempRow to the Row field, which is category 24."
                    $DifferenceList.Add($ADCheck) | Out-Null#>


                <#Try {
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
                    $TempStatus = "A"
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
                }#>