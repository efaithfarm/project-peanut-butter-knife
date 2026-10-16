$InstructionLabel = New-Object System.Windows.Forms.Label
$InstructionLabel.Location = '10,18'
$InstructionLabel.Text = "Enter the information below to process individually, or click the button to bulk process."
$InstructionLabel.AutoSize = $True
$MainWindow.Controls.Add($InstructionLabel)


$BrowseDialog = New-Object System.Windows.Forms.OpenFileDialog
$BrowseDialog.Filter = "CSV Files (*.csv)| *.csv"
$BrowseDialog.Title = "Choose CSV File"

$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Text = "Select CSV"
$BrowseButton.Location = '480,10'
$BrowseButton.Size = '120,23'
$BrowseButton.Add_Click(
    {
        $BrowseDialog.ShowDialog()
    }
)
$MainWindow.Controls.Add($BrowseButton)


$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = '480,347'
$SubmitButton.Size = '120,23'
$SubmitButton.Text = "Submit"
#$MainWindow.Controls.Add($SubmitButton)

            $MainWindow.Controls.Clear()
            $MainWindow.Controls.Add($MethodGroup)
            $MainWindow.Controls.Add($NextButton)

            
            # Add label and buttons to the SMGTypeGroup #
            $SMGTypeGroup.Controls.AddRange(@($SMGTypeLabel,$SMGTypeButtonWin,$SMGTypeButtonMac))
            
            # Add SMGTypeGroup to the SingleModeGroup #
            $SingleModeGroup.Controls.AddRange(@($SMGTypeGroup))