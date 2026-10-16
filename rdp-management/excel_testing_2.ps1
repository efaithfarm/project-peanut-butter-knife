$FilePath = "C:\Users\scohamil\Sewell Automotive Companies\IT Department - RA_Sheet\Perpetual_Remote_Access.xlsx"

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$eWorkBook = $Excel.Workbooks.Open($FilePath)

$Last30 = $eWorkBook.Sheets.Item("Last 30")

$Range = ($Last30.UsedRange.Rows).Count
<#
$Last30 | Get-Member

Write-Host $eWorkSheets

$Value = $Last30.Rows.Item(20).Columns.Item(20).Text
Write-Host $Value
#>

