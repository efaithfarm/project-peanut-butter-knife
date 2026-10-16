

$ExcelObject = New-Object -ComObject Excel.Application
$ExcelObject.Visible = $true

$ExcelWorkBook = $ExcelObject.Workbooks.Open("\\srvrdc42\filestorage\Documentation\Policies and Procedures\Remote Access\Perpetual_Remote_Access.xlsx")
#$ExcelObject.Workbooks.CheckOut("https://sewellautomotive.sharepoint.com/:x:/r/sites/ServersOnly/Shared%20Documents/General/TRA_CSVImport_03092021.xlsx?d=w42c2ccc6b3c2438f913619a72083df90&csf=1&web=1&e=db4fwe")

$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(5)
$ExcelCell = $ExcelWorkSheet.Cells.Item(11,3)

Write-host $ExcelCell.Text



$ExcelObject.Quit()

#$ExcelWorkBook.Save()
#$ExcelWorkBook.CheckIn()