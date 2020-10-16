# Create Local User from Excelfile 
Password = "xxx" | ConvertTo-SecureString -AsPlainText -Force 
$column = 2
$row = 1
do {
    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Open("C:\Public\XXX.xlsx") #Location Excelfile
    $Sheet = $Workbook.Sheets.Item(1)
    $Username = $sheet.cells.item($column, $row).Text
    $Username 
    New-LocalUser -Name $Username -Password $Password 
    $column++
} while ($Username)
