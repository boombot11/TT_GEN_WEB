param (
    [string]$excelFilePath,
    [string]$macroName
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($excelFilePath)
$excel.Run($macroName)
$workbook.Save()
$workbook.Close()
$excel.Quit()
