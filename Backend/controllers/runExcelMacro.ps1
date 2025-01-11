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
Start-Sleep -Seconds 2  # Add a 2-second delay