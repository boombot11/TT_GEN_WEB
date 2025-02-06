param(
    [string]$ExcelFilePath,
    [string]$ImageAbovePath,
    [string]$ImageBelowPath
)

# Create an Excel Application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # Don't show the Excel application window

# Open the workbook
$workbook = $excel.Workbooks.Open($ExcelFilePath)

# Run the macro (assuming the macro is inside the workbook and is called 'SetFontSizeForNonInitialSheets')
$excel.Application.Run("SetFontSizeForNonInitialSheets")

# Add images if paths are provided
if ($ImageAbovePath) {
    $worksheet = $workbook.Sheets.Item(1)  # Assuming image is added to the first sheet
    $worksheet.Pictures().Insert($ImageAbovePath)
}

if ($ImageBelowPath) {
    $worksheet = $workbook.Sheets.Item(1)
    $worksheet.Pictures().Insert($ImageBelowPath)
}

# Save the workbook
$workbook.Save()

# Close the workbook and Excel
$workbook.Close()
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
