param(
    [string]$ExcelFilePath,  # Path to the source Excel file (temp file)
    [string]$OutputExcelFilePath,  # Path to the output Excel file
    [string]$ImageAbovePath,  # Path to the image above (optional)
    [string]$ImageBelowPath   # Path to the image below (optional)
)

# Create an Excel Application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # Don't show the Excel application window

# Open the source workbook (temp file)
$sourceWorkbook = $excel.Workbooks.Open($ExcelFilePath)

# Create a new workbook for the output (or open an existing output file if it exists)
if (Test-Path $OutputExcelFilePath) {
    $outputWorkbook = $excel.Workbooks.Open($OutputExcelFilePath)
} else {
    $outputWorkbook = $excel.Workbooks.Add()  # If output file doesn't exist, create a new one
}

# Copy all sheets from the source workbook to the output workbook
foreach ($sheet in $sourceWorkbook.Sheets) {
    $sheet.Copy([Type]::Missing, $outputWorkbook.Sheets.Item($outputWorkbook.Sheets.Count))  # Copy each sheet
}

# Run the macro (assuming the macro is inside the source workbook and is called 'FontChange')
$excel.Application.Run("FontChange")


# Save the output workbook
$outputWorkbook.SaveAs($OutputExcelFilePath)

# Close the workbooks and Excel
$sourceWorkbook.Close($false)  # Don't save changes to the source workbook
$outputWorkbook.Close($false)  # Close the output workbook without saving it again
$excel.Quit()

# Release COM objects to avoid memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceWorkbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outputWorkbook)
