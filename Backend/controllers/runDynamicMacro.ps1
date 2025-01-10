param (
    [string]$filePath,
    [string]$macroName,
    [string]$userInputLab,
    [string]$userInputLecture
)

# Create an instance of Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Open the workbook
$workbook = $excel.Workbooks.Open($filePath)

# Run the macro with the passed parameters
$excel.Run($macroName, $userInputLab, $userInputLecture)

# Save the workbook after running the macro (in place)
$workbook.Save()

# Close the workbook and Excel
$workbook.Close()
$excel.Quit()

# Release the COM objects to avoid memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Macro executed successfully"
