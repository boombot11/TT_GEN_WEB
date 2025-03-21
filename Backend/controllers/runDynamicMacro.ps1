param (
    [string]$filePath,
    [string]$macroName,
    [string]$userInputLab,
    [string]$userInputLecture,
    [string]$track,  # Track is passed as a JSON string
    [string]$map   # Map is passed as a JSON string

 
)

# Create an instance of Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $true


# Deserialize AddOnEvents JSON

# Debugging
Write-Host "INPUTS ::"
Write-Host $track
Write-Host $map

$workbook = $excel.Workbooks.Open($filePath)
Write-Host "IN DYNAMIC MACRO OPENED WORKBOOK"

# Run the macro with the modified parameters
try {
    # Run the macro with user inputs, the converted Dictionaries, and the AddOnEvents
    $excel.Run($macroName, $userInputLab, $userInputLecture, $track, $map)
} catch {
    Write-Host "Error running macro: $_"
}

Write-Host "IN DYNAMIC MACRO RUN MACRO"

# Save the workbook after running the macro
$workbook.Save()
Write-Host "IN DYNAMIC MACRO SAVED WORKBOOK"

# Close the workbook and Excel
$workbook.Close()
$excel.Quit()

# Release the COM objects to avoid memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Macro executed successfully"
