param (
    [string]$filePath,
    [string]$macroName,
    [string]$userInputLab,
    [string]$userInputLecture,
    [string]$track,  # Track (not needed for AddOns processing)
    [string]$map,    # Map (not needed for AddOns processing)
    [string]$AddOns  # Add-ons passed as a simple string
)

# Create an instance of Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $true

# Split the AddOns parameter into an array using semicolon as the separator
$AddOnsArray = $AddOns.Split(';')

# Open the workbook
$workbook = $excel.Workbooks.Open($filePath)
Write-Host "Opened workbook: $filePath"

# Loop through the AddOns array and call the FillContentInTimeSlot function
foreach ($addOn in $AddOnsArray) {
    # Split each add-on by comma to get Day, SheetName, Content, and Time
    $addOnParts = $addOn.Split(',')

    $day = $addOnParts[0]        # Day (first item)
    $sheetName = $addOnParts[1]  # Sheet Name (second item)
    $content = $addOnParts[2]    # Content (third item)
    $time = $addOnParts[3]       # Time (fourth item)

    Write-Host "Processing Add-On: Day: $day, Sheet Name: $sheetName, Content: $content, Time: $time"

    # Call the VBA macro function to fill content in the specified time slot
    $excel.Run("FillContentInTimeSlot", $content, $sheetName, $time, $day)
}

# Save the workbook after filling the content
$workbook.Save()
Write-Host "Workbook saved."

# Wait for 2 seconds
Start-Sleep -Seconds 2

# Close the workbook and Excel application
$workbook.Close()
$excel.Quit()

# Release the COM objects to avoid memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Excel resources released and macro executed successfully."
