param (
    [string]$excelFilePath,        # Path to the input .xlsm file
    [string]$rooms,                # Array of room names (e.g., "61", "62")
    [string]$labs,                 # Array of lab names (e.g., "L1", "L2", "L3")
    [string]$newRoomFilePath,      # Path to save new room .xlsx file
    [string]$newLabFilePath,       # Path to save new lab .xlsx file
    [string]$newTeacherFilePath    # Path to save new teacher .xlsx file
)

# Ensure the input Excel file exists
if (-not (Test-Path $excelFilePath)) {
    Write-Host "Error: The specified Excel file does not exist."
    exit
}

# Create an Excel application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Open the original workbook (.xlsm)
Write-Host "Opening workbook: $excelFilePath"
$workbook = $excel.Workbooks.Open($excelFilePath)

# Check if the output workbook already exists and open it, otherwise create a new workbook
if (Test-Path $newTeacherFilePath) {
    $teacherWorkbook = $excel.Workbooks.Open($newTeacherFilePath)
    Write-Host "Opened existing teacher workbook: $newTeacherFilePath"
} else {
    $teacherWorkbook = $excel.Workbooks.Add()
    Write-Host "Created new teacher workbook."
}

# Function to check if a sheet name consists of only uppercase letters
function IsTeacherSheet($sheetName) {
    return $sheetName -match "^[A-Z]+$"  # Only uppercase letters, no numbers or special characters
}

# Loop through each sheet in the workbook
foreach ($sheet in $workbook.Sheets) {
    $sheetName = $sheet.Name
    Write-Host "Processing sheet: $sheetName"

    # Check if the sheet name is for a teacher
    if (IsTeacherSheet $sheetName) {
        Write-Host "Copying teacher sheet: $sheetName"
        $sheet.Copy()  # Copy to a new workbook
        $copiedSheet = $excel.ActiveSheet
        Write-Host "Teacher sheet copied: $($copiedSheet.Name)"
        
        # Clear existing sheets and move the copied sheet to the teacher workbook
        $teacherWorkbook.Sheets.Clear()
        $copiedSheet.Move($teacherWorkbook.Sheets.Item($teacherWorkbook.Sheets.Count))  # Move to teacher workbook
    }
}

# Save the updated teacher workbook in .xlsx format (overwrite the old one)
Write-Host "Saving updated teacher workbook..."
$teacherWorkbook.SaveAs($newTeacherFilePath, 51)  # 51 corresponds to the .xlsx format
Write-Host "Saved teacher workbook at: $newTeacherFilePath"

# Close the workbooks
Write-Host "Closing workbooks..."
$teacherWorkbook.Close()

# Close the original workbook and Excel application
$workbook.Close()
$excel.Quit()

# Release the COM objects to avoid memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($teacherWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Teacher sheets extracted successfully in .xlsx format!"
