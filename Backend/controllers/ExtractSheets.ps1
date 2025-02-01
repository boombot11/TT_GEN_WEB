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

# Check if the output workbooks already exist and open them, otherwise create new workbooks
if (Test-Path $newRoomFilePath) {
    $roomWorkbook = $excel.Workbooks.Open($newRoomFilePath)
    Write-Host "Opened existing room workbook: $newRoomFilePath"
} else {
    $roomWorkbook = $excel.Workbooks.Add()
    Write-Host "Created new room workbook."
}

if (Test-Path $newLabFilePath) {
    $labWorkbook = $excel.Workbooks.Open($newLabFilePath)
    Write-Host "Opened existing lab workbook: $newLabFilePath"
} else {
    $labWorkbook = $excel.Workbooks.Add()
    Write-Host "Created new lab workbook."
}

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
    Write-Host "Processing sheet: $sheetName $rooms $labs"

    # Check if the sheet is a room sheet
    if ($rooms -contains $sheetName) {
        Write-Host "Copying room sheet: $sheetName"
        $sheet.Copy()  # Copy to a new workbook
        $copiedSheet = $excel.ActiveSheet
        Write-Host "Room sheet copied: $($copiedSheet.Name)"
        # Clear existing sheets and move the copied sheet to the room workbook
        $roomWorkbook.Sheets.Clear()
        $copiedSheet.Move($roomWorkbook.Sheets.Item($roomWorkbook.Sheets.Count))  # Move to room workbook
    }
    # Check if the sheet is a lab sheet
    elseif ($labs -contains $sheetName) {
        Write-Host "Copying lab sheet: $sheetName"
        $sheet.Copy()  # Copy to a new workbook
        $copiedSheet = $excel.ActiveSheet
        Write-Host "Lab sheet copied: $($copiedSheet.Name)"
        # Clear existing sheets and move the copied sheet to the lab workbook
        $labWorkbook.Sheets.Clear()
        $copiedSheet.Move($labWorkbook.Sheets.Item($labWorkbook.Sheets.Count))  # Move to lab workbook
    }
    # Check if the sheet name is for a teacher
 
}

# Save the updated workbooks in .xlsx format (overwrite the old ones)
Write-Host "Saving updated workbooks..."

$roomWorkbook.SaveAs($newRoomFilePath, 51)  # 51 corresponds to the .xlsx format
Write-Host "Saved room workbook at: $newRoomFilePath"

$labWorkbook.SaveAs($newLabFilePath, 51)
Write-Host "Saved lab workbook at: $newLabFilePath"

$teacherWorkbook.SaveAs($newTeacherFilePath, 51)
Write-Host "Saved teacher workbook at: $newTeacherFilePath"

# Close the workbooks
Write-Host "Closing workbooks..."
$roomWorkbook.Close()
$labWorkbook.Close()
$teacherWorkbook.Close()

# Close the original workbook and Excel application
$workbook.Close()
$excel.Quit()

# Release the COM objects to avoid memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($roomWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($labWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($teacherWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null



Start-Sleep -Seconds 2  # Add a 2-second delay

Write-Host "Sheets extracted successfully in .xlsx format!"
