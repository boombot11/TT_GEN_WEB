param (
    [string]$ExcelFilePath,
    [string]$ImageAbovePath,
    [string]$ImageBelowPath,
    [string]$outputWordFilePath  # Corrected variable name to match the parameter
)

# Function to pause execution for a short period and retry the operation
function WaitForExcelAndWord {
    param (
        [Parameter(Mandatory=$true)]
        [object]$application
    )

    $attempt = 0
    $maxAttempts = 5

    # Retry mechanism to wait for Excel or Word to be ready
    while ($attempt -lt $maxAttempts) {
        try {
            $application.Visible = $false
            Start-Sleep -Seconds 1  # Adding sleep time to wait for app to be ready
            return $true
        }
        catch {
            $attempt++
            Write-Host "Attempt $attempt failed, retrying..."
            Start-Sleep -Seconds 2  # Retry after a small delay
        }
    }
    Write-Host "Failed to initialize Excel/Word after $maxAttempts attempts."
    return $false
}

# Ensure necessary libraries are loaded
$Excel = New-Object -ComObject Excel.Application
$Word = New-Object -ComObject Word.Application

$Excel.DisplayAlerts = $false
$Word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone

# Make sure Excel and Word are ready
if (-not (WaitForExcelAndWord -application $Excel)) {
    Write-Host "Excel is not ready. Exiting script."
    exit
}
if (-not (WaitForExcelAndWord -application $Word)) {
    Write-Host "Word is not ready. Exiting script."
    exit
}

$Word.Visible = $false
$Excel.Visible = $false

# Open the Excel file
$ExcelWorkbook = $Excel.Workbooks.Open($ExcelFilePath)

# Create a new Word document
$WordDoc = $Word.Documents.Add()

# Loop through all sheets in the Excel file
foreach ($Sheet in $ExcelWorkbook.Sheets) {
    # Add a page break before starting with the new sheet data (to ensure it starts on a new page)
    $WordDoc.Content.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)

    # Add the image above the page content (before the sheet data)
    $WordDoc.InlineShapes.AddPicture($ImageAbovePath)
    $WordDoc.Content.InsertParagraphAfter()  # Move to next line after the image

    # Copy sheet content to Word (copied as text or table)
    $ExcelRange = $Sheet.UsedRange
    $ExcelRange.Copy()

    # Paste it into Word (as a table)
    $WordDoc.Content.Paste()

    # Insert a paragraph break after pasting the content
    $WordDoc.Content.InsertParagraphAfter()

    # Add the image below the content (after the sheet data)
    $WordDoc.InlineShapes.AddPicture($ImageBelowPath)
    $WordDoc.Content.InsertParagraphAfter()  # Move to the next line after the image

    # Optional: Add another page break after each sheet if desired
    # This will place each sheet on a separate page with images above and below.
    $WordDoc.Content.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
}

# Save the Word document using the provided output path
$WordDoc.SaveAs([ref] $outputWordFilePath)

$WordDoc.Close($true)
Start-Sleep -Seconds 2  # Adjust the sleep time if needed

# Close Excel and Word
$ExcelWorkbook.Close()

$Word.Quit()
Start-Sleep -Seconds 2  # Adjust the sleep time if needed

# Clean up COM objects to prevent memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
