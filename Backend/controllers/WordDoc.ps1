param (
    [string]$ExcelFilePath,
    [string]$ImageAbovePath,
    [string]$ImageBelowPath,
    [string]$outputWordFilePath  # Corrected variable name to match the parameter
)

# Ensure necessary libraries are loaded
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false  # Initially hide Excel, but will show during copy
$Word = New-Object -ComObject Word.Application
$Word.Visible = $true
$Excel.DisplayAlerts = $false
$Word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone
$Word.Visible = $false 
$Excel.Visible = $false

# Open the Excel file
$ExcelWorkbook = $Excel.Workbooks.Open($ExcelFilePath)

# Create a new Word document
$WordDoc = $Word.Documents.Add()

# Add hardcoded text at the start of the document
$WordDoc.Content.InsertAfter("TESTTTTTTTTTTTTTTTTING EDITING")
$WordDoc.Content.InsertParagraphAfter()  # Ensure there's a paragraph break after the inserted text
# $WordDoc.Content.PasteAndFormat([Microsoft.Office.Interop.Word.WdRecoveryType]::wdFormatOriginalFormatting)

# Loop through all sheets in the Excel file
foreach ($Sheet in $ExcelWorkbook.Sheets) {
    Write-Host "Processing sheet: $($Sheet.Name)"
    
    # # Add a page break before starting with the new sheet data (to ensure it starts on a new page)
    # $WordDoc.Content.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)

    # Add the image above the page content (before the sheet data)
    # Write-Host "Inserting image above..."
    # $WordDoc.InlineShapes.AddPicture($ImageAbovePath)
    # $WordDoc.Content.InsertParagraphAfter()  # Move to next line after the image

    # Hardcoded range - Change the range as needed (e.g., "A1:G20")
    $ExcelRange = $Sheet.Range("A8:H29")  # Specify the range here

    Write-Host "Copying range from sheet: $($Sheet.Name)..."
    $rangeText = $ExcelRange.Value2  # This will give you a textual representation of the data
    Write-Host "Data from Excel: $($rangeText)"

    # Check if the Excel range has data
    if ($ExcelRange.Cells.Count -eq 0) {
        Write-Host "Excel range is empty. Skipping sheet: $($Sheet.Name)"
    } else {
        Write-Host "Excel range contains data. Copying and pasting into Word..."

         $Excel.Visible = $true  # Hide Excel after copying
        # Directly copy the range data from Excel (no need to select)
        $ExcelRange.Copy()

        # Small delay to ensure the clipboard is updated
        Start-Sleep -Seconds 3  # Increased delay to ensure content is copied

        # Paste it into Word (as a table)
        $WordDoc.Content.Paste()
        Write-Host "Pasted content from sheet: $($Sheet.Name)"
        
        # Hide Excel again after copying
        $Excel.Visible = $false  # Hide Excel after copying
    }

    # # Insert a paragraph break after pasting the content
    # $WordDoc.Content.InsertParagraphAfter()

    # # Add the image below the content (after the sheet data)
    # Write-Host "Inserting image below..."
    # # $WordDoc.InlineShapes.AddPicture($ImageBelowPath)
    # # $WordDoc.Content.InsertParagraphAfter()  # Move to the next line after the image

    # # Optional: Add another page break after each sheet if desired
    # $WordDoc.Content.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
}

# Save the Word document using the provided output path
Write-Host "Saving the Word document..."
$WordDoc.SaveAs([ref] $outputWordFilePath)
Start-Sleep -Seconds 2 
# Close the Word document
$WordDoc.Close($true)

Start-Sleep -Seconds 3  # Adjust the sleep time if needed

# Close Excel and Word
$ExcelWorkbook.Close()

# Quit the Word application
$Word.Quit()

Start-Sleep -Seconds 2  # Adjust the sleep time if needed

# Clean up COM objects to prevent memory leaks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)

Write-Host "COM objects released and Word document saved successfully."
