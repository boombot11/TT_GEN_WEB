param (
    [string]$ExcelFilePath,
    [string]$ImageAbovePath,
    [string]$ImageBelowPath,
    [string]$outputWordFilePath
)

# Create Word and Excel application objects
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false
$wordDoc = $wordApp.Documents.Add()

$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false
$excelApp.DisplayAlerts = $false
$workbook = $excelApp.Workbooks.Open($ExcelFilePath)

# Loop over each sheet in the Excel workbook
foreach ($sheet in $workbook.Sheets) {

    # Get a Range representing the end of the document
    $range = $wordDoc.Content
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    
   # Insert a page break if this is not the first page
    if ($wordDoc.Content.Paragraphs.Count -gt 1) {
        $range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
        $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    } else {
        # First page case: Ensure there's a proper paragraph to anchor the image
        $range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
        $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    }
    
    # Insert a paragraph marker (this paragraph will act as the anchor for subsequent shapes)
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    
    # ---------------------------
    # Add the first image (Image Above)
    # ---------------------------
    # Use the current range as the anchor
    $imageAbove = $wordDoc.Shapes.AddPicture(
        $ImageAbovePath, 
        $false,    # LinkToFile
        $true,     # SaveWithDocument
        0,         # Left (0 means use the anchor’s position)
        0,         # Top (0 means use the anchor’s position)
        26.7 * 28.35,  # Width in points for landscape (adjusted)
        4 * 28.35,      # Height in points (adjusted)
        $range    # Anchor
    )
    $imageAbove.WrapFormat.Type = 3  # wdWrapFront
    $imageAbove.LockAnchor = $true

    # Insert some extra paragraphs to create a gap
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

    # ---------------------------
    # Insert a TextBox for the Excel table
    # ---------------------------
    # Create the text box (again, anchor it to the current range)
    $textBox = $wordDoc.Shapes.AddTextbox(
        1,          # Orientation (msoTextOrientationHorizontal)
        0, 0,       # Left, Top (we’ll position later)
        780, 270,   # Width, Height in points (adjusted for landscape)
        $range      # Anchor
    )
    # Center the text box horizontally
    $pageWidth = $wordApp.ActiveDocument.PageSetup.PageWidth
    $textBox.Left = ($pageWidth - $textBox.Width) / 2
    $textBox.WrapFormat.Type = 0  # wdWrapNone

    # Optionally, adjust vertical position relative to the image:
    $textBox.Top = $imageAbove.Top + $imageAbove.Height + 10

    # ---------------------------
    # Copy the Excel table and paste into the text box
    # ---------------------------
    # (Make sure the range in Excel matches what you need)
$excelSheet = $sheet
$rangeExcel = $excelSheet.Range("A8:I30")

# Check if the range is empty
if ($rangeExcel.Cells.Count -eq 0) {
    Write-Host "Error: The selected range is empty."
    return
}

# Try to remove borders and copy the range as a picture
try {
    # Remove all borders from the range
    $rangeExcel.Borders([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).LineStyle = -4142
    $rangeExcel.Borders([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).LineStyle = -4142
    $rangeExcel.Borders([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).LineStyle = -4142
    $rangeExcel.Borders([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).LineStyle = -4142
    $rangeExcel.Borders([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideHorizontal).LineStyle = -4142
    $rangeExcel.Borders([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).LineStyle = -4142

    # Now, copy the range as a picture (without borders)
    $rangeExcel.CopyPicture(
        [Microsoft.Office.Interop.Excel.XlPictureAppearance]::xlScreen, 
        [Microsoft.Office.Interop.Excel.XlCopyPictureFormat]::xlPicture
    )
} catch {
    Write-Host "Error: Failed to copy range as picture. $_"
}

# Paste the copied picture into the text box
$textBox.TextFrame.TextRange.Paste()
# $imageInTextBox = $textBox.Shapes.Item(1)

# # Resize the image by adjusting its Width and Height properties
# $newWidth = $imageInTextBox.Width * 1.5  # Increase the width by 1.5 times
# $imageInTextBox.Width = $newWidth

# # Optional: Adjust the height to maintain aspect ratio, if needed
# $imageInTextBox.Height = $imageInTextBox.Height * 1

# Remove the border of the text box
$textBox.Line.Visible = $false
    # ---------------------------
    $offsetX = 2 * 28.35
    $textBox.Left = $textBox.Left + $offsetX

    # ---------------------------
    # Insert additional paragraphs to create space after the text box
    # ---------------------------
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

    # ---------------------------
    # Add the second image (Image Below)
    # ---------------------------
$rangeExcel2 = $sheet.Range("A34:H39")
if ($rangeExcel2.Cells.Count -eq 0) {
    Write-Host "Error: The selected range is empty."
    return
}

# Try to copy the range as a picture
try {
    $rangeExcel2.CopyPicture(
        [Microsoft.Office.Interop.Excel.XlPictureAppearance]::xlScreen, 
        [Microsoft.Office.Interop.Excel.XlCopyPictureFormat]::xlPicture
    )
} catch {
    Write-Host "Error: Failed to copy range as picture. $_"
}

# Create a new text box to hold the picture
$textBox2 = $wordDoc.Shapes.AddTextbox(
    1,          # Orientation (msoTextOrientationHorizontal)
    0, 0,       # Left, Top (we’ll position later)
    26.7 * 28.35,  # Width in points (same as ImageBelow)
    4 * 28.35,      # Height in points (same as ImageBelow)
    $range      # Anchor to the current range in Word
)

# Center the text box horizontally (same logic as before)
$pageWidth = $wordApp.ActiveDocument.PageSetup.PageWidth
$textBox2.Left = ($pageWidth - $textBox2.Width) / 2
$textBox2.WrapFormat.Type = 0  # wdWrapNone

# Position the second text box below the first one (imageAbove)
$textBox2.Top = $textBox.Top + $textBox.Height + 10

# Paste the Excel range as a picture into the text box
$textBox2.TextFrame.TextRange.Paste()

# Remove the border of the text box
$textBox2.Line.Visible = $false
}

# Save and close documents
$wordDoc.SaveAs([ref]$outputWordFilePath)
Start-Sleep -Seconds 2

$workbook.Close()
$excelApp.Quit()
$wordDoc.Close()
$wordApp.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDoc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
