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
# $wordDoc.PageSetup.Orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape
$wordDoc.PageSetup.TopMargin = 36  # Adjust top margin to create space
$wordDoc.PageSetup.BottomMargin = 36  # Adjust bottom margin to create space
$wordDoc.PageSetup.LeftMargin = 36  # Adjust left margin to create space
$wordDoc.PageSetup.RightMargin = 36  # Adjust right margin to create space

$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false
$excelApp.DisplayAlerts = $false
$workbook = $excelApp.Workbooks.Open($ExcelFilePath)

# Get the page dimensions in points (8.5 x 11 inches in landscape = 792 x 612 points)
$pageWidth = $wordDoc.PageSetup.PageWidth
$pageHeight = $wordDoc.PageSetup.PageHeight

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
    # Add the first image (Image Above) dynamically resized
    # ---------------------------
    $imageAbove = $wordDoc.Shapes.AddPicture(
        $ImageAbovePath, 
        $false,    # LinkToFile
        $true, 
  
        0,         # Left (0 means use the anchor’s position)
        0,         # Top (0 means use the anchor’s position)
                 26.7 * 28.35,  # Width in points for landscape (adjusted)
        4 * 28.35,      # Height in points (adjusted)    # SaveWithDocument
        $range     # Anchor
    )
    
    # Fit image to available width in landscape (with some padding for margins)
    $imageAbove.LockAspectRatio = $true
    $imageAbove.Width = $pageWidth - 72  # Subtracting margins (1 inch = 72 points)
    $imageAbove.Top = 0

    $imageAbove.WrapFormat.Type = 3  # wdWrapFront
    $imageAbove.LockAnchor = $true

    # Insert some extra paragraphs to create a gap
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

    # ---------------------------
    # Insert the Excel table as a Word table
    # ---------------------------
    # Define the Excel range to be copied
    $excelSheet = $sheet
    $rangeExcel = $excelSheet.Range("A8:I29")

    # Check if the range is empty
    if ($rangeExcel.Cells.Count -eq 0) {
        Write-Host "Error: The selected range is empty."
        return
    }

    # Copy the range as text
    $rangeExcel.Copy()

    # Insert the Excel data as a Word table (instead of inside a textbox)
    $range = $wordDoc.Content
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

    # Paste the Excel range as a Word table
    $wordTable = $range.PasteSpecial([Microsoft.Office.Interop.Word.WdRecoveryType]::wdFormatRTF)

    # Adjust the width and height of the table
    $wordTable = $wordDoc.Tables.Item($wordDoc.Tables.Count)

    # Set the table width to a specific size (in points)
    $wordTable.PreferredWidthType = [Microsoft.Office.Interop.Word.WdPreferredWidthType]::wdPreferredWidthPoints
    $wordTable.PreferredWidth = 500  # Constrain table width to 500 points
    $wordTable.AutoFitBehavior([Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow)

    # Set fixed width for all columns (15 points per column)
    foreach ($row in $wordTable.Rows) {
        foreach ($cell in $row.Cells) {
            # Access the range of each cell directly
            $cellRange = $cell.Range
            # Log the current font size before changing
            Write-Host "Before Change - Row: $($row.Index), Column: $($cell.ColumnIndex), Font Size: $($cellRange.Font.Size)"

            # Set the font size for the cell's range
            $cellRange.Font.Size = 10

            # Log the font size after applying the change
            Write-Host "After Change - Row: $($row.Index), Column: $($cell.ColumnIndex), Font Size: $($cellRange.Font.Size)"
        }
    }

    foreach ($column in $wordTable.Columns) {
        $column.Width = 105
    }

    # Set row height for each row
    foreach ($row in $wordTable.Rows) {
        $row.HeightRule = [Microsoft.Office.Interop.Word.WdRowHeightRule]::wdRowHeightExactly
        $row.Height = 16 # Adjust height for each row (in points)
    }

    # ---------------------------
    # Add the second image (Image Below) dynamically resized
    # ---------------------------
      $imageAbove = $wordDoc.Shapes.AddPicture(
        $ImageBelowPath, 
        $false,    # LinkToFile
        $true, 
  
        0,         # Left (0 means use the anchor’s position)
        -50,         # Top (0 means use the anchor’s position)
                 26.7 * 28.35,  # Width in points for landscape (adjusted)
        4 * 28.35,      # Height in points (adjusted)    # SaveWithDocument
        $range     # Anchor
    )
    
    # Fit image to available width in landscape (with some padding for margins)
    $imageAbove.LockAspectRatio = $true
    $imageAbove.Width = $pageWidth - 72  # Subtracting margins (1 inch = 72 points)
    $imageAbove.Top = 0

    $imageAbove.WrapFormat.Type = 3  # wdWrapFront
    $imageAbove.LockAnchor = $true

    # Insert some extra paragraphs to create a gap
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    $range.InsertParagraphAfter()
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
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
