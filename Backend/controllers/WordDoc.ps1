param (
    [string]$ExcelFilePath,
    [string]$ImageAbovePath,
    [string]$ImageBelowPath,
    [string]$outputWordFilePath
)

$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false  # Hide Word application (set to $true to show)
$wordDoc = $wordApp.Documents.Add()

# Open Excel file and get the workbook object
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false  # Hide Excel application (set to $true to show)
$excelApp.DisplayAlerts = $false  # Disable Excel pop-ups and alerts

$workbook = $excelApp.Workbooks.Open($ExcelFilePath)
# Loop over each sheet in the Excel workbook
$WordDoc.Content.InsertParagraphAfter()

foreach ($sheet in $workbook.Sheets) {
    # Start with a fresh page in Word
    $wordDoc.Content.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
    $WordDoc.Content.InsertParagraphAfter()

    # Add the first image (Image Above) as a floating shape
    $imageAbove = $wordDoc.Shapes.AddPicture($ImageAbovePath)

    # Set the height and width of the image in cm (1 cm = 28.35 points)
    $imageAbove.Height = 3 * 28.35  
    $imageAbove.Width = 18.88 * 28.35  

    # Set text wrapping for the image (wrap in front of the text)
    $imageAbove.WrapFormat.Type = 3  # 3 corresponds to wdWrapFront
    $imageAbove.LockAnchor = $true  # Fix the position relative to the text

    # Insert an artificial gap by adding blank paragraphs (or you can manually adjust spacing here)
    $wordDoc.Content.InsertParagraphAfter()
    $wordDoc.Content.InsertParagraphAfter()

    # Create a text box to hold the Excel table
    $textBox = $wordDoc.Shapes.AddTextbox(1, 0, 0, 500, 300)  # 1 for msoTextOrientationHorizontal

    # Center the text box horizontally by setting the Left property
    $pageWidth = $wordApp.ActiveDocument.PageSetup.PageWidth
    $textBoxWidth = $textBox.Width
    $textBox.Left = ($pageWidth - $textBoxWidth) / 2  # Center the text box on the page

    # Set the wrapping style to none for the text box to prevent it from wrapping around the text
    $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 0  # Align text in the text box (optional)
    $textBox.WrapFormat.Type = 0  # 0 corresponds to wdWrapNone (no text wrapping around the box)

    # Set a vertical offset for the text box to create space after the image
    $textBox.Top = $imageAbove.Top + $imageAbove.Height + 10  # Add a small gap (10 points) between image and table

    # Copy the table content from Excel (Range A8:I28)
      $excelSheet = $sheet
    $range = $excelSheet.Range("A8:H30")
    $range.CopyPicture([Microsoft.Office.Interop.Excel.XlPictureAppearance]::xlScreen, [Microsoft.Office.Interop.Excel.XlCopyPictureFormat]::xlPicture)


    # Paste the content into the text box
    $textBox.TextFrame.TextRange.Paste()

    # Remove the border of the text box (set Line.Visible to false)
    $textBox.Line.Visible = $false

    # Resize the pasted Excel table to fit within the text box
    $pastedShape = $textBox.TextFrame.TextRange
    $pastedShape.Select()  # Select the pasted content

    # Ensure maxWidth and maxHeight are defined correctly (non-zero values)
    $maxWidth = 450  # Set the maximum width (you can adjust this)
    $maxHeight = 250  # Set the maximum height (you can adjust this)
    
  
    # Insert additional space after the text box (you can adjust this gap as needed)
    $wordDoc.Content.InsertParagraphAfter()
    $wordDoc.Content.InsertParagraphAfter()

    # Add the second image (Image Below) as a floating shape
    $imageBelow = $wordDoc.Shapes.AddPicture($ImageBelowPath)

    # Set the height and width of the image in cm (1 cm = 28.35 points)
    $imageBelow.Height = 3 * 28.35
    $imageBelow.Width = 18.88 * 28.35

    # Set text wrapping for the image (wrap in front of the text)
    $imageBelow.WrapFormat.Type = 3  # 3 corresponds to wdWrapFront
    $imageBelow.LockAnchor = $true  # Fix the position relative to the text

    # Set a vertical offset for the second image (after the text box)
    $imageBelow.Top = [float]$textBox.Top + [float]$textBox.Height + 10  # Add a small gap (10 points) between table and second image

    # Apply Page Setup to the whole document (not to text boxes or shapes)
    $wordDoc.PageSetup.Orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientPortrait
    $wordDoc.PageSetup.TopMargin = 5
    $wordDoc.PageSetup.BottomMargin = 5
    $wordDoc.PageSetup.LeftMargin = 5
    $wordDoc.PageSetup.RightMargin = 5

    # Save the current Word document after each iteration
    $wordDoc.SaveAs([ref]$outputWordFilePath)
    Start-Sleep -Seconds 2
}

# Close Word and Excel applications
$workbook.Close()
$excelApp.Quit()
$wordDoc.Close()
$wordApp.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDoc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
