param(
    [string]$filePath
)

try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Open the workbook
    $workbook = $excel.Workbooks.Open($filePath)
    
    # Run the macro
    $excel.Run("CalculateUnusedTimings")
    
    # Get data from Statistics sheet
    $statSheet = $workbook.Worksheets.Item("Statistics")
    $lastRow = $statSheet.UsedRange.Rows.Count
    
    # # Prepare data for output
    # $result = @{
    #     sheets = @()
    #     totals = @{
    #         totalSlots = 0
    #         usedSlots = 0
    #         unusedSlots = 0
    #     }
    # }

    # # Collect data from each row (starting from row 2)
    # for ($i = 2; $i -le $lastRow; $i++) {
    #     try {
    #         $sheetName = $statSheet.Cells.Item($i, 1).Text
    #         $totalSlots = [int]$statSheet.Cells.Item($i, 2).Value
    #         $usedSlots = [int]$statSheet.Cells.Item($i, 3).Value
    #         $unusedSlots = [int]$statSheet.Cells.Item($i, 4).Value
    #         $utilization = $statSheet.Cells.Item($i, 5).Text
    #         $unusedTimings = $statSheet.Cells.Item($i, 6).Text

    #         # Clean up utilization percentage (remove % if present)
    #         $utilization = $utilization -replace '%',''
    #         $utilization = [math]::Round([double]$utilization, 2)
    #         $utilization = "$utilization%"

    #         # Clean up unused timings (remove extra spaces and normalize)
    #         $unusedTimings = $unusedTimings -replace '\s+',' '
    #         $unusedTimings = $unusedTimings.Trim()

    #         $sheetData = @{
    #             sheetName = $sheetName
    #             totalSlots = $totalSlots
    #             usedSlots = $usedSlots
    #             unusedSlots = $unusedSlots
    #             utilization = $utilization
    #             unusedTimings = $unusedTimings
    #         }
            
    #         $result.sheets += $sheetData
            
    #         # Update totals
    #         $result.totals.totalSlots += $totalSlots
    #         $result.totals.usedSlots += $usedSlots
    #         $result.totals.unusedSlots += $unusedSlots
    #     }
    #     catch {
    #         Write-Warning "Error processing row $i : $_"
    #         continue
    #     }
    # }
    
    # # Calculate overall utilization
    # if ($result.totals.totalSlots -gt 0) {
    #     $utilization = [math]::Round(($result.totals.usedSlots / $result.totals.totalSlots) * 100, 2)
    #     $result.totals.utilization = "$utilization%"
    # } else {
    #     $result.totals.utilization = "0.00%"
    # }
    
    # # Close workbook without saving
    $workbook.Close($false)
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($statSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Return as JSON
    $result | ConvertTo-Json -Depth 5
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}