function Format-ExcelWorksheet {
    param (
        [parameter()][System.IO.FileSystemInfo]$Item = (Get-ChildItem *.xlsx),
        [parameter()][Int32]$WorkSheetIndex = 1
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $false
    $excel.DisplayAlerts = $false
    $workBook = $Excel.Workbooks.Open($Item.Fullname)

    $workSheet = $workBook.Worksheets($WorkSheetIndex)
    
    # Format numbers.
    $range = $workSheet.Range("a1","z9")
    $range.NumberFormat = "000000.00"

    $newFormattedFileFullName = "$($Item.Fullname).formatted"
    $workbook.SaveAs($newFormattedFileFullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)

    Cleanup-Excel -Excel $excel -WorkBook $workBook

    return $newFormattedFileFullName
}

function Convert-ExcelToCsv {
    param (
        [parameter()][System.IO.FileSystemInfo]$Item = (Get-ChildItem *.xlsx)
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $false
    $excel.DisplayAlerts = $false
    $workBook = $Excel.Workbooks.Open($Item.Fullname)

    $csvFullName = "$($Item.Fullname).csv"
    $workbook.SaveAs($csvFullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)

    Cleanup-Excel -Excel $excel -WorkBook $workBook

    return $csvFullName
}

function Cleanup-Excel {
    param (
        [parameter(mandatory)][Microsoft.Office.Interop.Excel.ApplicationClass]$Excel,
        [parameter(mandatory)][System.__ComObject]$WorkBook
    )

    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$newFormattedFileFullName = Format-ExcelWorksheet
$csvFullName = Convert-ExcelToCsv -Item (Get-ChildItem $newFormattedFileFullName)
Import-Csv $csvFullName > "$csvFullName.txt"

# Clean temporary files used for calculations.
rm $newFormattedFileFullName
rm $csvFullName

Write-Output "The result file is '$csvFullName.txt'"