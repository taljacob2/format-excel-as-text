function Parse-ExcelWorksheet {
    param (
        [parameter()][System.IO.FileSystemInfo]$Item = (Get-ChildItem *.xlsx),
        [parameter()][Int32]$WorkSheetIndex = 1
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $false
    $excel.DisplayAlerts = $false
    $workBook = $Excel.Workbooks.Open($Item.Fullname)

    $workSheet = $workBook.Worksheets($WorkSheetIndex)

    $workSheet | Format-Table

    # Write-Output $workSheet
    # Write-Output $workSheet["A1"]

    Cleanup-Excel -Excel $excel -WorkBook $workBook
}

function Convert-ExcelToCsv {
    param (
        [parameter()][System.IO.FileSystemInfo]$Item = (Get-ChildItem *.xlsx)
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $false
    $excel.DisplayAlerts = $false
    $workBook = $Excel.Workbooks.Open($Item.Fullname)

    # $workbook.SaveAs("$($Item.Fullname).txt", 42)   # xlUnicodeText

    $csvFullName = "$($Item.Fullname).csv"
    $workbook.SaveAs($csvFullName, 6)   # csv

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

Parse-ExcelWorksheet
$csvFullName = Convert-ExcelToCsv
Import-Csv $csvFullName > "$csvFullName.txt" 