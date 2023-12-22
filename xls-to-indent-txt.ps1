function Parse-ExcelWorksheet {
    param (
        [parameter()][System.IO.FileSystemInfo]$Item = (Get-ChildItem *.xlsx),
        [parameter()][Int32]$WorkSheetNumber = 1
    )

    $Excel = New-Object -ComObject Excel.Application
    $Excel.visible = $false
    $Excel.DisplayAlerts = $false
    $workBook = $Excel.Workbooks.Open($Item.Fullname)

    $workSheet = $workBook.Worksheets($WorkSheetNumber)

    # Write-Output $workSheet
    Write-Output $workSheet["A1"]
}

function Convert-XlsToCsv {
    param (
        [parameter()][System.IO.FileSystemInfo]$Item = (Get-ChildItem *.xlsx)
    )

    $Excel = New-Object -ComObject Excel.Application
    $Excel.visible = $false
    $Excel.DisplayAlerts = $false
    $workBook = $Excel.Workbooks.Open($Item.Fullname)

    # $workbook.SaveAs("$($Item.Fullname).txt", 42)   # xlUnicodeText

    $csvFullName = "$($Item.Fullname).csv"
    $workbook.SaveAs($csvFullName, 6)   # csv

    # cleanup
    $Excel.Quit()

    return $csvFullName
}

Parse-ExcelWorksheet
$csvFullName = Convert-ExcelToCsv
Import-Csv $csvFullName > out.txt 