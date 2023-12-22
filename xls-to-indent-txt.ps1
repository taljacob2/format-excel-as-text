function Parse-Worksheet {
    param (
        [parameter(mandatory)][System.__ComObject]$WorkBook,
        [parameter()][Int32]$WorkSheetNumber = 1
    )

    $workSheet = $WorkBook.Worksheets($WorkSheetNumber)

    # Write-Output $workSheet
    Write-Output $workSheet["A1"]
}

$files = Get-ChildItem *.xls

$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.DisplayAlerts = $false

ForEach ($file in $files) {
    Write-Output "Loading File '$($file.Name)'..."
    $workBook = $Excel.Workbooks.Open($file.Fullname)
    
    Parse-Worksheet -WorkBook $workBook

    # $workbook.SaveAs("$($file.Fullname).txt", 42)   # xlUnicodeText
}

# cleanup
$Excel.Quit()
# [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
