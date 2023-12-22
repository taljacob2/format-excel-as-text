$files = Get-ChildItem *.xls*

$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.DisplayAlerts = $false

ForEach ($file in $files) {
    Write-Output "Loading File '$($file.Name)'..."
    $WorkBook = $Excel.Workbooks.Open($file.Fullname)
    $NewFilePath = [System.IO.Path]::ChangeExtension($file.Fullname, ".txt")
    $Workbook.SaveAs($NewFilepath, 42)   # xlUnicodeText
}

# cleanup
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
