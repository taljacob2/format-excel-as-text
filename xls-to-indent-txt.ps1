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

function Export-CsvAsTxtTable {
    param (
        [parameter(mandatory)][string]$CsvFullName
    )

    Import-Csv $CsvFullName > "$CsvFullName.txt"

    return "$CsvFullName.txt"
}

function Export-CsvAsDelimitedTxtTable {
    param (
        [parameter(mandatory)][string]$CsvFullName,
        [parameter()][Char]$Delimiter = '~'
    )

    $delimitedCsv = $(Import-Csv $CsvFullName)
    $delimitedCsv | Foreach-Object { 
        foreach ($property in $_.PSObject.Properties)
        {
            $property.Value = "$($property.Value)$Delimiter"
        }
    }

    Write-Output $delimitedCsv > "$CsvFullName.delimited.txt"

    return "$CsvFullName.delimited.txt"
}

$newFormattedFileFullName = Format-ExcelWorksheet
$csvFullName = Convert-ExcelToCsv -Item (Get-ChildItem $newFormattedFileFullName)
$csvAsTxtTableFullName = Export-CsvAsTxtTable -CsvFullName $csvFullName
$csvAsDelimitedTxtTable = Export-CsvAsDelimitedTxtTable -CsvFullName $csvFullName


# Clean temporary files used for calculations.
rm $newFormattedFileFullName
rm $csvFullName

# Write results to console.
Write-Output "New file: '$csvAsTxtTableFullName'"
Write-Output "New file: '$csvAsDelimitedTxtTable'"
