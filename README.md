# format-excel-as-text

A PowerShell script that receives an Excel file, formats it, and exports it as formatted text file(s).

## Usage

The following commands should be run with *powershell.exe*.

### Execute Inline Without Cloning

Syntax:

```ps1
& ([scriptblock]::Create((iwr https://raw.githubusercontent.com/taljacob2/format-excel-as-text/master/format-excel-as-text.ps1 -useb))) [-Path (<string>)]
```

Examples:

- Format the `.xls` Excel file in the current directory to text file(s).
  ```ps1
  & ([scriptblock]::Create((iwr https://raw.githubusercontent.com/taljacob2/format-excel-as-text/master/format-excel-as-text.ps1 -useb))) -Path "*.xls"
  ```

- Format an Excel file by its absolute path to text file(s).
  ```ps1
  & ([scriptblock]::Create((iwr https://raw.githubusercontent.com/taljacob2/format-excel-as-text/master/format-excel-as-text.ps1 -useb))) -Path "C:\Users\demo.xlsx"
  ```

### Execute Offline

#### Clone The Project

```
git clone https://github.com/taljacob2/format-excel-as-text.ps1
```

#### Run

```
.\format-excel-as-text.ps1.ps1 -Path <string>
```

In case you encouter an error, try running with:
```
powershell.exe -NoLogo -ExecutionPolicy Bypass -Command ".\format-excel-as-text.ps1 -Path <string>"
```

## Help

To view the full documentation of the script, run:
```
Get-Help .\format-excel-as-text.ps1 -Full
```
