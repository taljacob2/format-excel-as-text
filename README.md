# format-excel-as-text

A PowerShell script that receives an Excel file, formats it, and exports it as formatted text file(s).

## Usage

The following commands should be run with *powershell.exe*.

### Execute Inline Without Cloning

Syntax:

```ps1
& ([scriptblock]::Create((iwr https://raw.githubusercontent.com/taljacob2/format-excel-as-text/master/format-excel-as-text.ps1 -useb))) [[-Path] <String>] [[-RangeEndCell] <String>]
```

Examples:

- *Default:* Format the `.xlsx` Excel file in the current directory to text file(s), until cell "z999".

  ```ps1
  & ([scriptblock]::Create((iwr https://raw.githubusercontent.com/taljacob2/format-excel-as-text/master/format-excel-as-text.ps1 -useb)))
  ```

- Format the `.xls` Excel file in the current directory to text file(s), until cell "j23".

  ```ps1
  & ([scriptblock]::Create((iwr https://raw.githubusercontent.com/taljacob2/format-excel-as-text/master/format-excel-as-text.ps1 -useb))) -Path "*.xls" -RangeEndCell "j23"
  ```

- Format an Excel file by its absolute path to text file(s), until cell "z999".

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
.\format-excel-as-text.ps1.ps1 [[-Path] <String>] [[-RangeEndCell] <String>]
```

In case you encouter an error, try running with:

```
powershell.exe -NoLogo -ExecutionPolicy Bypass -Command ".\format-excel-as-text.ps1 [[-Path] <String>] [[-RangeEndCell] <String>]"
```

## Help

To view the full documentation of the script, run:

```
Get-Help .\format-excel-as-text.ps1 -Full
```
