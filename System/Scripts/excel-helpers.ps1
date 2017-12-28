﻿Add-Type -TypeDefinition @'
    public enum AmountLevel {
        Low,
        Norm,
        High
    };
'@

# импорт диапазонов нормальных данных для биожидкости из Normal concentrations blood.xlsx
function Import-BioFluid-Ranges () {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory = $True)] [string]$Path,   # путь к файлу с данными
      [Parameter(Mandatory = $True)] [string]$Name    # имя страницы
  )
  $excel = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $Path
  $wBook = $excel.Workbook
  $wSheets = $wBook.Worksheets
  $sheet = $wSheets[$Name]

  $tableData = @()

  for ($i = $sheet.Dimension.Start.Row; $i -le $sheet.Dimension.End.Row; $i++) {
    if (($sheet.Cells[$i, 3].Value -ge 0.0) -and $sheet.Cells[$i, 3].Value.GetType().Name -eq "Double") {
      $item = [PSCustomObject]@{
        type = $sheet.Cells[$i, 1].Value;
        name = $sheet.Cells[$i, 2].Value;
        lowConc = $sheet.Cells[$i, 3].Value;
        highConc = $sheet.Cells[$i, 4].Value;
      }
      $tableData += $item
    }
  }
  # foreach ($range in $wBook.Names) {
  #     if ($range.Name -eq $Name) {
  #         for ($rn = $range.Start.Row; $rn -le $range.End.Row; $rn++) {
  #             $values = @()
  #             for ($cn = $range.Start.Column + 2; $cn -le $range.End.Column; $cn++) {
  #                 $values += $sheet.Cells[$rn, $cn].Value
  #             }
  #             $tableData.sample += [PSCustomObject]@{
  #                 randNum = $sheet.Cells[$rn, $range.Start.Column].Value;
  #                 period  = $sheet.Cells[$rn, ($range.Start.Column + 1)].Value;
  #                 values  = $values;
  #             }
  #         }
  #     }
  # }
  
  $excel.Dispose()

  # заполнение $tableData.colSlice
  # for ($i = 0; $i -lt $tableData.sample[0].values.Count; $i++) {
  #     $timeSlice = @();
  #     $tableData.sample | ForEach-Object {$timeSlice += $_.values[$i]}
  #     $tableData.colSlice += [PSCustomObject]@{
  #         #time = $timeList[$i]
  #         values = $timeSlice
  #     }
  # }
  $tableData
}

# импорт стоплиста из Normal concentrations blood.xlsx
function Import-StopList () {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory = $True)] [string]$Path,   # путь к файлу с данными
      [Parameter(Mandatory = $True)] [string]$Name    # имя страницы
  )
  $excel = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $Path
  $wBook = $excel.Workbook
  $wSheets = $wBook.Worksheets
  $sheet = $wSheets[$Name]

  $tableData = @()

  for ($i = $sheet.Dimension.Start.Row; $i -le $sheet.Dimension.End.Row; $i++) {
    if ($sheet.Cells[$i, 2].Value) {
      $item = [PSCustomObject]@{
        type = $sheet.Cells[$i, 1].Value;
        name = $sheet.Cells[$i, 2].Value;
      }
      $tableData += $item
    }
  }
  
  $excel.Dispose()

  $tableData
}