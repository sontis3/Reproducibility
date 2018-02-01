Add-Type -TypeDefinition @'
    public enum AmountLevel {
        Low,
        Norm,
        High
    };
'@

################################################################### From excel-common.ps1
# формирование заголовка таблицы
function Add-TableTitle ($sheet, $rowNumber, $colNumber, $titles) {
  $i = $colNumber
  foreach ($item in $titles) {
      $sheet.Cells.Item($rowNumber, $i).Value = $item
      $sheet.Cells[$rowNumber, $i].Style.Border.Left.Style = "Thin"
      $i++
  }

  $sheet.Row($rowNumber).Style.HorizontalAlignment = "Center"
  $sheet.Row($rowNumber).Style.Font.Bold = $true
  $sheet.Row($rowNumber).Style.Border.Top.Style = $sheet.Row($rowNumber).Style.Border.Bottom.Style = "Thin"
}

# установить числовой формат столбцов из массива форматов начиная с n-го столбца
function Set-NumFormatColumns2 ($sheet, $colNumber, $formats) {
  $i = 0
  foreach ($item in $formats) {
      $sheet.Column($i + $colNumber).Style.Numberformat.Format = $item
      $i++
  }
}


#########################################################################################

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

# Вывод результатов в Excel
function Export-ExcelData() {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory = $True)] [string]$Path,             # путь файла
      [Parameter(Mandatory = $True)] [string[]]$dataInfo,       # жидкость пациент пол дата
      [Parameter(Mandatory = $False)] [System.Object]$outData   # данные, экспортитруемые в Excel
  )

  $excel = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $Path
  $wBook = $excel.Workbook

  $wSheets = $wBook.Worksheets     
  $sheet = $wSheets.Add("Результаты")
  $rowNumber = $sheet.Dimension.End.Row + 1

  Add-TableTitle $sheet $rowNumber 1 $dataInfo
  $sheet.Row($rowNumber).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
  $sheet.Row($rowNumber).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
  $rowNumber = $rowNumber + 2

  $measureUnit = "uM/L"
  if ($dataInfo[0] -eq "Urine") {
    $measureUnit = "uM/mmol crea"
  }
  $tableTitles1 = "Класс", "Метаболит", "min", "max", "Результат", "Отклонение"
  $tableTitles2 = $measureUnit, $measureUnit, $measureUnit

  Add-TableTitle $sheet $rowNumber 1 $tableTitles1
  $rowNumber++
  Add-TableTitle $sheet $rowNumber 3 $tableTitles2
  $rowNumber++

  foreach ($item in $outData) {
      $colNumber = 1
      $item.PSObject.Properties | ForEach-Object {
          $sheet.Cells[$rowNumber, $colNumber].Value = $_.value
          if ($colNumber -eq 6) {
            switch ($_.value) {
              "Low" { $sheet.Cells[$rowNumber, $colNumber].Value = "-"
                      $sheet.Cells[$rowNumber, ($colNumber-1)].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                      $sheet.Cells[$rowNumber, ($colNumber-1)].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
                      $sheet.Cells[$rowNumber, $colNumber].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                      $sheet.Cells[$rowNumber, $colNumber].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
                    }
              "Norm"  { $sheet.Cells[$rowNumber, $colNumber].Value = "" }
              "High"  { $sheet.Cells[$rowNumber, $colNumber].Value = "+"
                        $sheet.Cells[$rowNumber, ($colNumber-1)].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $sheet.Cells[$rowNumber, ($colNumber-1)].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightPink)
                        $sheet.Cells[$rowNumber, $colNumber].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $sheet.Cells[$rowNumber, $colNumber].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightPink)
                      }
              Default {$sheet.Cells[$rowNumber, $colNumber].Value = $_.value}
            }
          }
          # if ($_.value -eq "Failed") {
          #     $sheet.Cells[$rowNumber, $colNumber].Style.Font.Color.SetColor("Red")
          # }
          $colNumber++
      }
      $rowNumber++
  }

  Set-NumFormatColumns2 $sheet 3 "#0.000", "#0.000", "#0.000"
  $sheet.Column(6).Style.HorizontalAlignment = "Center"
  $sheet.Column(6).Style.Font.Bold = $true

  for ($i = 1; $i -lt $colNumber; $i++) {
      $sheet.Column($i).AutoFit()
  }
  
  $excel.Save()
  $excel.Dispose()
}
