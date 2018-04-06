Add-Type -TypeDefinition @'
    public enum AmountLevel {
        Low,
        Norm,
        High
    };
'@

. .\common-stat.ps1
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

# вывод строки данных в первый столбец $title
function Export-ExcelRow ($sheet, $rn, $cn, $title, $values) {
  $sheet.Row($rn).Style.HorizontalAlignment = "Center"
  $sheet.Cells[$rn, 1].Value = $title

  $values | ForEach-Object {
      if ($_ -is [string]) {
          $sheet.Cells[$rn, $cn].Value = $_
      } 
      else {
          $roundCount = $sheet.Cells[$rn, $cn].Style.Numberformat.Format.Split(".")[1].Length
          $sheet.Cells[$rn, $cn].Value = [math]::Round($_, $roundCount)
      }
      $cn++
  }
}
function Export-BoldExcelRow ($sheet, $rn, $cn, $title, $format, $values) {
  if ($format) {
      $sheet.Row($rn).Style.Numberformat.Format = $format
  }
  $sheet.Row($rn).Style.Font.Bold = $true
  Export-ExcelRow $sheet $rn $cn $title $values
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
  if ($sheet) {
    for ($i = $sheet.Dimension.Start.Row; $i -le $sheet.Dimension.End.Row; $i++) {
      if ($sheet.Cells[$i, 2].Value) {
        $item = [PSCustomObject]@{
          type = $sheet.Cells[$i, 1].Value;
          name = $sheet.Cells[$i, 2].Value;
        }
        $tableData += $item
      }
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

# вывод общих статистических данных
function Export-ExcelCommonStats ($sheet, $rn, $cn, $format, $stats) {
  Export-BoldExcelRow $sheet $rn $cn "Mean" $format $stats.means.value
  $rn++
  Export-BoldExcelRow $sheet $rn $cn "Gmean" $format $stats.gmeans.value
  $rn++
  Export-BoldExcelRow $sheet $rn $cn "Медиана" $format $stats.median.value
  $rn++
  Export-BoldExcelRow $sheet $rn $cn "Минимум" $format $stats.minimum.value
  $rn++
  Export-BoldExcelRow $sheet $rn $cn "Максимум" $format $stats.maximum.value
  $rn++
  Export-BoldExcelRow $sheet $rn $cn "SD" $format $stats.SD.value
  $rn++
  Export-BoldExcelRow $sheet $rn $cn "CV" $format $stats.RSD.value
  $sheet.Row($rn).Style.Font.Color.SetColor("Red")
  $sheet.Row($rn).Style.Border.Bottom.Style = "Thin"
  $rn++

  $rn
}

# формирование сводного отчета
function Export-ExcelSummary () {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory = $True)] [string]$Path,             # путь файла
      [Parameter(Mandatory = $True)] [string]$Diagnose,         # диагноз (1-я колонкаs)
      [Parameter(Mandatory = $True)] [System.Object]$outData,   # данные, экспортитруемые в Excel
      [Parameter(Mandatory = $True)] [System.Object]$stats      # общая статистика
  )

  $excel = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $Path
  $wBook = $excel.Workbook
  $wSheets = $wBook.Worksheets

  $titles = $outData[0].dataValues | Select-Object -ExpandProperty Peak_Name
  if (!$wSheets.Item("Сводная")) {
      $sheet = $wSheets.Add("Сводная")

      $rowNumber = $sheet.Dimension.End.Row + 1
      Add-TableTitle $sheet $rowNumber 1 (@("Диагноз", "Имя") + $titles)
      $rowNumber++
  } else {
    $sheet = $wSheets["Сводная"]
    $rowNumber = $sheet.Dimension.End.Row + 1
  }

  foreach ($item in $outData) {
    $colNumber = 1
    $sheet.Cells[$rowNumber, $colNumber++].Value = $Diagnose
    $sheet.Cells[$rowNumber, $colNumber++].Value = $item.name
    $titles | ForEach-Object {
      $amount = $item.dataValues | Where-Object Peak_Name -EQ $_ | Select-Object -ExpandProperty Amount
      if ($amount.GetType().Name -eq "String") {
        $sheet.Cells[$rowNumber, $colNumber++].Value = [System.Double]$amount
      } else {
        Write-Warning ("Дубликат имени " + $_)
        $sheet.Cells[$rowNumber, $colNumber++].Value = "?????"
      }
    }

    $sheet.Row($rowNumber).Style.Numberformat.Format = "#0.00000"
    $rowNumber++
  }

  $rowNumber = Export-ExcelCommonStats $sheet $rowNumber 3 "#0.000" $stats

  $allCells = $sheet.Cells[1, 1, $sheet.Dimension.End.Row, $sheet.Dimension.End.Column]
  $allCells.AutoFitColumns()

  $excel.Save()
  $excel.Dispose()
}

# формирование отчета T-test
function Export-T () {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory = $True)] [string]$Path,             # путь файла
      [Parameter(Mandatory = $True)] [System.Object]$t          # T-test
  )

  $excel = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $Path
  $wBook = $excel.Workbook
  $wSheets = $wBook.Worksheets
  if (!$wSheets.Item("Сводная")) {
      $sheet = $wSheets.Add("Сводная")
  } else {
    $sheet = $wSheets["Сводная"]
  }
  $rowNumber = $sheet.Dimension.End.Row + 1

  $colNumber = 1
  $sheet.Cells[$rowNumber, $colNumber].Value = "t-value"
  $colNumber = 3
  foreach ($item in $t) {
    $sheet.Cells[$rowNumber, $colNumber++].Value = $item.tValue
  }
  $sheet.Row($rowNumber).Style.Numberformat.Format = "#0.00000"
  $rowNumber++

  $colNumber = 1
  $sheet.Cells[$rowNumber, $colNumber].Value = "df"
  $colNumber = 3
  foreach ($item in $t) {
    $sheet.Cells[$rowNumber, $colNumber++].Value = $item.dF
  }
  $sheet.Row($rowNumber).Style.Numberformat.Format = "#0"
  $rowNumber++

  $colNumber = 1
  $sheet.Cells[$rowNumber, $colNumber].Value = "p"
  $colNumber = 3
  foreach ($item in $t) {
    if ($item.p -lt 0.05) {
      $sheet.Cells[$rowNumber, $colNumber].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
      $sheet.Cells[$rowNumber, $colNumber].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightPink)
      # $sheet.Cells[$rowNumber, $colNumber].Style.Font.Color.SetColor("Red")
    }
    $sheet.Cells[$rowNumber, $colNumber++].Value = $item.p
  }
  $sheet.Row($rowNumber).Style.Numberformat.Format = "#0.00000"
  $rowNumber++

  $colNumber = 1
  $sheet.Cells[$rowNumber, $colNumber].Value = "F-ratio"
  $colNumber = 3
  foreach ($item in $t) {
    $sheet.Cells[$rowNumber, $colNumber++].Value = $item.fva
  }
  $sheet.Row($rowNumber).Style.Numberformat.Format = "#0.00000"
  $rowNumber++

  $colNumber = 1
  $sheet.Cells[$rowNumber, $colNumber].Value = "p variances"
  $colNumber = 3
  foreach ($item in $t) {
    if ($item.pVar -lt 0.05) {
      $sheet.Cells[$rowNumber, $colNumber].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
      $sheet.Cells[$rowNumber, $colNumber].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightPink)
    }
    $sheet.Cells[$rowNumber, $colNumber++].Value = $item.pVar
  }
  $sheet.Row($rowNumber).Style.Numberformat.Format = "#0.00000"
  $rowNumber++

  $allCells = $sheet.Cells[1, 1, $sheet.Dimension.End.Row, $sheet.Dimension.End.Column]
  $allCells.AutoFitColumns()

  $excel.Save()
  $excel.Dispose()
}

