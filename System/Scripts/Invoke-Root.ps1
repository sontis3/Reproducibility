﻿[CmdletBinding()]
Param(
  [Parameter(Mandatory = $False)] [string]$inPath
)

if (!$inPath) {
  $scriptPath = Split-Path $MyInvocation.MyCommand.path
  $systemPath = Split-Path -Parent $scriptPath
  $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data'
}

#-------------------------------------------------------------------------------
# загрузка нормальных диапазонов
. .\Invoke-Init.ps1 -inPath $inPath
. ".\xml-helpers.ps1"

#-----------------------------------------------------------------------------------------------
function Process-Data() {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory=$True)] [System.Object[]]$ranges,
      [Parameter(Mandatory=$True)] [System.Object[]]$stopList,
      [Parameter(Mandatory=$True)] [System.Object[]]$inData
  )
  $result = @()

  foreach ($item in $inData) {
    $rng = $ranges.Where({$_.name -eq $item.Peak_Name})
    if ($rng) {
      $amount = [Double]$item.Amount
      if ($amount -lt $rng[0].lowConc) {
        $level = [AmountLevel]::Low
      } elseif ($amount -gt $rng[0].highConc) {
        $level = [AmountLevel]::High
      } else {
        $level = [AmountLevel]::Norm
      }
      $result += [PSCustomObject]@{
        type = $rng[0].type;
        name = $rng[0].name;
        lowConc = $rng[0].lowConc;
        highConc = $rng[0].highConc;
        amount = $amount;
        level = $level;
      }
    } elseif (!$stopList.Where({$_.name -eq $item.Peak_Name})) {
      Write-Warning ("Не найден диапазон для " + $item.Peak_Name)
    }
  }

  $result
}

function Process-XmlDataMult() {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory=$True)] [string]$Path,                        # каталог с исходными файлами
      [Parameter(Mandatory=$True)] [string]$tmpPath,                     # каталог временных(результат) файлов
      [Parameter(Mandatory=$True)] [string]$Format,                      # формат данных
      [Parameter(Mandatory=$True)] [System.Object[]]$fields,
      [Parameter(Mandatory=$True)] [System.Object]$ranges,
      [Parameter(Mandatory=$True)] [System.Object[]]$stopList
  )

  Write-Host '-------------- start job (Load from multiple xml to stage)---------------'
  # удаление старых файлов
  $tmpFilePath = Join-Path -Path $tmpPath -ChildPath '*.rem'
  if(Test-Path $tmpFilePath){
      Remove-Item $tmpFilePath
  }

  Get-ChildItem -Path $Path -Filter *.xml |
      ForEach-Object {
          switch -regex ($_.Name) {
            $BioFluidNames[0] { $selFluidName = $BioFluidNames[0] }
            $BioFluidNames[1] { $selFluidName = $BioFluidNames[1] }
            $BioFluidNames[2] { $selFluidName = $BioFluidNames[2] }
            Default {
              Write-Warning ("Не найден код жидкости для " + $_.Name)
              $selFluidName = $null
            }
          }

          if ($selFluidName -and $ranges.$selFluidName) {
            $inData = Get-XmlData -FilePath $_.FullName -Format $Format -Filter " " -fields ($fields  | Select-Object -ExpandProperty "in")

            $outData = Process-Data -ranges $ranges.$selFluidName -stopList $stopList -indata $inData
          }
      }

  Write-Host '-------------- end job (Load from multiple xml to stage)---------------'
}
#--------------------------------------------------------------------------------------------------

$ranges = Import-Norm-Ranges
$stopList = Import-StopList -Path $normPath -Name $StopListName

$fields = @(
  [PSCustomObject]@{in = "Number"; out = "Number"; },
  [PSCustomObject]@{in = "Peak_Name"; out = "Name"; },
  [PSCustomObject]@{in = "Amount"; out = "Amount"; }
)


Process-XmlDataMult -Path $inPath -tmpPath $tmpPath -Format "AlexPasha" -fields $fields -ranges $ranges -stopList $stopList

$qwe = 123

# Импорт корректировки из Excel и формирование correct.sample.json
# . .\Invoke-PostCorrect.ps1 -inPath $inPath

# Формирование таблиц 1.2 + (TableData1.json + TableData2.json)
# . .\Invoke-CommonStat.ps1 -inPath $inPath

# . .\Invoke-PkStat.ps1 -inPath $inPath
