[CmdletBinding()]
Param(
  [Parameter(Mandatory = $False)] [string]$inPath
)

if (!$inPath) {
  $scriptPath = Split-Path $MyInvocation.MyCommand.path
  $systemPath = Split-Path -Parent $scriptPath
  # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\All'
  # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Amino Acids'
  # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Fatty Acids_All'
  # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Fatty Acids_Free'
  # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'ProstateCancer'
  $inPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\All'
}

$normPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Normal concentrations.xlsx'
# $normPath = Split-Path -Parent $systemPath | Join-Path -ChildPath "Normal concentrations.xlsx"
# $normPath = Split-Path -Parent $systemPath | Join-Path -ChildPath "Amino acids Normal concentrations.xlsx" 
# $normPath = Split-Path -Parent $systemPath | Join-Path -ChildPath "Fatty acids All Normal concentrations.xlsx" 
# $normPath = Split-Path -Parent $systemPath | Join-Path -ChildPath "Fatty acids Free Normal concerntrations.xlsx" 

#-------------------------------------------------------------------------------
# загрузка нормальных диапазонов
. .\Invoke-Init.ps1 -inPath $inPath
. ".\xml-helpers.ps1"

#-----------------------------------------------------------------------------------------------
function Process-Data() {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$True)] [string]$fluidName,
    [Parameter(Mandatory=$True)] [System.Object[]]$ranges,
    [Parameter(Mandatory=$False)] [System.Object[]]$stopList,
    [Parameter(Mandatory=$True)] [System.Object[]]$inData
  )
  $result = @()

  $urineFactor = "Creatinine"
  $urineFactorValue = 0.0
  # поиск фактора для урины
  if ($fluidName -eq "Urine") {
    $item = $inData.Where({$_.Peak_Name -eq $urineFactor})
    if ($item) {
      $urineFactorValue = [Double]$item.Amount
    }
  }

  foreach ($item in $inData) {
    $rng = $ranges.Where({$_.name -eq $item.Peak_Name})
    if ($rng) {
      if ($fluidName -eq "Urine") {
        if ($urineFactorValue -eq 0.0) {
          Write-Error ("$urineFactor = $urineFactorValue")
          $amount = 0.0
        }
        else {
          $amount = [Double]$item.Amount / $urineFactorValue
        }
      }
      else {
        $amount = [Double]$item.Amount
      }
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
    } elseif (!$stopList -or !$stopList.Where({$_.name -eq $item.Peak_Name})) {
      Write-Warning ($fluidName + ": Не найден диапазон для " + $item.Peak_Name)
    }
  }

  $result
}

function Process-XmlDataMult() {
  [CmdletBinding()]
  Param(
      [Parameter(Mandatory=$True)] [string]$Path,                       # каталог с исходными файлами
      [Parameter(Mandatory=$True)] [string]$tmpPath,                    # каталог временных(результат) файлов
      [Parameter(Mandatory=$True)] [string]$Format,                     # формат данных
      [Parameter(Mandatory=$True)] [System.Object[]]$fields,            # представление данных in/out
      [Parameter(Mandatory=$True)] [System.Object]$ranges,              # диапазоны
      [Parameter(Mandatory=$False)] [System.Object[]]$stopList           # неучитываемые параметры
  )

  Write-Host '-------------- start job (Load from multiple xml to stage)---------------'
  # удаление старых файлов
  $tmpFilePath = Join-Path -Path $tmpPath -ChildPath '*.*'
  if(Test-Path $tmpFilePath){
    Remove-Item $tmpFilePath
  }

  Get-ChildItem -Path $Path -Filter *.xml |
      ForEach-Object {
        Write-Host '-------------- ' + $_.Name + ' ---------------'
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

            $outData = Process-Data -fluidName $selFluidName -ranges $ranges.$selFluidName -stopList $stopList -indata $inData

            $baseName = $_ | Select-Object -ExpandProperty BaseName
            $fileName = $baseName + ".xlsx"
            $ExcelPath = Join-Path -Path $tmpPath -ChildPath $fileName
            $dataInfo = $baseName.Split("_")

            Export-ExcelData -Path $ExcelPath -dataInfo $dataInfo -outData $outData
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
