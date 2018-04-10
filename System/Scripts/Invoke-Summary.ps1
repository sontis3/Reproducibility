[CmdletBinding()]
Param(
    [Parameter(Mandatory = $False)] [string]$inPath,
    [Parameter(Mandatory = $False)] [string]$outPath
)

$ErrorActionPreference = "Stop"

if (!$inPath) {
    $scriptPath = Split-Path $MyInvocation.MyCommand.path
    $systemPath = Split-Path -Parent $scriptPath
    # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\All'
    # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Amino Acids'
    # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Fatty Acids_All'
    # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Fatty Acids_Free'
    # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'ProstateCancer'
    # $inPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\PH-PC Serum'
    # $inPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\PH-PC Urine'
    # $inPath = 'Z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\QC\Serum'
    # $inPath = 'Z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\QC\Urine'
    $inPath = 'Z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\PH-PC Serum AA'
}

#-------------------------------------------------------------------------------
# загрузка драйвера Excel и нормальных диапазонов
. .\Invoke-Init.ps1 -inPath $inPath
# поддержка XML
. ".\xml-helpers.ps1"

#-----------------------------------------------------------------------------------------------
# загрузить исходные данные
function Get-XmlDataMult() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)] [string]$Path,               # каталог с исходными файлами
        [Parameter(Mandatory = $True)] [string]$Format,             # формат данных
        [Parameter(Mandatory = $True)] [System.Object[]]$fields,    # представление данных in/out
        [Parameter(Mandatory = $True)] [string]$filterTemplate      # шаблон фильтра
    )
  
    Write-Host '-------------- start job (Load from multiple xml to stage)---------------'
    $tableData = @()

    Get-ChildItem -Path $Path -Filter $filterTemplate | Sort-Object -Property Name |
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


        $inData = Get-XmlData -FilePath $_.FullName -Format $Format -Filter " " -fields ($fields  | Select-Object -ExpandProperty "in")

        $baseName = $_ | Select-Object -ExpandProperty BaseName
        $dataInfo = $baseName.Split("_")

        $item = [PSCustomObject]@{
            name = $dataInfo[1];
            dataValues = $inData;
            selFluidName = $selFluidName;
        }

        $tableData += $item
    }
  
    Write-Host '-------------- end job (Load from multiple xml to stage)---------------'
    
    $tableData
}
  
function Prepare-Stats () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)] [System.Object]$data      # данные, экспортитруемые в Excel
    )

    $urineFactor = "Creatinine"
    $data | ForEach-Object {
        if ($_.SelFluidName -eq "Urine") {
            $dataValues = $_.dataValues
            $item = $dataValues.Where({$_.Peak_Name -eq $urineFactor})
            if ($item) {
              $urineFactorValue = [Double]$item.Amount
            }
            $dataValues | ForEach-Object {
                $_.Amount /= $urineFactorValue
            }
        }
    }

    $titles = $data[0].dataValues | Select-Object -ExpandProperty Peak_Name
    $samples  = @()
    $titles | ForEach-Object {
      $selSample = $data | Select-Object -ExpandProperty dataValues | Where-Object Peak_Name -EQ $_
      $sample = $selSample | Select-Object -Property @{name="Amount";expression={[System.Double]$_.Amount}} | Select-Object -ExpandProperty Amount
      $samples += [PSCustomObject]@{
        name = $_
        values = $sample
      }
    }
  
    Get-CommonStats -samples $samples
}
########################################################################################

$fields = @(
    [PSCustomObject]@{in = "Number"; out = "Number"; },
    [PSCustomObject]@{in = "Peak_Name"; out = "Name"; },
    [PSCustomObject]@{in = "Amount"; out = "Amount"; }
)

# удаление старых файлов
$tmpFilePath = Join-Path -Path $tmpPath -ChildPath '*.*'
if (Test-Path $tmpFilePath) {
    Remove-Item $tmpFilePath
}
$ExcelPath = Join-Path -Path $tmpPath -ChildPath "summary.xlsx"

$phData = Get-XmlDataMult -Path $inPath -Format "AlexPasha" -fields $fields -filterTemplate "*_PH*.xml"
$pcData = Get-XmlDataMult -Path $inPath -Format "AlexPasha" -fields $fields -filterTemplate "*_PC*.xml"

$phStats = Prepare-Stats -data $phData
$pcStats = Prepare-Stats -data $pcData

Export-ExcelSummary -Path $ExcelPath -Diagnose "PH" -outData $phData -stats $phStats
Export-ExcelSummary -Path $ExcelPath -Diagnose "PC" -outData $pcData -stats $pcStats

$t = @()
for ($i = 0; $i -lt $phStats.means.Count; $i++) {
    $t += Get-T -nT $phData.Count -meanT $phStats.means[$i].value -sdT $phStats.SD[$i].value -nR $pcData.Count -meanR $pcStats.means[$i].value -sdR $pcStats.SD[$i].value
}

Export-T -Path $ExcelPath -t $t

# $qcData = Get-XmlDataMult -Path $inPath -Format "AlexPasha" -fields $fields -filterTemplate "*.xml"
# $qcStats = Prepare-Stats -data $qcData
# Export-ExcelSummary -Path $ExcelPath -Diagnose "QC" -outData $qcData -stats $qcStats
