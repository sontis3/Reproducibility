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
    $inPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\PH-PC Serum'
    # $inPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\PH-PC Urine'
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
        $inData = Get-XmlData -FilePath $_.FullName -Format $Format -Filter " " -fields ($fields  | Select-Object -ExpandProperty "in")

        $baseName = $_ | Select-Object -ExpandProperty BaseName
        $dataInfo = $baseName.Split("_")

        $item = [PSCustomObject]@{
            name = $dataInfo[1];
            dataValues = $inData;
        }

        $tableData += $item
    }
  
    Write-Host '-------------- end job (Load from multiple xml to stage)---------------'
    
    $tableData
}
  
$fields = @(
    [PSCustomObject]@{in = "Number"; out = "Number"; },
    [PSCustomObject]@{in = "Peak_Name"; out = "Name"; },
    [PSCustomObject]@{in = "Amount"; out = "Amount"; }
)

$phData = Get-XmlDataMult -Path $inPath -Format "AlexPasha" -fields $fields -filterTemplate "*_PH*.xml"
$pcData = Get-XmlDataMult -Path $inPath -Format "AlexPasha" -fields $fields -filterTemplate "*_PC*.xml"

# удаление старых файлов
$tmpFilePath = Join-Path -Path $tmpPath -ChildPath '*.*'
if (Test-Path $tmpFilePath) {
    Remove-Item $tmpFilePath
}
$ExcelPath = Join-Path -Path $tmpPath -ChildPath "summary.xlsx"

Export-ExcelSummary -Path $ExcelPath -Diagnose "PH" -outData $phData
Export-ExcelSummary -Path $ExcelPath -Diagnose "PC" -outData $pcData