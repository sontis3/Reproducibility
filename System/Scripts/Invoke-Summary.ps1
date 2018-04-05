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
}

#-------------------------------------------------------------------------------
# загрузка драйвера Excel и нормальных диапазонов
. .\Invoke-Init.ps1 -inPath $inPath
# поддержка XML
. ".\xml-helpers.ps1"

#-----------------------------------------------------------------------------------------------

function Get-XmlDataMult() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)] [string]$Path, # каталог с исходными файлами
        [Parameter(Mandatory = $True)] [string]$tmpPath, # каталог временных(результат) файлов
        [Parameter(Mandatory = $True)] [string]$Format, # формат данных
        [Parameter(Mandatory = $True)] [System.Object[]]$fields, # представление данных in/out
        [Parameter(Mandatory = $True)] [string]$filterTemplate              # шаблон фильтра
    )
  
    Write-Host '-------------- start job (Load from multiple xml to stage)---------------'
    # удаление старых файлов
    $tmpFilePath = Join-Path -Path $tmpPath -ChildPath '*.*'
    if (Test-Path $tmpFilePath) {
        Remove-Item $tmpFilePath
    }

    $tableData = @()

    Get-ChildItem -Path $Path -Filter $filterTemplate | Sort-Object -Property Name |
        ForEach-Object {
        Write-Host '-------------- ' + $_.Name + ' ---------------'
        # switch -regex ($_.Name) {
        #     $BioFluidNames[0] { $selFluidName = $BioFluidNames[0] }
        #     $BioFluidNames[1] { $selFluidName = $BioFluidNames[1] }
        #     $BioFluidNames[2] { $selFluidName = $BioFluidNames[2] }
        #     Default {
        #         Write-Warning ("Не найден код жидкости для " + $_.Name)
        #         $selFluidName = $null
        #     }
        # }
  
        # if ($selFluidName -and $ranges.$selFluidName) {
            $inData = Get-XmlData -FilePath $_.FullName -Format $Format -Filter " " -fields ($fields  | Select-Object -ExpandProperty "in")
  
        #     $outData = Process-Data -fluidName $selFluidName -ranges $ranges.$selFluidName -stopList $stopList -indata $inData
  
            $baseName = $_ | Select-Object -ExpandProperty BaseName
        #     $fileName = $baseName + ".xlsx"
        #     $ExcelPath = Join-Path -Path $tmpPath -ChildPath $fileName
            $dataInfo = $baseName.Split("_")

            $item = [PSCustomObject]@{
                name = $dataInfo[1];
                inData = $inData;
            }

            $tableData += $item

        #     Export-ExcelData -Path $ExcelPath -dataInfo $dataInfo -outData $outData
        # }
    }
  
    Write-Host '-------------- end job (Load from multiple xml to stage)---------------'
    
    $tableData
}
  

$fields = @(
    [PSCustomObject]@{in = "Number"; out = "Number"; },
    [PSCustomObject]@{in = "Peak_Name"; out = "Name"; },
    [PSCustomObject]@{in = "Amount"; out = "Amount"; }
)

$phData = Get-XmlDataMult -Path $inPath -tmpPath $tmpPath -Format "AlexPasha" -fields $fields -filterTemplate "*_PH*.xml"
$pcData = Get-XmlDataMult -Path $inPath -tmpPath $tmpPath -Format "AlexPasha" -fields $fields -filterTemplate "*_PC*.xml"
