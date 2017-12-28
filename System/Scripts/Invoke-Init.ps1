[CmdletBinding()]
Param(
  [Parameter(Mandatory = $False)]
  [string]$inPath
)
# $OutputEncoding = [Console]::OutputEncoding
# $OutputEncoding = [system.Text.Encoding]::UTF8
# $OutputEncoding = [System.Text.Encoding]::Default

$ErrorActionPreference = "Stop"
$scriptPath = Split-Path $MyInvocation.MyCommand.path
$systemPath = Split-Path -Parent $scriptPath
if (!$inPath) {
  $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data' 
}

$normPath = Split-Path -Parent $systemPath | Join-Path -ChildPath "Normal concentrations blood.xlsx" 

$tmpPath = $inPath | Join-Path -ChildPath 'result'
if ($isCleanStart -and (Test-Path $tmpPath)) {
  Remove-Item $tmpPath -Recurse
}
if (!(Test-Path $tmpPath)) {
  New-Item -ItemType Directory -Path $tmpPath | Out-Null
}

# $tmpExcelPath = Join-Path -Path $tmpPath -ChildPath 'out.xlsx'
$EpPlusPath = Split-Path -Parent $scriptPath | Join-Path -ChildPath "packages\EPPlus 4.1\lib" | Join-Path -ChildPath EPPlus.dll 
[Reflection.Assembly]::LoadFrom($EpPlusPath)# | Out-Null

. .\excel-helpers.ps1

$BioFluidNames = @("Serum")
# $BioFluidNames = @("Serum", "Saliva", "Urine")
$StopListName = "Stop List"

# загрузка диапазонов по всем жидкостям
function Import-Norm-Ranges () {
#   [CmdletBinding()]
#   Param(
#     [Parameter(Mandatory = $True)] [string]$Path # путь к файлу с данными
#   )
  
  $ranges = @{}

  $BioFluidNames.ForEach( {
      $item = Import-BioFluid-Ranges -Path $normPath -Name $_
      $ranges.Add($_, $item)
    })

    $ranges
}