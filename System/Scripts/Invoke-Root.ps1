[CmdletBinding()]
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

$ranges = Import-Norm-Ranges

$fields = @(
  [PSCustomObject]@{in = "Number"; out = "Number"; },
  [PSCustomObject]@{in = "Peak_Name"; out = "Name"; },
  [PSCustomObject]@{in = "Amount"; out = "Amount"; }
)

Import-XmlDataMult -Path $inPath -tmpPath $tmpPath -Format "AlexPasha" -fields ($fields  | Select-Object -ExpandProperty "in")

$qwe = 123

# Импорт корректировки из Excel и формирование correct.sample.json
# . .\Invoke-PostCorrect.ps1 -inPath $inPath

# Формирование таблиц 1.2 + (TableData1.json + TableData2.json)
# . .\Invoke-CommonStat.ps1 -inPath $inPath

# . .\Invoke-PkStat.ps1 -inPath $inPath
