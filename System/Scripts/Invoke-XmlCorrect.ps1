[CmdletBinding()]
Param(
    [Parameter(Mandatory=$False)] [string]$inPath,
    [Parameter(Mandatory=$False)] [string]$outPath
)

$ErrorActionPreference = "Stop"

# $scriptPath = Split-Path $MyInvocation.MyCommand.path
# if (!$inPath) {
#     $systemPath = Split-Path -Parent $scriptPath
    # $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\All' 
    # $inPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\All'
    # $outPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Amino Acids' 
# }
$inPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\All'

. ".\xml-helpers.ps1"

$outPath = 'z:\UBUNTU\02 - Проекты в работе\32 - Воспроизводимость после хранения\Павел Prostate Cancer\Amino Acids\Urine' 
Correct-XmlDataMult -inPath $inPath -outPath $outPath -corrFunc "Set_Creatinine_Element"

# $outPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Fatty Acids_All' 
# Correct-XmlDataMult -inPath $inPath -outPath $outPath -corrFunc "Set_Creatinine_Element"

# $outPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data\Fatty Acids_Free' 
# Correct-XmlDataMult -inPath $inPath -outPath $outPath -corrFunc "Set_Creatinine_Element"