# получить путь по фильтру
function AlexPasha () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)] [string]$Filter            # фильтр
    )

    ("//Results/Compounds")
}

# выбрать данные из xml
function Get-XmlData () { 
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)] [string]$FilePath,                     # исходный файл
        [Parameter(Mandatory=$True)] [string]$Format,                       # формат данных
        [Parameter(Mandatory=$True)] [string]$Filter,                       # фильтр
        [Parameter(Mandatory=$True)] [System.Object]$fields
    )
    
    # создание шаблона объекта
    $si = New-Object psobject
    foreach ($p in $fields) {
        $si | Add-Member -MemberType NoteProperty -Name $p -Value "0"
    }

    $result = @()
    [xml] $xmlSample = Get-Content $FilePath

    $nsUri = $xmlSample.DocumentElement.NamespaceURI
    $nsm = New-Object Xml.XmlNamespaceManager $xmlSample.NameTable
    $nsm.AddNamespace("c", $nsUri)

    $rc = $xmlSample.SelectNodes((&($Format) -Filter $Filter), $nsm)
    
    foreach ($item in $rc) {
        $sampleItem = $si.psobject.Copy()
        foreach ($f in $fields) {
            $srcField = $item
            $srcField = $srcField.SelectSingleNode($f, $nsm)
            $sampleItem.$f = $srcField.InnerText
        }

        $result += $sampleItem
        # $result += [PSCustomObject]@{           # !! такой тип для правильной обработки полей
        #     name = $item.name;
        #     type = $item.type;
        #     stdconc = $item.stdconc;
        #     analconc = $item.COMPOUND.PEAK.analconc
        # }
    }
    $result
}

function Set_Creatinine_Element () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)] [string]$FilePath,                      # исходный файл
        [Parameter(Mandatory=$True)] [System.Object]$elt                     # xml-элемент с данными по creatinine
    )
    [xml] $xmlSample = Get-Content $FilePath

    if (!$xmlSample.Results.Compounds.Where({$_.Peak_Name -eq "Creatinine"})[0]) {
        $xmlSample.Save(($FilePath + ".old"))

        $num = $xmlSample.Results.Compounds.Count
        $z = $xmlSample.ImportNode($elt, $True)
        $z.Number = ($num + 1).ToString()
        $result = $xmlSample.Results.AppendChild($z)
    
        $xmlSample.Save(($FilePath))
    }
}

function Correct-XmlDataMult () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)] [string]$inPath,                  # каталог с исходными файлами
        [Parameter(Mandatory=$False)] [string]$outPath,                 # каталог с корректруемыми файлами
        [Parameter(Mandatory=$False)] [string]$corrFunc                 # ф-ция корректровки
    )

    Write-Host '-------------- start job (Correct multiple xml)---------------'
    # исходный файл
    if (!$inPath) {
        # если нет исходного файла в параметрах
        $scriptPath = Split-Path $script:MyInvocation.MyCommand.path
        $systemPath = Split-Path -Parent $scriptPath
        $inPath = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data' 
    }

    Get-ChildItem -Path $outPath -Filter "Urine*.xml" |
        ForEach-Object {
            $outFilePath = $_.FullName
            Write-Host "-------------- $outFilePath ---------------"
            $baseName = $_ | Select-Object -ExpandProperty BaseName
            $dataInfo = $baseName.Split("_")
            $identNamePart = $dataInfo[1..($dataInfo.Count-2)] -join "_"
            Get-ChildItem -Path $inPath -Filter "Urine*$identNamePart*.xml" |
                ForEach-Object {
                    $inFilePath = $_.FullName
                }

            [xml] $xmlSample = Get-Content $inFilePath
            $sss = $xmlSample.Results.Compounds.Where({$_.Peak_Name -eq "Creatinine"})[0]

            &($corrFunc) -FilePath $outFilePath -elt $sss
        }

    Write-Host '-------------- end job (Correct multiple xml)---------------'

}
