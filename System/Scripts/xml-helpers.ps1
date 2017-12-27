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

