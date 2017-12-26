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

#-----------------------------------------------------------------------------------------------
function Import-XmlDataMult() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)] [string]$Path,                        # каталог с исходными файлами
        [Parameter(Mandatory=$True)] [string]$tmpPath,                     # каталог временных(результат) файлов
        [Parameter(Mandatory=$True)]  [string]$Format,                     # формат данных
        [Parameter(Mandatory=$True)]  [System.Object]$fields
    )

    Write-Host '-------------- start job (Load from multiple xml to stage)---------------'
    # удаление старых файлов
    $tmpFilePath = Join-Path -Path $tmpPath -ChildPath '*.rem'
    if(Test-Path $tmpFilePath){
        Remove-Item $tmpFilePath
    }

    $xmlInput = @()
    Get-ChildItem -Path $Path -Filter *.xml |
        ForEach-Object {
            $xmlInput += Get-XmlData -FilePath $_.FullName -Format $Format -Filter " " -fields $fields
        }

    Write-Host '-------------- end job (Load from multiple csv to stage)---------------'
}

#------------------------------- Корректировка файлов XML
# корректровка Inpath4
function Correct-XmlSingle () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)] [string]$FilePath                       # исходный файл
    )
    $fileName = Split-Path $FilePath -Leaf
    $parts = $fileName.Split("-")
    $cycle = $parts[-2]
    [xml] $xmlSample = Get-Content $FilePath
    $samples = $xmlSample.SelectNodes("//QUANDATASET/GROUPDATA/GROUP/SAMPLELISTDATA/SAMPLE")
    # $xmlSample.QUANDATASET.GROUPDATA.GROUP.SAMPLELISTDATA.SAMPLE | ForEach-Object {
    #     $_.Attributes.
    # }
    $qcLevels = "LLOQQC", "LQC", "MQC", "HQC"
    $i = 0
    $j = 1
    foreach ($sample in $samples) {
        $sample.task = $cycle                       # простановка номера цикла выделенного из имени файла
        if ($sample.type -eq "Analyte") {
            $parts = $sample.name.Split("-")
            switch ($parts[-3]) {
                "01" { $parts[-3] = "1-17" }
                "02" { $parts[-3] = "1-18" }
                "03" { $parts[-3] = "1-19" }
                "04" { $parts[-3] = "1-20" }
                "05" { $parts[-3] = "1-21" }
                "06" { $parts[-3] = "1-22" }
                "07" { $parts[-3] = "1-23" }
                "08" { $parts[-3] = "1-24" }
                "09" { $parts[-3] = "1-01" }
                "10" { $parts[-3] = "1-02" }
                "11" { $parts[-3] = "1-03" }
                "12" { $parts[-3] = "1-04" }
                "13" { $parts[-3] = "1-05" }
                "14" { $parts[-3] = "1-06" }
                "15" { $parts[-3] = "1-07" }
                "16" { $parts[-3] = "1-08" }
                "17" { $parts[-3] = "1-09" }
                "18" { $parts[-3] = "1-10" }
                "19" { $parts[-3] = "1-11" }
                "20" { $parts[-3] = "1-12" }
                "21" { $parts[-3] = "1-13" }
                "22" { $parts[-3] = "1-14" }
                "23" { $parts[-3] = "1-15" }
                "24" { $parts[-3] = "1-16" }
                Default {Write-Host ('ERROR: Нераспознанный номер ' + $sample.name)}
            }
            $sample.name = $parts -join "-"
        }
        if ($sample.type -eq "QC") {
            $sample.name = $qcLevels[$i++] + "-$j"
            if ($i -ge $qcLevels.Count) {
                $i = 0
                $j += 1
            }
        }
    }

    $xmlSample.Save(($FilePath + "_c"))
}

function Correct-24 () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)] [string]$FilePath                       # исходный файл
    )
    $fileName = Split-Path $FilePath -Leaf
    $parts = $fileName.Split(".")
    $parts = $parts[0].Split("-")
    $cycle = $parts[-1]
    [xml] $xmlSample = Get-Content $FilePath
    $samples = $xmlSample.SelectNodes("//QUANDATASET/GROUPDATA/GROUP/SAMPLELISTDATA/SAMPLE")
    foreach ($sample in $samples) {
        $parts = $sample.name.Split("-")
        if ($sample.type -eq "Analyte") {
            switch ($parts[-4]) {
                "2" { $parts[-4] = "1" }
                Default {
                    Write-Host ('FILE:  ' + $FilePath)
                    Write-Host ('ERROR: Нераспознанный номер ' + $sample.name)
                }
            }
            $sample.name = $parts -join "-"
        } elseif ($sample.type -eq "Standard") {
            switch ($parts[-2]) {
                "LLOQC" {
                    $parts[-2] = "LLOQQC"
                    $sample.type = "QC"
                }
            }
            $sample.name = $parts -join "-"
        }
    }

    $xmlSample.Save(($FilePath + ".c"))
}

function Correct-LLOQC () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)] [string]$FilePath                       # исходный файл
    )
    $fileName = Split-Path $FilePath -Leaf
    $parts = $fileName.Split(".")
    $parts = $parts[0].Split("-")
    $cycle = $parts[-1]
    [xml] $xmlSample = Get-Content $FilePath
    $samples = $xmlSample.SelectNodes("//QUANDATASET/GROUPDATA/GROUP/SAMPLELISTDATA/SAMPLE")
    foreach ($sample in $samples) {
        $parts = $sample.name.Split("-")
        if ($sample.type -eq "QC") {
            switch ($parts[-2]) {
                "LLOQC" {
                    $parts[-2] = "LLOQQC"
                }
            }
            $sample.name = $parts -join "-"
        }
    }

    $xmlSample.Save(($FilePath + ".c"))
}

function Correct-XmlDataMult () {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)] [string]$Path,                    # каталог с исходными файлами
        [Parameter(Mandatory=$False)] [string]$corrFunc                 # ф-ция корректровки
    )

    Write-Host '-------------- start job (Correct multiple xml)---------------'
    # исходный файл
    if (!$Path) {
        # если нет исходного файла в параметрах
        $scriptPath = Split-Path $script:MyInvocation.MyCommand.path
        $systemPath = Split-Path -Parent $scriptPath
        $Path = Split-Path -Parent $systemPath | Join-Path -ChildPath 'In-Data' 
    }

    Get-ChildItem -Path $Path -Filter "*.xml" |
        ForEach-Object {
            &($corrFunc) -FilePath $_.FullName
        }

    Write-Host '-------------- end job (Correct multiple xml)---------------'

}

