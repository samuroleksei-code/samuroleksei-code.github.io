$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$previewRoot = Join-Path $root "previews"
$assetsRoot = Join-Path $previewRoot "assets"
$wordNamespaceUri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$relationshipNamespaceUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Escape-Html {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Read-ZipXml {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) {
        return $null
    }

    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)

    try {
        [xml]$xml = $reader.ReadToEnd()
        return $xml
    }
    finally {
        $reader.Close()
        $stream.Close()
    }
}

function New-NamespaceManager {
    param([xml]$XmlDocument)

    $ns = New-Object System.Xml.XmlNamespaceManager($XmlDocument.NameTable)
    [void]$ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    [void]$ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    [void]$ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    [void]$ns.AddNamespace("v", "urn:schemas-microsoft-com:vml")
    return $ns
}

function Get-RelationshipMap {
    param([xml]$RelationshipsXml)

    $map = @{}
    if (-not $RelationshipsXml) {
        return $map
    }

    foreach ($node in $RelationshipsXml.Relationships.Relationship) {
        $map[$node.Id] = $node.Target
    }

    return $map
}

function Get-NodeTextHtml {
    param(
        [System.Xml.XmlNode]$Node,
        $NamespaceManager
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $inlineNodes = $Node.SelectNodes(".//*[local-name()='t' or local-name()='tab' or local-name()='br' or local-name()='cr' or local-name()='noBreakHyphen' or local-name()='softHyphen']")

    foreach ($inlineNode in $inlineNodes) {
        switch ($inlineNode.LocalName) {
            "t" { $parts.Add((Escape-Html $inlineNode.InnerText)) }
            "tab" { $parts.Add("&emsp;") }
            "br" { $parts.Add("<br>") }
            "cr" { $parts.Add("<br>") }
            "noBreakHyphen" { $parts.Add("-") }
            "softHyphen" { $parts.Add("-") }
        }
    }

    return ($parts -join "")
}

function Get-ParagraphImageHtml {
    param(
        [System.Xml.XmlNode]$Paragraph,
        $NamespaceManager,
        [hashtable]$RelationshipMap,
        [string]$AssetRelativePath
)

    $images = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $blips = $Paragraph.SelectNodes(".//*[local-name()='blip']")
    foreach ($blip in $blips) {
        $embed = $blip.GetAttribute("embed", $relationshipNamespaceUri)
        if ($embed -and $RelationshipMap.ContainsKey($embed)) {
            $target = [System.IO.Path]::GetFileName($RelationshipMap[$embed])
            if ($target -and -not $seen.ContainsKey($target)) {
                $seen[$target] = $true
                $images.Add("<figure class=`"doc-image`"><img src=`"$AssetRelativePath/$target`" alt=`"Зображення з документа`"></figure>")
            }
        }
    }

    $legacyImages = $Paragraph.SelectNodes(".//*[local-name()='imagedata']")
    foreach ($imageNode in $legacyImages) {
        $embed = $imageNode.GetAttribute("id", $relationshipNamespaceUri)
        if ($embed -and $RelationshipMap.ContainsKey($embed)) {
            $target = [System.IO.Path]::GetFileName($RelationshipMap[$embed])
            if ($target -and -not $seen.ContainsKey($target)) {
                $seen[$target] = $true
                $images.Add("<figure class=`"doc-image`"><img src=`"$AssetRelativePath/$target`" alt=`"Зображення з документа`"></figure>")
            }
        }
    }

    return $images
}

function Convert-Paragraph {
    param(
        [System.Xml.XmlNode]$Paragraph,
        $NamespaceManager,
        [hashtable]$RelationshipMap,
        [string]$AssetRelativePath
)

    $textHtml = Get-NodeTextHtml -Node $Paragraph -NamespaceManager $NamespaceManager
    $images = Get-ParagraphImageHtml -Paragraph $Paragraph -NamespaceManager $NamespaceManager -RelationshipMap $RelationshipMap -AssetRelativePath $AssetRelativePath

    $styleNode = $Paragraph.SelectSingleNode("./*[local-name()='pPr']/*[local-name()='pStyle']")
    $alignmentNode = $Paragraph.SelectSingleNode("./*[local-name()='pPr']/*[local-name()='jc']")
    $listNode = $Paragraph.SelectSingleNode("./*[local-name()='pPr']/*[local-name()='numPr']")

    $styleValue = if ($styleNode) { $styleNode.GetAttribute("val", $wordNamespaceUri) } else { "" }
    $alignment = if ($alignmentNode) { $alignmentNode.GetAttribute("val", $wordNamespaceUri) } else { "" }

    if ($listNode -and $textHtml) {
        $textHtml = "&#8226; " + $textHtml
    }

    $tag = "p"
    switch -Regex ($styleValue) {
        "Title|Heading1" { $tag = "h1"; break }
        "Heading2|Subtitle" { $tag = "h2"; break }
        "Heading3" { $tag = "h3"; break }
        "Heading4" { $tag = "h4"; break }
    }

    $classes = New-Object System.Collections.Generic.List[string]
    if ($alignment -eq "center") {
        $classes.Add("is-centered")
    }
    elseif ($alignment -eq "right") {
        $classes.Add("is-right")
    }

    $markup = New-Object System.Collections.Generic.List[string]
    $plainText = $textHtml -replace "<br>", "" -replace "&emsp;", "" -replace "&#8226;", "" -replace "\s+", ""

    if (-not [string]::IsNullOrWhiteSpace($plainText)) {
        $classAttribute = if ($classes.Count -gt 0) { " class=`"" + ($classes -join " ") + "`"" } else { "" }
        $markup.Add("<$tag$classAttribute>$textHtml</$tag>")
    }

    foreach ($imageMarkup in $images) {
        $markup.Add($imageMarkup)
    }

    return ($markup -join "`n")
}

function Convert-Table {
    param(
        [System.Xml.XmlNode]$Table,
        $NamespaceManager
    )

    $rowsHtml = New-Object System.Collections.Generic.List[string]
    $rows = $Table.SelectNodes("./*[local-name()='tr']")

    foreach ($row in $rows) {
        $cellHtml = New-Object System.Collections.Generic.List[string]
        $cells = $row.SelectNodes("./*[local-name()='tc']")

        foreach ($cell in $cells) {
            $paragraphs = $cell.SelectNodes("./*[local-name()='p']")
            $pieces = New-Object System.Collections.Generic.List[string]

            foreach ($paragraph in $paragraphs) {
                $textHtml = Get-NodeTextHtml -Node $paragraph -NamespaceManager $NamespaceManager
                $plainText = $textHtml -replace "<br>", "" -replace "&emsp;", "" -replace "\s+", ""

                if (-not [string]::IsNullOrWhiteSpace($plainText)) {
                    $pieces.Add("<p>$textHtml</p>")
                }
            }

            if ($pieces.Count -eq 0) {
                $pieces.Add("<p>&nbsp;</p>")
            }

            $cellHtml.Add("<td>" + ($pieces -join "") + "</td>")
        }

        $rowsHtml.Add("<tr>" + ($cellHtml -join "") + "</tr>")
    }

    return "<div class=`"table-wrap`"><table>" + ($rowsHtml -join "") + "</table></div>"
}

function Export-MediaFiles {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$TargetDirectory
    )

    Ensure-Directory -Path $TargetDirectory

    foreach ($entry in $Zip.Entries) {
        if ($entry.FullName -like "word/media/*") {
            $targetName = [System.IO.Path]::GetFileName($entry.FullName)
            $targetPath = Join-Path $TargetDirectory $targetName

            $sourceStream = $entry.Open()
            $fileStream = [System.IO.File]::Create($targetPath)

            try {
                $sourceStream.CopyTo($fileStream)
            }
            finally {
                $fileStream.Close()
                $sourceStream.Close()
            }
        }
    }
}

function Convert-DocxToPreview {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [string]$Title,
        [string]$Subtitle
    )

    $zip = [System.IO.Compression.ZipFile]::OpenRead($SourcePath)

    try {
        [xml]$documentXml = Read-ZipXml -Zip $zip -EntryName "word/document.xml"
        [xml]$relationshipsXml = Read-ZipXml -Zip $zip -EntryName "word/_rels/document.xml.rels"
        $namespaceManager = New-NamespaceManager -XmlDocument $documentXml
        $relationshipMap = Get-RelationshipMap -RelationshipsXml $relationshipsXml

        $slug = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
        $assetDirectory = Join-Path $assetsRoot $slug
        $assetRelativePath = "assets/$slug"

        Ensure-Directory -Path $previewRoot
        Export-MediaFiles -Zip $zip -TargetDirectory $assetDirectory

        $content = New-Object System.Collections.Generic.List[string]
        $documentNode = [System.Xml.XmlNode]$documentXml.DocumentElement
        $bodyNode = $documentNode.ChildNodes | Where-Object { $_.LocalName -eq "body" } | Select-Object -First 1
        $bodyNodes = if ($bodyNode) { $bodyNode.ChildNodes } else { @() }

        foreach ($node in $bodyNodes) {
            switch ($node.LocalName) {
                "p" {
                    $paragraphMarkup = Convert-Paragraph -Paragraph $node -NamespaceManager $namespaceManager -RelationshipMap $relationshipMap -AssetRelativePath $assetRelativePath
                    if (-not [string]::IsNullOrWhiteSpace($paragraphMarkup)) {
                        $content.Add($paragraphMarkup)
                    }
                }
                "tbl" {
                    $content.Add((Convert-Table -Table $node -NamespaceManager $namespaceManager))
                }
            }
        }

        $html = @"
<!DOCTYPE html>
<html lang="uk">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Title</title>
    <link rel="stylesheet" href="preview.css">
</head>
<body>
    <div class="preview-page">
        <header class="preview-header">
            <p class="preview-eyebrow">Онлайн-перегляд документа</p>
            <h1>$Title</h1>
            <p class="preview-subtitle">$Subtitle</p>
        </header>
        <main class="preview-content">
$($content -join "`n")
        </main>
    </div>
</body>
</html>
"@

        Set-Content -Path $OutputPath -Value $html -Encoding UTF8
    }
    finally {
        $zip.Dispose()
    }
}

Ensure-Directory -Path $previewRoot
Ensure-Directory -Path $assetsRoot

$documents = @(
    @{
        Source = (Join-Path $root "documents/characteristic.docx")
        Output = (Join-Path $previewRoot "characteristic.html")
        Title = "Характеристика студента 4 курсу"
        Subtitle = "Автоматично згенерований перегляд Word-документа для сайту практики."
    },
    @{
        Source = (Join-Path $root "documents/plan.docx")
        Output = (Join-Path $previewRoot "plan.html")
        Title = "Індивідуальний план"
        Subtitle = "Автоматично згенерований перегляд основного плану проходження практики."
    },
    @{
        Source = (Join-Path $root "documents/report.docx")
        Output = (Join-Path $previewRoot "report.html")
        Title = "Звіт з практики"
        Subtitle = "Автоматично згенерований перегляд підсумкового звіту з виробничої практики."
    }
)

foreach ($document in $documents) {
    Convert-DocxToPreview `
        -SourcePath $document.Source `
        -OutputPath $document.Output `
        -Title $document.Title `
        -Subtitle $document.Subtitle
}

Get-ChildItem -Path $previewRoot -File | Select-Object Name, Length
