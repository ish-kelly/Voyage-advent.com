Add-Type -AssemblyName System.IO.Compression.FileSystem

$xlsx = 'C:\Users\kelly\Downloads\website_product_catalog.xlsx'
$outDir = 'C:\Users\kelly\Downloads\Voyage_Advent_COMPLETE_FIXED\assets\product-images'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsx)

# Parse Products sheet rows
$sheetEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
$sr = New-Object System.IO.StreamReader($sheetEntry.Open())
$sheetXml = $sr.ReadToEnd()
$sr.Dispose()
[xml]$sheetDoc = $sheetXml
$sheetNs = New-Object System.Xml.XmlNamespaceManager -ArgumentList $sheetDoc.NameTable
$sheetNs.AddNamespace('a', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

$rowsByNumber = @{}
foreach ($row in $sheetDoc.SelectNodes('//a:worksheet/a:sheetData/a:row', $sheetNs)) {
    $rowNumber = [int]($row.GetAttribute('r'))
    $cells = @{}
    foreach ($cell in $row.SelectNodes('a:c', $sheetNs)) {
        $ref = $cell.GetAttribute('r')
        $col = ($ref -replace '\d', '')
        $type = $cell.GetAttribute('t')
        if ($type -eq 'inlineStr') {
            $value = ($cell.SelectSingleNode('a:is/a:t', $sheetNs)).InnerText
        }
        else {
            $value = $cell.InnerText
        }
        $cells[$col] = $value
    }

    if ($cells.ContainsKey('A') -and $cells['A'] -match 'P\d+' -and $cells.ContainsKey('B')) {
        $rowsByNumber[$rowNumber] = [pscustomobject]@{
            ProductId = $cells['A']
            ProductName = $cells['B']
            ImageFilename = $cells['H']
            Category = $cells['C']
        }
    }
}

# Parse drawing to row-to-image relationship ids
$drawingEntry = $zip.GetEntry('xl/drawings/drawing1.xml')
$drSr = New-Object System.IO.StreamReader($drawingEntry.Open())
$drawingXml = $drSr.ReadToEnd()
$drSr.Dispose()
[xml]$drawingDoc = $drawingXml
$drawingNs = New-Object System.Xml.XmlNamespaceManager -ArgumentList $drawingDoc.NameTable
$drawingNs.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
$drawingNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

$relsEntry = $zip.GetEntry('xl/drawings/_rels/drawing1.xml.rels')
$rSr = New-Object System.IO.StreamReader($relsEntry.Open())
$relsXml = $rSr.ReadToEnd()
$rSr.Dispose()
[xml]$relsDoc = $relsXml
$relsNs = New-Object System.Xml.XmlNamespaceManager -ArgumentList $relsDoc.NameTable
$relsNs.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')

$targetById = @{}
foreach ($rel in $relsDoc.SelectNodes('//rel:Relationship', $relsNs)) {
    $id = $rel.GetAttribute('Id')
    $target = $rel.GetAttribute('Target')
    $targetById[$id] = $target.TrimStart('/')
}

$map = @{}
foreach ($anchor in $drawingDoc.SelectNodes('//a:oneCellAnchor', $drawingNs)) {
    $from = $anchor.SelectSingleNode('a:from', $drawingNs)
    $rowNumber = [int]($from.SelectSingleNode('a:row', $drawingNs).InnerText)
    $pic = $anchor.SelectSingleNode('a:pic', $drawingNs)
    $blip = $pic.SelectSingleNode('a:blipFill/a:blip', $drawingNs)
    $relId = $blip.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $target = $targetById[$relId]
    $mediaName = Split-Path $target -Leaf

    if ($rowsByNumber.ContainsKey($rowNumber)) {
        $product = $rowsByNumber[$rowNumber]
        $destName = if ([string]::IsNullOrWhiteSpace($product.ImageFilename)) { $mediaName } else { $product.ImageFilename }
        $destPath = Join-Path $outDir $destName

        $entry = $zip.GetEntry($target)
        if ($entry) {
            $stream = $entry.Open()
            $fs = [System.IO.File]::Open($destPath, [System.IO.FileMode]::Create)
            $stream.CopyTo($fs)
            $fs.Close()
            $stream.Close()
        }

        $map[$product.ProductName] = $destName
        Write-Host "$($product.ProductId)|$($product.ProductName)|$destName"
    }
}

$zip.Dispose()
Write-Host "Images extracted to: $outDir"
