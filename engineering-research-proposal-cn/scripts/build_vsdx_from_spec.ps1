param(
    [Parameter(Mandatory=$true)][string]$InputJson,
    [Parameter(Mandatory=$true)][string]$OutputVsdx,
    [string]$PreviewSvg,
    [string]$PreviewPng,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputJson).Path
$outputPath = [IO.Path]::GetFullPath($OutputVsdx)
if ([IO.Path]::GetExtension($outputPath) -ne '.vsdx') {
    throw 'OutputVsdx must use the .vsdx extension.'
}
if ((Test-Path -LiteralPath $outputPath) -and -not $Force) {
    throw "Refusing to overwrite existing file: $outputPath"
}

function Convert-HexToRgbFormula([string]$Hex) {
    if ($Hex -notmatch '^#?[0-9A-Fa-f]{6}$') {
        throw "Invalid RGB color: $Hex"
    }
    $value = $Hex.TrimStart('#')
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return "RGB($r,$g,$b)"
}

function Set-FormulaIfPresent($Shape, [string]$CellName, [string]$Formula) {
    if ($Shape.CellExistsU($CellName, 0) -ne 0) {
        $Shape.CellsU($CellName).FormulaU = $Formula
    }
}

$spec = Get-Content -Raw -Encoding UTF8 -LiteralPath $inputPath | ConvertFrom-Json
if (-not $spec.nodes -or $spec.nodes.Count -eq 0) {
    throw 'The diagram specification must contain at least one node.'
}

$outputDirectory = [IO.Path]::GetDirectoryName($outputPath)
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$app = $null
$doc = $null
$page = $null
try {
    $app = New-Object -ComObject Visio.Application
    $app.Visible = $false
    $app.AlertResponse = 7
    $doc = $app.Documents.Add('')
    $page = $app.ActivePage

    $pageWidth = if ($spec.page.width_in) { [double]$spec.page.width_in } else { 11.69 }
    $pageHeight = if ($spec.page.height_in) { [double]$spec.page.height_in } else { 8.27 }
    $page.PageSheet.CellsU('PageWidth').FormulaU = "$pageWidth in"
    $page.PageSheet.CellsU('PageHeight').FormulaU = "$pageHeight in"
    Set-FormulaIfPresent $page.PageSheet 'PageLeftMargin' '0.25 in'
    Set-FormulaIfPresent $page.PageSheet 'PageRightMargin' '0.25 in'
    Set-FormulaIfPresent $page.PageSheet 'PageTopMargin' '0.25 in'
    Set-FormulaIfPresent $page.PageSheet 'PageBottomMargin' '0.25 in'

    $defaultFont = if ($spec.style.font) { [string]$spec.style.font } else { 'Microsoft YaHei' }
    $defaultFontSize = if ($spec.style.font_size_pt) { [double]$spec.style.font_size_pt } else { 10 }
    $defaultFill = if ($spec.style.fill) { [string]$spec.style.fill } else { '#D9EAF7' }
    $defaultLine = if ($spec.style.line) { [string]$spec.style.line } else { '#2F5597' }
    $defaultText = if ($spec.style.text) { [string]$spec.style.text } else { '#1F1F1F' }
    $lineWeight = if ($spec.style.line_weight_pt) { [double]$spec.style.line_weight_pt } else { 1.25 }

    $shapes = @{}
    foreach ($node in $spec.nodes) {
        $id = [string]$node.id
        if ([string]::IsNullOrWhiteSpace($id) -or $shapes.ContainsKey($id)) {
            throw "Node ids must be non-empty and unique: $id"
        }
        $x = [double]$node.x
        $y = [double]$node.y
        $width = if ($node.width) { [double]$node.width } else { 2.0 }
        $height = if ($node.height) { [double]$node.height } else { 0.8 }
        $shape = $page.DrawRectangle($x - $width / 2, $y - $height / 2, $x + $width / 2, $y + $height / 2)
        $shape.Text = [string]$node.text

        $fill = if ($node.fill) { [string]$node.fill } else { $defaultFill }
        $line = if ($node.line) { [string]$node.line } else { $defaultLine }
        $textColor = if ($node.text_color) { [string]$node.text_color } else { $defaultText }
        $fontSize = if ($node.font_size_pt) { [double]$node.font_size_pt } else { $defaultFontSize }

        $shape.CellsU('FillForegnd').FormulaU = Convert-HexToRgbFormula $fill
        $shape.CellsU('FillPattern').FormulaU = '1'
        $shape.CellsU('LineColor').FormulaU = Convert-HexToRgbFormula $line
        $shape.CellsU('LineWeight').FormulaU = "$lineWeight pt"
        Set-FormulaIfPresent $shape 'Rounding' '0.08 in'
        Set-FormulaIfPresent $shape 'Char.Font' "FONT(`"$defaultFont`")"
        Set-FormulaIfPresent $shape 'Char.Size' "$fontSize pt"
        Set-FormulaIfPresent $shape 'Char.Color' (Convert-HexToRgbFormula $textColor)
        Set-FormulaIfPresent $shape 'Para.HorzAlign' '1'
        Set-FormulaIfPresent $shape 'VerticalAlign' '1'
        Set-FormulaIfPresent $shape 'TxtWidth' 'Width-0.12 in'
        Set-FormulaIfPresent $shape 'TxtHeight' 'Height-0.08 in'
        $shapes[$id] = $shape
    }

    $connectors = @()
    foreach ($edge in $spec.connectors) {
        $fromId = [string]$edge.from
        $toId = [string]$edge.to
        if (-not $shapes.ContainsKey($fromId) -or -not $shapes.ContainsKey($toId)) {
            throw "Connector references an unknown node: $fromId -> $toId"
        }
        $connector = $page.Drop($app.ConnectorToolDataObject, 0, 0)
        $connector.CellsU('BeginX').GlueTo($shapes[$fromId].CellsU('PinX'))
        $connector.CellsU('EndX').GlueTo($shapes[$toId].CellsU('PinX'))
        $connector.CellsU('LineColor').FormulaU = Convert-HexToRgbFormula $defaultLine
        $connector.CellsU('LineWeight').FormulaU = "$lineWeight pt"
        Set-FormulaIfPresent $connector 'EndArrow' '13'
        Set-FormulaIfPresent $connector 'ConLineRouteExt' '0'
        Set-FormulaIfPresent $connector 'ShapeRouteStyle' '1'
        Set-FormulaIfPresent $connector 'Char.Font' "FONT(`"$defaultFont`")"
        Set-FormulaIfPresent $connector 'Char.Size' "$defaultFontSize pt"
        if ($edge.label) {
            $connector.Text = [string]$edge.label
        }
        $connectors += $connector
    }

    $doc.SaveAs($outputPath)
    $previewPath = $null
    if (-not [string]::IsNullOrWhiteSpace($PreviewSvg)) {
        $previewPath = [IO.Path]::GetFullPath($PreviewSvg)
        if ([IO.Path]::GetExtension($previewPath) -ne '.svg') {
            throw 'PreviewSvg must use the .svg extension.'
        }
        New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($previewPath)) | Out-Null
        $page.Export($previewPath)
    }
    $previewPngPath = $null
    if (-not [string]::IsNullOrWhiteSpace($PreviewPng)) {
        $previewPngPath = [IO.Path]::GetFullPath($PreviewPng)
        if ([IO.Path]::GetExtension($previewPngPath) -ne '.png') {
            throw 'PreviewPng must use the .png extension.'
        }
        New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($previewPngPath)) | Out-Null
        $page.Export($previewPngPath)
    }
    [pscustomobject]@{
        Source = $inputPath
        Output = $outputPath
        Nodes = $shapes.Count
        Connectors = $connectors.Count
        Preview = $previewPath
        PreviewPng = $previewPngPath
        Bytes = (Get-Item -LiteralPath $outputPath).Length
    } | Format-List
} finally {
    foreach ($connector in @($connectors)) {
        if ($null -ne $connector) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($connector) | Out-Null }
    }
    foreach ($shape in @($shapes.Values)) {
        if ($null -ne $shape) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shape) | Out-Null }
    }
    if ($null -ne $page) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($page) | Out-Null }
    if ($null -ne $doc) {
        try { $doc.Close() } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null
    }
    if ($null -ne $app) {
        try { $app.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
