param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,

    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function Resolve-AssetPath {
    param([string]$BaseDirectory, [string]$Value)

    if ([System.IO.Path]::IsPathRooted($Value)) {
        return [System.IO.Path]::GetFullPath($Value)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $BaseDirectory $Value))
}

function Parse-HexColor {
    param([string]$Value)

    $hex = $Value.TrimStart('#')
    if ($hex.Length -ne 6) { throw "background_color must be a 6-digit RGB hex value." }
    return [System.Drawing.Color]::FromArgb(
        255,
        [Convert]::ToInt32($hex.Substring(0, 2), 16),
        [Convert]::ToInt32($hex.Substring(2, 2), 16),
        [Convert]::ToInt32($hex.Substring(4, 2), 16)
    )
}

function Get-CornerAverageColor {
    param([System.Drawing.Bitmap]$Bitmap)

    $points = @(
        @(0, 0),
        @(($Bitmap.Width - 1), 0),
        @(0, ($Bitmap.Height - 1)),
        @(($Bitmap.Width - 1), ($Bitmap.Height - 1))
    )
    $red = 0
    $green = 0
    $blue = 0
    foreach ($point in $points) {
        $color = $Bitmap.GetPixel($point[0], $point[1])
        $red += $color.R
        $green += $color.G
        $blue += $color.B
    }
    return [System.Drawing.Color]::FromArgb(255, [int]($red / 4), [int]($green / 4), [int]($blue / 4))
}

function Remove-EdgeConnectedBackground {
    param(
        [System.Drawing.Bitmap]$Bitmap,
        [System.Drawing.Color]$BackgroundColor,
        [int]$Tolerance
    )

    $width = $Bitmap.Width
    $height = $Bitmap.Height
    $count = $width * $height
    $visited = [bool[]]::new($count)
    $background = [bool[]]::new($count)
    $queue = [System.Collections.Generic.Queue[int]]::new()
    $toleranceSquared = $Tolerance * $Tolerance

    for ($x = 0; $x -lt $width; $x++) {
        $queue.Enqueue($x)
        $queue.Enqueue((($height - 1) * $width) + $x)
    }
    for ($y = 1; $y -lt ($height - 1); $y++) {
        $queue.Enqueue($y * $width)
        $queue.Enqueue(($y * $width) + ($width - 1))
    }

    while ($queue.Count -gt 0) {
        $index = $queue.Dequeue()
        if ($visited[$index]) { continue }
        $visited[$index] = $true

        $x = $index % $width
        $y = [int][math]::Floor($index / $width)
        $color = $Bitmap.GetPixel($x, $y)
        $dr = [int]$color.R - [int]$BackgroundColor.R
        $dg = [int]$color.G - [int]$BackgroundColor.G
        $db = [int]$color.B - [int]$BackgroundColor.B
        if (($dr * $dr + $dg * $dg + $db * $db) -gt $toleranceSquared) { continue }

        $background[$index] = $true
        if ($x -gt 0) { $queue.Enqueue($index - 1) }
        if ($x -lt ($width - 1)) { $queue.Enqueue($index + 1) }
        if ($y -gt 0) { $queue.Enqueue($index - $width) }
        if ($y -lt ($height - 1)) { $queue.Enqueue($index + $width) }
    }

    for ($index = 0; $index -lt $count; $index++) {
        if (-not $background[$index]) { continue }
        $x = $index % $width
        $y = [int][math]::Floor($index / $width)
        $color = $Bitmap.GetPixel($x, $y)
        $Bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(0, $color.R, $color.G, $color.B))
    }
}

function Keep-LargestOpaqueComponent {
    param(
        [System.Drawing.Bitmap]$Bitmap,
        [int]$AlphaThreshold = 8
    )

    $width = $Bitmap.Width
    $height = $Bitmap.Height
    $count = $width * $height
    $visited = [bool[]]::new($count)
    $best = [int[]]@()

    for ($start = 0; $start -lt $count; $start++) {
        if ($visited[$start]) { continue }
        $visited[$start] = $true
        $startX = $start % $width
        $startY = [int][math]::Floor($start / $width)
        if ($Bitmap.GetPixel($startX, $startY).A -le $AlphaThreshold) { continue }

        $queue = [System.Collections.Generic.Queue[int]]::new()
        $component = [System.Collections.Generic.List[int]]::new()
        $queue.Enqueue($start)
        while ($queue.Count -gt 0) {
            $index = $queue.Dequeue()
            $component.Add($index)
            $x = $index % $width
            $y = [int][math]::Floor($index / $width)

            for ($dy = -1; $dy -le 1; $dy++) {
                for ($dx = -1; $dx -le 1; $dx++) {
                    if ($dx -eq 0 -and $dy -eq 0) { continue }
                    $nextX = $x + $dx
                    $nextY = $y + $dy
                    if ($nextX -lt 0 -or $nextX -ge $width -or $nextY -lt 0 -or $nextY -ge $height) { continue }
                    $next = ($nextY * $width) + $nextX
                    if ($visited[$next]) { continue }
                    $visited[$next] = $true
                    if ($Bitmap.GetPixel($nextX, $nextY).A -gt $AlphaThreshold) { $queue.Enqueue($next) }
                }
            }
        }

        if ($component.Count -gt $best.Count) { $best = $component.ToArray() }
    }

    if ($best.Count -eq 0) { throw "No opaque component remains after background removal." }
    $keep = [bool[]]::new($count)
    foreach ($index in $best) { $keep[$index] = $true }
    for ($index = 0; $index -lt $count; $index++) {
        if ($keep[$index]) { continue }
        $x = $index % $width
        $y = [int][math]::Floor($index / $width)
        $color = $Bitmap.GetPixel($x, $y)
        if ($color.A -gt 0) { $Bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(0, $color.R, $color.G, $color.B)) }
    }
}

function Get-OpaqueMetrics {
    param(
        [System.Drawing.Bitmap]$Bitmap,
        [int]$AlphaThreshold = 8
    )

    $opaqueCount = 0
    $minX = $Bitmap.Width
    $minY = $Bitmap.Height
    $maxX = -1
    $maxY = -1
    $touchesEdge = $false
    for ($y = 0; $y -lt $Bitmap.Height; $y++) {
        for ($x = 0; $x -lt $Bitmap.Width; $x++) {
            if ($Bitmap.GetPixel($x, $y).A -le $AlphaThreshold) { continue }
            $opaqueCount++
            if ($x -lt $minX) { $minX = $x }
            if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }
            if ($y -gt $maxY) { $maxY = $y }
            if ($x -eq 0 -or $y -eq 0 -or $x -eq ($Bitmap.Width - 1) -or $y -eq ($Bitmap.Height - 1)) { $touchesEdge = $true }
        }
    }

    $bounds = if ($opaqueCount -gt 0) { @($minX, $minY, ($maxX - $minX + 1), ($maxY - $minY + 1)) } else { @() }
    return [pscustomobject]@{
        opaque_count = $opaqueCount
        opaque_ratio = [math]::Round($opaqueCount / [double]($Bitmap.Width * $Bitmap.Height), 5)
        opaque_bounds_px = $bounds
        touches_edge = $touchesEdge
    }
}

function Draw-Checkerboard {
    param(
        [System.Drawing.Graphics]$Graphics,
        [System.Drawing.Rectangle]$Rectangle
    )

    $size = 12
    $light = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
    $dark = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(232, 235, 238))
    try {
        for ($y = $Rectangle.Top; $y -lt $Rectangle.Bottom; $y += $size) {
            for ($x = $Rectangle.Left; $x -lt $Rectangle.Right; $x += $size) {
                $brush = if ((([int](($x - $Rectangle.Left) / $size) + [int](($y - $Rectangle.Top) / $size)) % 2) -eq 0) { $light } else { $dark }
                $Graphics.FillRectangle($brush, $x, $y, [math]::Min($size, $Rectangle.Right - $x), [math]::Min($size, $Rectangle.Bottom - $y))
            }
        }
    }
    finally {
        $light.Dispose()
        $dark.Dispose()
    }
}

$manifestFullPath = [System.IO.Path]::GetFullPath($ManifestPath)
$manifestDirectory = Split-Path -Parent $manifestFullPath
$manifest = Get-Content -Raw -LiteralPath $manifestFullPath -Encoding UTF8 | ConvertFrom-Json
if (-not $manifest.source) { throw "Manifest requires source." }
if (-not $manifest.assets -or $manifest.assets.Count -eq 0) { throw "Manifest requires at least one asset." }

$sourcePath = Resolve-AssetPath $manifestDirectory ([string]$manifest.source)
$outputDirectoryValue = if ($manifest.output_dir) { [string]$manifest.output_dir } else { "assets" }
$outputDirectory = Resolve-AssetPath $manifestDirectory $outputDirectoryValue
$contactSheetValue = if ($manifest.contact_sheet) { [string]$manifest.contact_sheet } else { "asset-contact-sheet.png" }
$contactSheetPath = Resolve-AssetPath $manifestDirectory $contactSheetValue
$reportValue = if ($manifest.report) { [string]$manifest.report } else { "asset-extraction-report.json" }
$reportPath = Resolve-AssetPath $manifestDirectory $reportValue
$maxAreaRatio = if ($null -ne $manifest.max_asset_area_ratio) { [double]$manifest.max_asset_area_ratio } else { 0.35 }
$strictValidation = if ($null -ne $manifest.strict_validation) { [bool]$manifest.strict_validation } else { $false }
$saveRawCrops = if ($null -ne $manifest.save_raw_crops) { [bool]$manifest.save_raw_crops } else { $true }
$rawOutputDirectoryValue = if ($manifest.raw_output_dir) { [string]$manifest.raw_output_dir } else { Join-Path $outputDirectoryValue "_raw" }
$rawOutputDirectory = Resolve-AssetPath $manifestDirectory $rawOutputDirectoryValue
$defaultComponentLossLimit = if ($null -ne $manifest.max_component_loss_ratio) { [double]$manifest.max_component_loss_ratio } else { 0.08 }

New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
if ($saveRawCrops) { New-Item -ItemType Directory -Force -Path $rawOutputDirectory | Out-Null }
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $contactSheetPath) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $reportPath) | Out-Null

$source = [System.Drawing.Bitmap]::new($sourcePath)
$sourceWidth = $source.Width
$sourceHeight = $source.Height
$results = [System.Collections.Generic.List[object]]::new()
try {
    foreach ($asset in $manifest.assets) {
        if (-not $asset.id) { throw "Every asset requires id." }
        if (-not $asset.source_rect_px -or $asset.source_rect_px.Count -ne 4) { throw "Asset '$($asset.id)' requires source_rect_px [x,y,width,height]." }

        if ($strictValidation) {
            if (-not $asset.visual_role) { throw "Asset '$($asset.id)' requires visual_role when strict_validation=true." }
            if (-not $asset.raster_reason) { throw "Asset '$($asset.id)' requires raster_reason when strict_validation=true." }
            if (-not $asset.expected_content -or $asset.expected_content.Count -eq 0) { throw "Asset '$($asset.id)' requires expected_content when strict_validation=true." }
            if (-not ($asset.PSObject.Properties.Name -contains "forbidden_content")) { throw "Asset '$($asset.id)' requires forbidden_content when strict_validation=true." }
            if (-not ($asset.PSObject.Properties.Name -contains "native_surroundings")) { throw "Asset '$($asset.id)' requires native_surroundings when strict_validation=true." }
        }

        $x = [int]$asset.source_rect_px[0]
        $y = [int]$asset.source_rect_px[1]
        $width = [int]$asset.source_rect_px[2]
        $height = [int]$asset.source_rect_px[3]
        if ($x -lt 0 -or $y -lt 0 -or $width -le 0 -or $height -le 0 -or ($x + $width) -gt $source.Width -or ($y + $height) -gt $source.Height) {
            throw "Asset '$($asset.id)' is outside the source bounds."
        }

        $areaRatio = ($width * $height) / [double]($sourceWidth * $sourceHeight)
        $allowLarge = ($asset.allow_large -eq $true)
        if ($areaRatio -gt $maxAreaRatio -and -not $allowLarge) {
            throw "Asset '$($asset.id)' covers $([math]::Round($areaRatio * 100, 1))% of the source; split it or set allow_large only for a documented exception."
        }

        $outputValue = if ($asset.output) { [string]$asset.output } else { "$($asset.id).png" }
        $outputPath = Resolve-AssetPath $outputDirectory $outputValue
        $rawOutputPath = Resolve-AssetPath $rawOutputDirectory ([System.IO.Path]::GetFileNameWithoutExtension($outputValue) + "_raw.png")
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $outputPath) | Out-Null
        if ((Test-Path -LiteralPath $outputPath) -and -not $Overwrite) { throw "Output exists: $outputPath. Use -Overwrite to replace it." }
        if ($saveRawCrops -and (Test-Path -LiteralPath $rawOutputPath) -and -not $Overwrite) { throw "Raw crop exists: $rawOutputPath. Use -Overwrite to replace it." }

        $warnings = [System.Collections.Generic.List[string]]::new()
        $componentLossRatio = 0.0
        $backgroundColorDistance = $null
        $bitmap = [System.Drawing.Bitmap]::new($width, $height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        try {
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
                $graphics.DrawImage(
                    $source,
                    [System.Drawing.Rectangle]::new(0, 0, $width, $height),
                    [System.Drawing.Rectangle]::new($x, $y, $width, $height),
                    [System.Drawing.GraphicsUnit]::Pixel
                )
            }
            finally {
                $graphics.Dispose()
            }

            if ($saveRawCrops) { $bitmap.Save($rawOutputPath, [System.Drawing.Imaging.ImageFormat]::Png) }
            $rawCornerColor = Get-CornerAverageColor $bitmap

            $alphaMethod = if ($asset.alpha_method) { [string]$asset.alpha_method } else { "none" }
            if ($alphaMethod -eq "edge-connected") {
                $backgroundColor = if ($asset.background_color) { Parse-HexColor ([string]$asset.background_color) } else { Get-CornerAverageColor $bitmap }
                $tolerance = if ($null -ne $asset.tolerance) { [int]$asset.tolerance } else { 28 }
                Remove-EdgeConnectedBackground $bitmap $backgroundColor $tolerance
            }
            elseif ($alphaMethod -ne "none" -and $alphaMethod -ne "matched-background") {
                throw "Asset '$($asset.id)' has unsupported alpha_method '$alphaMethod'."
            }

            if ($asset.destination_background_color -and ($alphaMethod -eq "none" -or $alphaMethod -eq "matched-background")) {
                $destinationColor = Parse-HexColor ([string]$asset.destination_background_color)
                $dr = [int]$rawCornerColor.R - [int]$destinationColor.R
                $dg = [int]$rawCornerColor.G - [int]$destinationColor.G
                $db = [int]$rawCornerColor.B - [int]$destinationColor.B
                $backgroundColorDistance = [math]::Round([math]::Sqrt(($dr * $dr) + ($dg * $dg) + ($db * $db)), 2)
                $backgroundTolerance = if ($null -ne $asset.background_match_tolerance) { [double]$asset.background_match_tolerance } else { 18.0 }
                if ($backgroundColorDistance -gt $backgroundTolerance) { $warnings.Add("crop_background_differs_from_destination") }
                if ($strictValidation -and $alphaMethod -eq "matched-background" -and $backgroundColorDistance -gt $backgroundTolerance -and $asset.allow_background_mismatch -ne $true) {
                    throw "Asset '$($asset.id)' corner background differs from destination_background_color by $backgroundColorDistance. Re-crop, change the background method, or document allow_background_mismatch=true."
                }
            }

            $preFilterMetrics = Get-OpaqueMetrics $bitmap
            $componentFilter = if ($asset.component_filter) { [string]$asset.component_filter } else { "none" }
            if ($componentFilter -eq "largest") {
                if ($alphaMethod -ne "edge-connected") { throw "Asset '$($asset.id)' can use component_filter=largest only with alpha_method=edge-connected." }
                if ($strictValidation) {
                    if ($asset.expected_connected_parts -ne 1) { throw "Asset '$($asset.id)' can use component_filter=largest in strict mode only with expected_connected_parts=1." }
                    if (-not $asset.component_filter_reason) { throw "Asset '$($asset.id)' requires component_filter_reason when component_filter=largest in strict mode." }
                }
                Keep-LargestOpaqueComponent $bitmap
            }
            elseif ($componentFilter -ne "none") {
                throw "Asset '$($asset.id)' has unsupported component_filter '$componentFilter'."
            }

            $postFilterMetrics = Get-OpaqueMetrics $bitmap
            if ($preFilterMetrics.opaque_count -gt 0) {
                $componentLossRatio = [math]::Round(($preFilterMetrics.opaque_count - $postFilterMetrics.opaque_count) / [double]$preFilterMetrics.opaque_count, 5)
            }
            if ($componentFilter -eq "largest" -and $componentLossRatio -gt 0.02) {
                $warnings.Add("component_filter_removed_$([math]::Round($componentLossRatio * 100, 1))_percent")
            }
            $componentLossLimit = if ($null -ne $asset.max_component_loss_ratio) { [double]$asset.max_component_loss_ratio } else { $defaultComponentLossLimit }
            if ($strictValidation -and $componentFilter -eq "largest" -and $componentLossRatio -gt $componentLossLimit -and $asset.allow_component_loss -ne $true) {
                throw "Asset '$($asset.id)' lost $([math]::Round($componentLossRatio * 100, 1))% of opaque content under component_filter=largest. Inspect the raw crop or set allow_component_loss=true with a documented reason."
            }
            if ($postFilterMetrics.touches_edge) { $warnings.Add("opaque_content_touches_crop_edge") }
            if ($alphaMethod -eq "edge-connected" -and $postFilterMetrics.opaque_ratio -gt 0.98) { $warnings.Add("edge_background_removal_may_have_failed") }

            $bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $bitmap.Dispose()
        }

        $results.Add([pscustomobject]@{
            id = [string]$asset.id
            visual_role = [string]$asset.visual_role
            source_rect_px = @($x, $y, $width, $height)
            output_path = $outputPath
            raw_output_path = if ($saveRawCrops) { $rawOutputPath } else { $null }
            alpha_method = $alphaMethod
            component_filter = $componentFilter
            component_loss_ratio = $componentLossRatio
            background_color_distance = $backgroundColorDistance
            opaque_bounds_px = $postFilterMetrics.opaque_bounds_px
            touches_edge = $postFilterMetrics.touches_edge
            warnings = @($warnings)
            area_ratio = [math]::Round($areaRatio, 5)
        })
    }
}
finally {
    $source.Dispose()
}

$columns = [math]::Min(2, $results.Count)
$cellWidth = 680
$cellHeight = 310
$rows = [int][math]::Ceiling($results.Count / [double]$columns)
$sheet = [System.Drawing.Bitmap]::new($columns * $cellWidth, $rows * $cellHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
try {
    $graphics = [System.Drawing.Graphics]::FromImage($sheet)
    $font = [System.Drawing.Font]::new("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $smallFont = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $textBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(32, 36, 40))
    $borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(180, 185, 190), 1)
    try {
        $graphics.Clear([System.Drawing.Color]::White)
        for ($index = 0; $index -lt $results.Count; $index++) {
            $column = $index % $columns
            $row = [int][math]::Floor($index / $columns)
            $left = $column * $cellWidth
            $top = $row * $cellHeight
            $graphics.DrawString($results[$index].id, $font, $textBrush, $left + 12, $top + 10)

            $gap = 12
            $previewWidth = [int](($cellWidth - 36) / 2)
            $previewHeight = $cellHeight - 124
            $rawPreview = [System.Drawing.Rectangle]::new($left + 12, $top + 58, $previewWidth, $previewHeight)
            $processedPreview = [System.Drawing.Rectangle]::new($rawPreview.Right + $gap, $top + 58, $previewWidth, $previewHeight)
            $rawPreviewPath = if ($results[$index].raw_output_path) { $results[$index].raw_output_path } else { $results[$index].output_path }
            $previewItems = @(
                [pscustomobject]@{ label = "raw source crop"; path = $rawPreviewPath; rectangle = $rawPreview },
                [pscustomobject]@{ label = "processed asset"; path = $results[$index].output_path; rectangle = $processedPreview }
            )

            foreach ($previewItem in $previewItems) {
                $graphics.DrawString($previewItem.label, $smallFont, $textBrush, $previewItem.rectangle.Left, $top + 39)
                Draw-Checkerboard $graphics $previewItem.rectangle
                $graphics.DrawRectangle($borderPen, $previewItem.rectangle)
                $assetImage = [System.Drawing.Image]::FromFile($previewItem.path)
                try {
                    $scale = [math]::Min($previewItem.rectangle.Width / [double]$assetImage.Width, $previewItem.rectangle.Height / [double]$assetImage.Height)
                    $drawWidth = [int][math]::Round($assetImage.Width * $scale)
                    $drawHeight = [int][math]::Round($assetImage.Height * $scale)
                    $drawLeft = $previewItem.rectangle.Left + [int](($previewItem.rectangle.Width - $drawWidth) / 2)
                    $drawTop = $previewItem.rectangle.Top + [int](($previewItem.rectangle.Height - $drawHeight) / 2)
                    $graphics.DrawImage($assetImage, $drawLeft, $drawTop, $drawWidth, $drawHeight)
                }
                finally {
                    $assetImage.Dispose()
                }
            }

            $warningText = if ($results[$index].warnings.Count -gt 0) { $results[$index].warnings -join "; " } else { "none" }
            $rectText = "rect: [$($results[$index].source_rect_px -join ', ')]  alpha: $($results[$index].alpha_method)  component: $($results[$index].component_filter)`nloss: $([math]::Round($results[$index].component_loss_ratio * 100, 1))%  warnings: $warningText"
            $textRectangle = [System.Drawing.RectangleF]::new($left + 12, $top + $cellHeight - 57, $cellWidth - 24, 52)
            $graphics.DrawString($rectText, $smallFont, $textBrush, $textRectangle)
        }
        if ((Test-Path -LiteralPath $contactSheetPath) -and -not $Overwrite) { throw "Contact sheet exists: $contactSheetPath. Use -Overwrite to replace it." }
        $sheet.Save($contactSheetPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $font.Dispose()
        $smallFont.Dispose()
        $textBrush.Dispose()
        $borderPen.Dispose()
    }
}
finally {
    $sheet.Dispose()
}

$report = [pscustomobject]@{
    source_path = $sourcePath
    source_size_px = @($sourceWidth, $sourceHeight)
    asset_count = $results.Count
    max_asset_area_ratio = $maxAreaRatio
    strict_validation = $strictValidation
    save_raw_crops = $saveRawCrops
    raw_output_dir = if ($saveRawCrops) { $rawOutputDirectory } else { $null }
    contact_sheet = $contactSheetPath
    assets = $results
}
if ((Test-Path -LiteralPath $reportPath) -and -not $Overwrite) { throw "Report exists: $reportPath. Use -Overwrite to replace it." }
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8

$report | ConvertTo-Json -Depth 8
