param(
    [Parameter(Mandatory = $true)]
    [string]$SceneMapPath,

    [ValidateSet("planning", "final")]
    [string]$Phase = "planning",

    [switch]$AllowDocumentedExceptions
)

$ErrorActionPreference = "Stop"

function Has-Property {
    param([object]$Value, [string]$Name)
    return $null -ne $Value -and ($Value.PSObject.Properties.Name -contains $Name)
}

function Add-Issue {
    param(
        [System.Collections.Generic.List[string]]$List,
        [string]$Message
    )
    $List.Add($Message)
}

$fullPath = [System.IO.Path]::GetFullPath($SceneMapPath)
if (-not (Test-Path -LiteralPath $fullPath)) { throw "Scene map not found: $fullPath" }
$map = Get-Content -LiteralPath $fullPath -Raw -Encoding UTF8 | ConvertFrom-Json
$errors = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()

if (-not $map.source -or -not $map.source.width_px -or -not $map.source.height_px) { Add-Issue $errors "source width_px and height_px are required." }
if (-not $map.canvas -or -not $map.canvas.width_pt -or -not $map.canvas.height_pt) { Add-Issue $errors "canvas width_pt and height_pt are required." }
if (-not $map.fidelity_profile) { Add-Issue $errors "fidelity_profile is required." }
if (-not $map.mode) { Add-Issue $errors "mode is required." }
if (-not $map.reference_inventory) { Add-Issue $errors "reference_inventory is required." }
else {
    foreach ($field in @("complete", "regions", "counts", "required_ids", "unresolved_ambiguities", "authorized_omissions")) {
        if (-not (Has-Property $map.reference_inventory $field)) { Add-Issue $errors "reference_inventory is missing $field." }
    }
    if ($map.reference_inventory.complete -ne $true) { Add-Issue $errors "reference_inventory.complete must be true before drawing." }
}

$objects = @($map.objects)
$links = @($map.links)
$allEntries = @($objects) + @($links)
$ids = @($allEntries | ForEach-Object { [string]$_.id })

foreach ($group in ($ids | Group-Object)) {
    if (-not $group.Name) { Add-Issue $errors "Every object and link requires a non-empty id." }
    elseif ($group.Count -gt 1) { Add-Issue $errors "Duplicate id: $($group.Name)." }
}

$allowedStatus = @("planned", "drawn", "verified")
$allowedCategories = @("panel", "text_item", "visual_core", "evidence_tile", "plot", "native_symbol", "inset", "caption")
$regions = @($map.reference_inventory.regions | ForEach-Object { [string]$_ })
$sourceWidth = if ($map.source.width_px) { [double]$map.source.width_px } else { 0 }
$sourceHeight = if ($map.source.height_px) { [double]$map.source.height_px } else { 0 }

foreach ($object in $objects) {
    $id = [string]$object.id
    foreach ($field in @("parent_region", "kind", "required", "status", "source_category", "source_rect_px", "mapping", "layer", "editability")) {
        if (-not (Has-Property $object $field)) { Add-Issue $errors "Object '$id' is missing $field." }
    }
    if ($regions.Count -gt 0 -and $regions -notcontains [string]$object.parent_region) { Add-Issue $errors "Object '$id' references unknown parent_region '$($object.parent_region)'." }
    if ($allowedStatus -notcontains [string]$object.status) { Add-Issue $errors "Object '$id' has invalid status '$($object.status)'." }
    if ($Phase -eq "final" -and $object.required -eq $true -and $object.status -ne "verified") { Add-Issue $errors "Required object '$id' is not verified in final phase." }
    if ($allowedCategories -notcontains [string]$object.source_category) { Add-Issue $errors "Object '$id' has invalid source_category '$($object.source_category)'." }
    if ($object.mapping -ne "scaled" -and -not (Has-Property $object "mapping_adjustment_reason")) { Add-Issue $errors "Object '$id' uses mapping '$($object.mapping)' without mapping_adjustment_reason." }

    $rect = @($object.source_rect_px)
    if ($rect.Count -ne 4) {
        Add-Issue $errors "Object '$id' requires source_rect_px [x,y,width,height]."
    }
    elseif ($sourceWidth -gt 0 -and $sourceHeight -gt 0) {
        $x = [double]$rect[0]
        $y = [double]$rect[1]
        $width = [double]$rect[2]
        $height = [double]$rect[3]
        if ($x -lt 0 -or $y -lt 0 -or $width -le 0 -or $height -le 0 -or ($x + $width) -gt $sourceWidth -or ($y + $height) -gt $sourceHeight) {
            Add-Issue $errors "Object '$id' has source_rect_px outside source bounds."
        }
    }

    if ($object.source_category -eq "text_item" -or $object.source_category -eq "caption") {
        if (-not (Has-Property $object "text")) { Add-Issue $errors "Text object '$id' is missing exact text." }
        if ($object.line_breaks -ne "exact") { Add-Issue $errors "Text object '$id' requires line_breaks='exact'." }
        if (-not (Has-Property $object "style")) { Add-Issue $errors "Text object '$id' is missing measured font/alignment/color style." }
    }

    if ($object.kind -eq "picture") {
        foreach ($field in @("asset", "visual_role", "raster_reason", "alpha_method", "expected_content", "forbidden_content", "native_surroundings", "fidelity_priority")) {
            if (-not (Has-Property $object $field)) { Add-Issue $errors "Picture '$id' is missing $field." }
        }
        if (-not $object.expected_content -or $object.expected_content.Count -eq 0) { Add-Issue $errors "Picture '$id' requires non-empty expected_content." }
    }

    if ($object.source_category -eq "plot" -and -not (Has-Property $object "plot_plan")) {
        Add-Issue $errors "Plot '$id' is missing plot_plan for axes, labels, legend, and visible traces."
    }
}

$objectIds = @($objects | ForEach-Object { [string]$_.id })
foreach ($link in $links) {
    $id = [string]$link.id
    foreach ($field in @("parent_region", "required", "status", "from", "to", "direction", "topology_verified")) {
        if (-not (Has-Property $link $field)) { Add-Issue $errors "Link '$id' is missing $field." }
    }
    if ($regions.Count -gt 0 -and $regions -notcontains [string]$link.parent_region) { Add-Issue $errors "Link '$id' references unknown parent_region '$($link.parent_region)'." }
    if ($allowedStatus -notcontains [string]$link.status) { Add-Issue $errors "Link '$id' has invalid status '$($link.status)'." }
    if ($Phase -eq "final" -and $link.required -eq $true -and $link.status -ne "verified") { Add-Issue $errors "Required link '$id' is not verified in final phase." }
    if ($Phase -eq "final" -and $link.required -eq $true -and $link.topology_verified -ne $true) { Add-Issue $errors "Required link '$id' has not passed topology verification." }
    if ($objectIds -notcontains [string]$link.from) { Add-Issue $errors "Link '$id' references missing source object '$($link.from)'." }
    if ($objectIds -notcontains [string]$link.to) { Add-Issue $errors "Link '$id' references missing target object '$($link.to)'." }
    if (-not (Has-Property $link "source_route_px") -and -not (Has-Property $link "route_kind")) { Add-Issue $errors "Link '$id' requires source_route_px or route_kind." }
}

if ($map.reference_inventory) {
    $requiredIds = @($map.reference_inventory.required_ids | ForEach-Object { [string]$_ })
    $mappedRequiredIds = @($allEntries | Where-Object { $_.required -eq $true } | ForEach-Object { [string]$_.id })
    foreach ($id in $requiredIds) {
        if (($ids | Where-Object { $_ -eq $id }).Count -ne 1) { Add-Issue $errors "reference_inventory.required_ids entry '$id' does not resolve exactly once." }
    }
    foreach ($id in $mappedRequiredIds) {
        if ($requiredIds -notcontains $id) { Add-Issue $errors "Required entry '$id' is missing from reference_inventory.required_ids." }
    }
    foreach ($id in $requiredIds) {
        if ($mappedRequiredIds -notcontains $id) { Add-Issue $errors "reference_inventory.required_ids contains '$id', but the entry is not marked required." }
    }

    $categoryMap = @{
        panels = "panel"
        text_items = "text_item"
        visual_cores = "visual_core"
        evidence_tiles = "evidence_tile"
        plots = "plot"
        native_symbols = "native_symbol"
        insets = "inset"
        captions = "caption"
    }
    if ($map.reference_inventory.counts) {
        foreach ($property in $map.reference_inventory.counts.PSObject.Properties) {
            $expected = [int]$property.Value
            if ($property.Name -eq "links") {
                $actual = @($links | Where-Object { $_.required -eq $true }).Count
            }
            elseif ($categoryMap.ContainsKey($property.Name)) {
                $category = $categoryMap[$property.Name]
                $actual = @($objects | Where-Object { $_.required -eq $true -and $_.source_category -eq $category }).Count
            }
            else {
                Add-Issue $warnings "Unrecognized reference_inventory.counts key '$($property.Name)' was not reconciled."
                continue
            }
            if ($actual -ne $expected) { Add-Issue $errors "Inventory count '$($property.Name)' expected $expected but mapped $actual required entries." }
        }
    }

    $ambiguities = @($map.reference_inventory.unresolved_ambiguities)
    $omissions = @($map.reference_inventory.authorized_omissions)
    if (-not $AllowDocumentedExceptions -and $ambiguities.Count -gt 0) { Add-Issue $errors "Unresolved source ambiguities remain; review them or pass -AllowDocumentedExceptions after disclosure." }
    if (-not $AllowDocumentedExceptions -and $omissions.Count -gt 0) { Add-Issue $errors "Authorized omissions are recorded; pass -AllowDocumentedExceptions only after user approval." }
}

$report = [pscustomobject]@{
    scene_map = $fullPath
    phase = $Phase
    object_count = $objects.Count
    link_count = $links.Count
    required_count = @($allEntries | Where-Object { $_.required -eq $true }).Count
    error_count = $errors.Count
    warning_count = $warnings.Count
    errors = @($errors)
    warnings = @($warnings)
    valid = ($errors.Count -eq 0)
}

$report | ConvertTo-Json -Depth 8
if ($errors.Count -gt 0) { exit 1 }
