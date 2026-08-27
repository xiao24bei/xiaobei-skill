[CmdletBinding()]
param(
    [string]$PluginRoot = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($PluginRoot)) {
    $pluginCache = Join-Path ([Environment]::GetFolderPath("UserProfile")) ".codex\plugins\cache\scientific-illustrator-tools\scientific-illustrator"
    if (-not (Test-Path -LiteralPath $pluginCache -PathType Container)) {
        throw "Scientific Illustrator plugin cache was not found: $pluginCache"
    }
    $candidates = @(
        Get-ChildItem -LiteralPath $pluginCache -Directory |
            Where-Object {
                (Test-Path -LiteralPath (Join-Path $_.FullName "scripts\powerpoint-server.mjs") -PathType Leaf) -and
                (Test-Path -LiteralPath (Join-Path $_.FullName "scripts\powerpoint-bridge.ps1") -PathType Leaf)
            } |
            Sort-Object LastWriteTime -Descending
    )
    if ($candidates.Count -lt 1) {
        throw "No installed Scientific Illustrator PowerPoint bridge was found under: $pluginCache"
    }
    $PluginRoot = $candidates[0].FullName
}

$resolvedRoot = (Resolve-Path -LiteralPath $PluginRoot).Path
$serverPath = Join-Path $resolvedRoot "scripts\powerpoint-server.mjs"
$bridgePath = Join-Path $resolvedRoot "scripts\powerpoint-bridge.ps1"
$keeperPath = Join-Path $resolvedRoot "scripts\powerpoint-keeper.ps1"

foreach ($requiredPath in @($serverPath, $bridgePath, $keeperPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required live PowerPoint bridge file is missing: $requiredPath"
    }
}

$server = Get-Content -LiteralPath $serverPath -Raw
$bridge = Get-Content -LiteralPath $bridgePath -Raw
$keeper = Get-Content -LiteralPath $keeperPath -Raw

$checks = [ordered]@{
    server_uses_one_bridge_for_sequence = $server -match 'return\s+runBridge\("draw_sequence",\s*\{\s*operations,\s*step_delay_ms:\s*delay\s*\}\s*\);'
    server_does_not_reconnect_per_object = $server -notmatch 'result:\s*await\s+runBridge\(action,\s*operation\)'
    bridge_has_sequence_action = $bridge -match '"draw_sequence"\s*\{\s*return\s+Invoke-DrawSequence\s+\$Arguments\s*\}'
    bridge_pins_application = $bridge.Contains('$script:PowerPointSessionApplication = $application')
    bridge_pins_presentation = $bridge.Contains('$script:PowerPointSessionPresentation = $presentation')
    bridge_reuses_pinned_application = $bridge.Contains('if ($null -ne $script:PowerPointSessionApplication)')
    bridge_reuses_pinned_presentation = $bridge.Contains('if ($null -ne $script:PowerPointSessionPresentation)')
    bridge_targets_pinned_window = $bridge.Contains('$script:PowerPointSessionPresentation.Windows.Item(1)')
    bridge_reports_session_mode = $bridge.Contains('session_mode = "single_pinned_com_bridge"')
    bridge_starts_task_keeper = $bridge.Contains('Start-PowerPointKeeper $application $presentation')
    keeper_holds_application_reference = $keeper.Contains('$application = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")')
    keeper_holds_presentation_reference = $keeper.Contains('$null = $presentation.Name')
    keeper_tracks_expected_powerpoint_pid = $keeper.Contains('$actualProcessId -ne $ExpectedPowerPointProcessId')
}

$failedChecks = @($checks.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $_.Key })
$result = [ordered]@{
    valid = ($failedChecks.Count -eq 0)
    plugin_root = $resolvedRoot
    checks = $checks
    failed_checks = $failedChecks
}

$result | ConvertTo-Json -Depth 5
if ($failedChecks.Count -gt 0) {
    exit 1
}
