[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Plan,
    [ValidateRange(10, 1800)][int]$TimeoutSeconds = 180,
    [string]$ExportPreview = "",
    [switch]$CreateBlank,
    [switch]$CheckOnly,
    [switch]$Wait
)
$ErrorActionPreference = "Stop"

$planFile = (Resolve-Path -LiteralPath $Plan).Path
$planDir = Split-Path -Parent $planFile
$config = Get-Content -Raw -Encoding UTF8 -LiteralPath $planFile | ConvertFrom-Json

$isObserver = ($null -ne $config.cues) -and ($null -eq $config.actions) -and ($null -eq $config.operations) -and ($null -eq $config.elements)
$driverName = if ($isObserver) { "swarm_observer.py" } else { "swarm_live_draw.py" }
$driverPath = Join-Path $PSScriptRoot $driverName

$pythonCmd = Get-Command python.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $pythonCmd) {
    $pythonCmd = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (-not $pythonCmd) {
    throw "Python was not found on PATH. Install Python 3.10+ with pywin32 (pip install pywin32)."
}
$python = $pythonCmd.Source

if ($CheckOnly) {
    if ($isObserver) { throw "-CheckOnly is supported for swarm_live_draw.py action plans." }
    & $python -X utf8 $driverPath $planFile --check
    if ($LASTEXITCODE -ne 0) { throw "Swarm drawing plan failed preflight check." }
    return
}

# Automatically ensure PowerPoint is visibly open, maximized, and showing the target deck.
if (-not $isObserver -and $config.deck) {
    $deckPath = if ([System.IO.Path]::IsPathRooted([string]$config.deck)) {
        [System.IO.Path]::GetFullPath([string]$config.deck)
    } else {
        [System.IO.Path]::GetFullPath((Join-Path $planDir ([string]$config.deck)))
    }
    $ppt = $null
    try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch {}
    $needOpen = $CreateBlank -or ($null -eq $ppt) -or (-not (Test-Path -LiteralPath $deckPath))
    if (-not $needOpen -and $null -ne $ppt) {
        $alreadyOpen = $false
        for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
            try {
                if ([System.IO.Path]::GetFullPath($ppt.Presentations.Item($i).FullName).Equals($deckPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $alreadyOpen = $true
                    $ppt.Presentations.Item($i).Windows.Item(1).Activate()
                    break
                }
            } catch {}
        }
        if (-not $alreadyOpen) { $needOpen = $true }
    }
    if ($needOpen) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $deckPath) | Out-Null
        if ($null -eq $ppt) {
            $ppt = New-Object -ComObject PowerPoint.Application
        }
        $ppt.Visible = -1
        try { $ppt.WindowState = 3 } catch {} # ppWindowMaximized
        if ($CreateBlank -or (-not (Test-Path -LiteralPath $deckPath))) {
            $pres = $ppt.Presentations.Add(-1)
            if ($config.canvas -and $config.canvas.Count -eq 2) {
                $pres.PageSetup.SlideWidth = [single]$config.canvas[0]
                $pres.PageSetup.SlideHeight = [single]$config.canvas[1]
            }
            $null = $pres.Slides.Add(1, 12) # ppLayoutBlank
            $pres.SaveAs($deckPath)
        } else {
            $pres = $ppt.Presentations.Open($deckPath, 0, 0, -1)
        }
        $pres.Windows.Item(1).Activate()
        Start-Sleep -Milliseconds 250
    }
}

$runId = [Guid]::NewGuid().ToString("N")
$stdout = Join-Path $planDir "swarm-$runId.log"
$stderr = "$stdout.err"

$argList = "-u -X utf8 `"$driverPath`" `"$planFile`" --timeout $TimeoutSeconds"
if (-not $isObserver -and $ExportPreview) {
    $argList += " --export-preview `"$ExportPreview`""
}

$proc = Start-Process -FilePath $python -ArgumentList $argList -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
$timer = [Diagnostics.Stopwatch]::StartNew()
$ready = $false
while ($timer.Elapsed.TotalSeconds -lt 25) {
    if ($proc.HasExited) {
        $errText = if (Test-Path -LiteralPath $stderr) { Get-Content -Raw -Encoding UTF8 -LiteralPath $stderr } else { "" }
        throw "Swarm cursor engine exited ($($proc.ExitCode)). $errText"
    }
    if ((Test-Path -LiteralPath $stdout) -and ((Get-Content -LiteralPath $stdout -Encoding UTF8) -contains "READY")) {
        $ready = $true
        break
    }
    Start-Sleep -Milliseconds 100
}
if (-not $ready) {
    if (-not $proc.HasExited) { $proc.Kill() }
    throw "PowerPoint swarm overlay did not attach within 25s. Check $stderr and ensure the presentation is active."
}

if ($Wait) {
    $proc.WaitForExit()
    if ($proc.ExitCode -ne 0) {
        $errText = if (Test-Path -LiteralPath $stderr) { Get-Content -Raw -Encoding UTF8 -LiteralPath $stderr } else { "" }
        throw "Swarm drawing failed ($($proc.ExitCode)). $errText"
    }
}

[pscustomobject]@{
    pid = $proc.Id
    driver = $driverName
    log = $stdout
    errors = $stderr
    completed = $proc.HasExited
} | ConvertTo-Json -Compress
