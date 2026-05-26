param(
    [Parameter(Mandatory = $true)]
    [string]$VbaFile,

    [string]$MacroName = "ReconstructFromImage",

    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

function Write-JsonResult {
    param(
        [string]$Status,
        [hashtable]$Extra = @{}
    )

    $payload = @{ status = $Status }
    foreach ($key in $Extra.Keys) {
        $payload[$key] = $Extra[$key]
    }
    $payload | ConvertTo-Json -Depth 6
}

try {
    $resolvedVba = (Resolve-Path -LiteralPath $VbaFile).Path
} catch {
    Write-JsonResult "vba_file_not_found" @{ path = $VbaFile; error = $_.Exception.Message }
    exit 2
}

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
} catch {
    Write-JsonResult "powerpoint_com_unavailable" @{
        vba_file = $resolvedVba
        error = $_.Exception.Message
        fallback = "Install Microsoft PowerPoint or use the manual .bas import path. WPS automation is not assumed compatible with this PowerPoint COM runner."
    }
    exit 2
}

try {
    $presentation = $powerPoint.Presentations.Add()
    try {
        [void]$presentation.VBProject.VBComponents.Import($resolvedVba)
    } catch {
        Write-JsonResult "vba_import_blocked" @{
            vba_file = $resolvedVba
            error = $_.Exception.Message
            fallback = "In PowerPoint Trust Center, enable trusted macro execution and 'Trust access to the VBA project object model', then rerun. Otherwise import the .bas file manually."
        }
        exit 1
    }

    [void]$powerPoint.Run($MacroName)
    $activePresentation = $powerPoint.ActivePresentation

    $savedPath = $null
    if ($OutputPath -ne "") {
        $target = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        $parent = Split-Path -Parent $target
        if ($parent -and !(Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Force -Path $parent | Out-Null
        }
        $activePresentation.SaveAs($target, 25)
        $savedPath = $target
    }

    Write-JsonResult "ran" @{
        vba_file = $resolvedVba
        macro = $MacroName
        saved_path = $savedPath
    }
    exit 0
} catch {
    Write-JsonResult "automation_failed" @{
        vba_file = $resolvedVba
        macro = $MacroName
        error = $_.Exception.Message
        fallback = "Open PowerPoint, import or paste the .bas file into the VBA editor, and run the macro manually."
    }
    exit 1
}
