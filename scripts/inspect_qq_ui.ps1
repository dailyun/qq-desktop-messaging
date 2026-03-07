param(
    [int]$MaxDepth = 4,
    [string]$WindowNameRegex = "QQ",
    [string]$OutputPath,
    [switch]$IncludeScreenshotMetadata
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. "$PSScriptRoot\qq_uia_common.ps1"
. "$PSScriptRoot\qq_visual_common.ps1"

try {
    $windowContext = Get-QQWindowContext -WindowNameRegex $WindowNameRegex
    if ($null -eq $windowContext) {
        $result = New-QQFailureResult -Operation "inspect" -FailureCode "window_not_found" -Message "Could not find a QQ main window matching regex '$WindowNameRegex'." -Data @{
            window_name_regex = $WindowNameRegex
        }
        ConvertTo-QQJson -InputObject $result
        exit 1
    }

    $lines = Export-ElementTree -Root $windowContext.Element -MaxDepth $MaxDepth
    if ($OutputPath) {
        $parent = Split-Path -Parent $OutputPath
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent | Out-Null
        }
        $lines | Set-Content -Path $OutputPath -Encoding ASCII
    }

    $screenshot = $null
    if ($IncludeScreenshotMetadata) {
        $screenshot = Save-ScreenshotRegion -Rectangle (Get-RegionRectangleFromWindow -WindowContext $windowContext -Region full)
    }

    $result = New-QQSuccessResult -Operation "inspect" -Data @{
        window_name = $windowContext.Name
        window_rect = $windowContext.Rectangle
        max_depth = $MaxDepth
        ui_tree = $lines
        output_path = $OutputPath
        screenshot_metadata = $screenshot
        tesseract_available = (Test-TesseractAvailable)
    }
    ConvertTo-QQJson -InputObject $result
} catch {
    $result = New-QQFailureResult -Operation "inspect" -FailureCode "window_not_found" -Message $_.Exception.Message -Data @{
        window_name_regex = $WindowNameRegex
    }
    ConvertTo-QQJson -InputObject $result
    exit 1
}
