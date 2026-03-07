Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. "$PSScriptRoot\qq_uia_common.ps1"

function Get-TesseractCommand {
    $cmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($null -eq $cmd) {
        $fallbacks = @(
            "C:\Program Files\Tesseract-OCR\tesseract.exe",
            "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
        )
        foreach ($candidate in $fallbacks) {
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
        return $null
    }

    return $cmd.Source
}

function Get-TesseractDataPrefix {
    $candidates = @(
        (Join-Path $PSScriptRoot "tessdata"),
        (Join-Path (Split-Path $PSScriptRoot -Parent) "tessdata"),
        "C:\Program Files\Tesseract-OCR\tessdata"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Test-TesseractAvailable {
    return ($null -ne (Get-TesseractCommand))
}

function Get-NormalizedOcrText {
    param(
        [AllowEmptyString()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    return (($Text -replace "\s+", "") -replace "[`"'|,.;:!?()\[\]{}<>_~\-]+", "").Trim()
}

function Get-RegionRectangleFromWindow {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [ValidateSet("search", "sidebar", "title", "messages", "input", "full")]
        [string]$Region
    )

    $rect = $WindowContext.Rectangle
    switch ($Region) {
        "full" {
            return Get-RectangleObject -Left $rect.Left -Top $rect.Top -Width $rect.Width -Height $rect.Height
        }
        "search" {
            return Get-RectangleObject -Left ([int]($rect.Left + $rect.Width * 0.055)) -Top ([int]($rect.Top + $rect.Height * 0.035)) -Width ([int]($rect.Width * 0.17)) -Height ([int]($rect.Height * 0.05))
        }
        "sidebar" {
            return Get-RectangleObject -Left $rect.Left -Top $rect.Top -Width ([int]($rect.Width * 0.28)) -Height $rect.Height
        }
        "title" {
            return Get-RectangleObject -Left ([int]($rect.Left + $rect.Width * 0.28)) -Top $rect.Top -Width ([int]($rect.Width * 0.72)) -Height ([int]($rect.Height * 0.09))
        }
        "messages" {
            return Get-RectangleObject -Left ([int]($rect.Left + $rect.Width * 0.28)) -Top ([int]($rect.Top + $rect.Height * 0.09)) -Width ([int]($rect.Width * 0.72)) -Height ([int]($rect.Height * 0.72))
        }
        "input" {
            return Get-RectangleObject -Left ([int]($rect.Left + $rect.Width * 0.28)) -Top ([int]($rect.Top + $rect.Height * 0.81)) -Width ([int]($rect.Width * 0.72)) -Height ([int]($rect.Height * 0.17))
        }
    }
}

function Save-ScreenshotRegion {
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Rectangle]$Rectangle,
        [string]$OutputPath,
        [int]$Scale = 2
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path ([System.IO.Path]::GetTempPath()) ("qq-" + [guid]::NewGuid().ToString("N") + ".png")
    }

    $bitmap = New-Object System.Drawing.Bitmap($Rectangle.Width, $Rectangle.Height)
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.CopyFromScreen($Rectangle.Left, $Rectangle.Top, 0, 0, $bitmap.Size)
        } finally {
            $graphics.Dispose()
        }

        if ($Scale -gt 1) {
            $scaled = New-Object System.Drawing.Bitmap(($Rectangle.Width * $Scale), ($Rectangle.Height * $Scale))
            try {
                $scaledGraphics = [System.Drawing.Graphics]::FromImage($scaled)
                try {
                    $scaledGraphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                    $scaledGraphics.DrawImage($bitmap, 0, 0, $scaled.Width, $scaled.Height)
                } finally {
                    $scaledGraphics.Dispose()
                }
                $scaled.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
            } finally {
                $scaled.Dispose()
            }
        } else {
            $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
    } finally {
        $bitmap.Dispose()
    }

    return [pscustomobject]@{
        path = $OutputPath
        left = $Rectangle.Left
        top = $Rectangle.Top
        width = $Rectangle.Width
        height = $Rectangle.Height
        scale = $Scale
    }
}

function Invoke-TesseractTsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,
        [string]$Language = "chi_sim+eng"
    )

    $tesseract = Get-TesseractCommand
    if ($null -eq $tesseract) {
        throw "Tesseract executable not found in PATH."
    }

    $dataPrefix = Get-TesseractDataPrefix
    $stderrPath = Join-Path ([System.IO.Path]::GetTempPath()) ("qq-tesseract-" + [guid]::NewGuid().ToString("N") + ".log")
    try {
        if ($dataPrefix) {
            $output = & $tesseract $ImagePath stdout --tessdata-dir $dataPrefix --psm 6 -l $Language -c tessedit_create_tsv=1 2> $stderrPath
        } else {
            $output = & $tesseract $ImagePath stdout --psm 6 -l $Language -c tessedit_create_tsv=1 2> $stderrPath
        }
        if ($LASTEXITCODE -ne 0) {
            $stderr = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { "" }
            throw "Tesseract failed: $stderr"
        }
    } finally {
        if (Test-Path -LiteralPath $stderrPath) {
            Remove-Item -LiteralPath $stderrPath -Force
        }
    }

    return $output
}

function ConvertFrom-TesseractTsv {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Lines,
        [Parameter(Mandatory = $true)]
        [int]$OffsetLeft,
        [Parameter(Mandatory = $true)]
        [int]$OffsetTop
    )

    $rows = @()
    foreach ($line in ($Lines | Select-Object -Skip 1)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $parts = $line -split "`t"
        if ($parts.Count -lt 12) {
            continue
        }
        $text = $parts[11].Trim()
        if (-not $text) {
            continue
        }
        $conf = 0
        [int]::TryParse($parts[10], [ref]$conf) | Out-Null
        $left = 0
        $top = 0
        $width = 0
        $height = 0
        [int]::TryParse($parts[6], [ref]$left) | Out-Null
        [int]::TryParse($parts[7], [ref]$top) | Out-Null
        [int]::TryParse($parts[8], [ref]$width) | Out-Null
        [int]::TryParse($parts[9], [ref]$height) | Out-Null
        $rows += [pscustomobject]@{
            Text = $text
            Confidence = $conf
            Left = $OffsetLeft + $left
            Top = $OffsetTop + $top
            Width = $width
            Height = $height
            CenterX = $OffsetLeft + $left + [int]($width / 2)
            CenterY = $OffsetTop + $top + [int]($height / 2)
        }
    }

    return $rows
}

function Get-OcrDataForRegion {
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Rectangle]$Rectangle,
        [string]$Language = "chi_sim+eng"
    )

    $shot = Save-ScreenshotRegion -Rectangle $Rectangle
    try {
        $tsv = Invoke-TesseractTsv -ImagePath $shot.path -Language $Language
        $scaledWords = ConvertFrom-TesseractTsv -Lines $tsv -OffsetLeft 0 -OffsetTop 0
        $words = foreach ($word in $scaledWords) {
            [pscustomobject]@{
                Text = $word.Text
                Confidence = $word.Confidence
                Left = $shot.left + [int]($word.Left / $shot.scale)
                Top = $shot.top + [int]($word.Top / $shot.scale)
                Width = [int]($word.Width / $shot.scale)
                Height = [int]($word.Height / $shot.scale)
                CenterX = $shot.left + [int]($word.CenterX / $shot.scale)
                CenterY = $shot.top + [int]($word.CenterY / $shot.scale)
            }
        }
        return [pscustomobject]@{
            Screenshot = $shot
            Words = $words
        }
    } finally {
        if (Test-Path -LiteralPath $shot.path) {
            Remove-Item -LiteralPath $shot.path -Force
        }
    }
}

function Find-OcrAnchor {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Words,
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $needle = Get-NormalizedOcrText -Text $Text
    $wordCandidates = @(
        $Words | Where-Object {
            $candidate = Get-NormalizedOcrText -Text $_.Text
            $candidate -and ($candidate -eq $needle -or $candidate.Contains($needle) -or $needle.Contains($candidate))
        }
    )
    if ($wordCandidates.Count -eq 0) {
        return $null
    }

    return ($wordCandidates | Sort-Object @{Expression = { if ((Get-NormalizedOcrText -Text $_.Text) -eq $needle) { 0 } else { 1 } } }, @{Expression = { -1 * $_.Confidence } } | Select-Object -First 1)
}

function Get-OcrTextLines {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Words
    )

    if (@($Words).Count -eq 0) {
        return @()
    }

    $groups = @()
    foreach ($word in ($Words | Sort-Object Top, Left)) {
        $bucket = $null
        foreach ($existing in $groups) {
            if ([math]::Abs($existing.Top - $word.Top) -le 12) {
                $bucket = $existing
                break
            }
        }

        if ($null -eq $bucket) {
            $bucket = [pscustomobject]@{
                Top = $word.Top
                Items = New-Object System.Collections.Generic.List[object]
            }
            $groups += $bucket
        }

        $bucket.Items.Add($word)
    }

    $lines = foreach ($group in ($groups | Sort-Object Top)) {
        $text = (($group.Items | Sort-Object Left | ForEach-Object { $_.Text.Trim() }) -join " ").Trim()
        if ($text) {
            [pscustomobject]@{
                Top = $group.Top
                Text = $text
                Confidence = [math]::Round((($group.Items | Measure-Object Confidence -Average).Average), 2)
                Left = (($group.Items | Measure-Object Left -Minimum).Minimum)
                Right = (($group.Items | ForEach-Object { $_.Left + $_.Width } | Measure-Object -Maximum).Maximum)
                CenterX = [int](( (($group.Items | Measure-Object Left -Minimum).Minimum) + (($group.Items | ForEach-Object { $_.Left + $_.Width } | Measure-Object -Maximum).Maximum) ) / 2)
                CenterY = [int]($group.Top + 12)
            }
        }
    }

    return @($lines | Select-Object -Unique Text, Top, Confidence, Left, Right, CenterX, CenterY)
}

function Focus-SearchBoxByVision {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext
    )

    $searchRect = Get-RegionRectangleFromWindow -WindowContext $WindowContext -Region search
    $x = $searchRect.Left + [int]($searchRect.Width / 2)
    $y = $searchRect.Top + [int]($searchRect.Height / 2)
    Invoke-ClickAtPoint -X $x -Y $y

    return [pscustomobject]@{
        X = $x
        Y = $y
        Region = @{
            Left = $searchRect.Left
            Top = $searchRect.Top
            Width = $searchRect.Width
            Height = $searchRect.Height
        }
    }
}

function Set-SearchBoxTextByVision {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$ConversationName
    )

    $focus = Focus-SearchBoxByVision -WindowContext $WindowContext
    Start-Sleep -Milliseconds 120
    [System.Windows.Forms.SendKeys]::SendWait("^a")
    Start-Sleep -Milliseconds 80
    [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
    Start-Sleep -Milliseconds 100
    Set-Clipboard -Value $ConversationName
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 500
    return $focus
}

function Find-ConversationLineInSidebar {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$ConversationName,
        [string]$Language = "chi_sim+eng"
    )

    $sidebar = Get-RegionRectangleFromWindow -WindowContext $WindowContext -Region sidebar
    $searchRect = Get-RegionRectangleFromWindow -WindowContext $WindowContext -Region search
    $ocr = Get-OcrDataForRegion -Rectangle $sidebar -Language $Language
    $lines = @(
        Get-OcrTextLines -Words $ocr.Words | Where-Object {
            $_.Top -gt ($searchRect.Top + $searchRect.Height + 18)
        }
    )
    $needle = Get-NormalizedOcrText -Text $ConversationName
    $lineCandidates = @(
        $lines | Where-Object {
            $candidate = Get-NormalizedOcrText -Text $_.Text
            $candidate -and (
                $candidate -eq $needle -or
                $candidate.Contains($needle) -or
                ($candidate.Length -ge 2 -and $needle.Contains($candidate))
            )
        }
    )
    if ($lineCandidates.Count -eq 0) {
        return $null
    }

    $best = $lineCandidates | Sort-Object @{Expression = { if ((Get-NormalizedOcrText -Text $_.Text) -eq $needle) { 0 } else { 1 } } }, Top | Select-Object -First 1
    return [pscustomobject]@{
        Line = $best
        Screenshot = $ocr.Screenshot
    }
}

function Open-ConversationByVision {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$ConversationName,
        [string]$Language = "chi_sim+eng"
    )

    $lineResult = Find-ConversationLineInSidebar -WindowContext $WindowContext -ConversationName $ConversationName -Language $Language
    if ($null -eq $lineResult) {
        Set-SearchBoxTextByVision -WindowContext $WindowContext -ConversationName $ConversationName | Out-Null
        Start-Sleep -Milliseconds 300
        $lineResult = Find-ConversationLineInSidebar -WindowContext $WindowContext -ConversationName $ConversationName -Language $Language
        if ($null -eq $lineResult) {
            return $null
        }
    }

    Invoke-ClickAtPoint -X $lineResult.Line.CenterX -Y $lineResult.Line.CenterY
    Start-Sleep -Milliseconds 500

    return [pscustomobject]@{
        Method = "vision_anchor_click"
        Anchor = $lineResult.Line
        Screenshot = $lineResult.Screenshot
    }
}

function Test-ConversationTitleVisibleByVision {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$ConversationName,
        [string]$Language = "chi_sim+eng"
    )

    $titleRect = Get-RegionRectangleFromWindow -WindowContext $WindowContext -Region title
    $ocr = Get-OcrDataForRegion -Rectangle $titleRect -Language $Language
    return ($null -ne (Find-OcrAnchor -Words $ocr.Words -Text $ConversationName))
}

function Get-VisibleMessagesByVision {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [int]$Last = 20,
        [string]$Language = "chi_sim+eng"
    )

    $messageRect = Get-RegionRectangleFromWindow -WindowContext $WindowContext -Region messages
    $ocr = Get-OcrDataForRegion -Rectangle $messageRect -Language $Language
    $lines = @(Get-OcrTextLines -Words $ocr.Words)
    if ($Last -gt 0 -and $lines.Count -gt $Last) {
        $lines = $lines[($lines.Count - $Last)..($lines.Count - 1)]
    }

    return [pscustomobject]@{
        Screenshot = $ocr.Screenshot
        Lines = $lines
        Confidence = if ($lines.Count -gt 0) { [math]::Round((($lines | Measure-Object Confidence -Average).Average), 2) } else { 0 }
    }
}

function Focus-InputByVision {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext
    )

    $inputRect = Get-RegionRectangleFromWindow -WindowContext $WindowContext -Region input
    $x = $inputRect.Left + [int]($inputRect.Width / 2)
    $y = $inputRect.Top + [int]($inputRect.Height / 2)
    Invoke-ClickAtPoint -X $x -Y $y

    return [pscustomobject]@{
        Method = "vision_input_click"
        X = $x
        Y = $y
        Region = @{
            Left = $inputRect.Left
            Top = $inputRect.Top
            Width = $inputRect.Width
            Height = $inputRect.Height
        }
    }
}
