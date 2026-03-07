Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ("QQUia.NativeMethods" -as [type])) {
Add-Type @"
using System;
using System.Runtime.InteropServices;

namespace QQUia {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    }
}
"@
}

$script:QqMouseLeftDown = 0x0002
$script:QqMouseLeftUp = 0x0004

function New-QQFailureResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation,
        [Parameter(Mandatory = $true)]
        [string]$FailureCode,
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [hashtable]$Data
    )

    $result = [ordered]@{
        ok = $false
        operation = $Operation
        failure_code = $FailureCode
        message = $Message
    }

    if ($Data) {
        foreach ($key in $Data.Keys) {
            $result[$key] = $Data[$key]
        }
    }

    return [pscustomobject]$result
}

function New-QQSuccessResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation,
        [hashtable]$Data
    )

    $result = [ordered]@{
        ok = $true
        operation = $Operation
    }

    if ($Data) {
        foreach ($key in $Data.Keys) {
            $result[$key] = $Data[$key]
        }
    }

    return [pscustomobject]$result
}

function ConvertTo-QQJson {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject
    )

    return ($InputObject | ConvertTo-Json -Depth 8)
}

function Get-QQMainWindow {
    param(
        [string]$WindowNameRegex = "QQ"
    )

    $windowCandidates = @()
    $processes = @(Get-Process -Name QQ -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
    foreach ($process in $processes) {
        try {
            $title = $process.MainWindowTitle
            if ($title -notmatch $WindowNameRegex) {
                continue
            }
            $element = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$process.MainWindowHandle)
            if ($null -eq $element) {
                continue
            }
            $rect = $element.Current.BoundingRectangle
            $area = [math]::Max(0, $rect.Width) * [math]::Max(0, $rect.Height)
            $windowCandidates += [pscustomobject]@{
                Element = $element
                Name = $title
                Area = $area
            }
        } catch {
            continue
        }
    }

    if ($windowCandidates.Count -eq 0) {
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $children = $root.FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.Condition]::TrueCondition
        )
        foreach ($child in $children) {
            try {
                $name = $child.Current.Name
                $type = $child.Current.ControlType.ProgrammaticName
                if ($type -ne "ControlType.Window") {
                    continue
                }
                if ($name -match $WindowNameRegex) {
                    $rect = $child.Current.BoundingRectangle
                    $area = [math]::Max(0, $rect.Width) * [math]::Max(0, $rect.Height)
                    $windowCandidates += [pscustomobject]@{
                        Element = $child
                        Name = $name
                        Area = $area
                    }
                }
            } catch {
                continue
            }
        }
    }

    if ($windowCandidates.Count -eq 0) {
        return $null
    }

    return ($windowCandidates | Sort-Object Area -Descending | Select-Object -First 1)
}

function Get-ElementRectangle {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    $rect = $Element.Current.BoundingRectangle
    return [pscustomobject]@{
        Left = [int][math]::Round($rect.Left)
        Top = [int][math]::Round($rect.Top)
        Width = [int][math]::Round($rect.Width)
        Height = [int][math]::Round($rect.Height)
        Right = [int][math]::Round($rect.Right)
        Bottom = [int][math]::Round($rect.Bottom)
    }
}

function Get-RectangleObject {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Left,
        [Parameter(Mandatory = $true)]
        [int]$Top,
        [Parameter(Mandatory = $true)]
        [int]$Width,
        [Parameter(Mandatory = $true)]
        [int]$Height
    )

    return [System.Drawing.Rectangle]::FromLTRB($Left, $Top, $Left + $Width, $Top + $Height)
}

function Set-QQForeground {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    $handle = [IntPtr]$Element.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero) {
        throw "The QQ window has no native handle."
    }

    [QQUia.NativeMethods]::ShowWindowAsync($handle, 5) | Out-Null
    Start-Sleep -Milliseconds 200
    [QQUia.NativeMethods]::SetForegroundWindow($handle) | Out-Null
    Start-Sleep -Milliseconds 350
}

function Test-QQForeground {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    $handle = [IntPtr]$Element.Current.NativeWindowHandle
    return ([QQUia.NativeMethods]::GetForegroundWindow() -eq $handle)
}

function Get-QQWindowContext {
    param(
        [string]$WindowNameRegex = "QQ"
    )

    $windowInfo = @(Get-QQMainWindow -WindowNameRegex $WindowNameRegex) | Select-Object -First 1
    if ($null -eq $windowInfo -or $windowInfo.PSObject.Properties.Name -notcontains "Element") {
        return $null
    }

    Set-QQForeground -Element $windowInfo.Element

    return [pscustomobject]@{
        Element = $windowInfo.Element
        Name = $windowInfo.Name
        Rectangle = Get-ElementRectangle -Element $windowInfo.Element
    }
}

function Get-ElementTextValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    $parts = @()
    try {
        $patternObj = $null
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$patternObj)) {
            $value = ([System.Windows.Automation.ValuePattern]$patternObj).Current.Value
            if ($value) {
                $parts += $value
            }
        }
    } catch {
    }

    try {
        $name = $Element.Current.Name
        if ($name) {
            $parts += $name
        }
    } catch {
    }

    return ($parts | Where-Object { $_ -and $_.Trim() } | Select-Object -Unique) -join " "
}

function Get-Descendants {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Root
    )

    return $Root.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition
    )
}

function Convert-ControlTypeName {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    try {
        $programmatic = $Element.Current.ControlType.ProgrammaticName
        if ($programmatic -like "ControlType.*") {
            return $programmatic.Substring("ControlType.".Length)
        }
        return $programmatic
    } catch {
        return "Unknown"
    }
}

function Export-ElementTree {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Root,
        [int]$MaxDepth = 4
    )

    $walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
    $lines = New-Object System.Collections.Generic.List[string]

    function Visit-Node {
        param(
            [System.Windows.Automation.AutomationElement]$Node,
            [int]$Depth
        )

        if ($Depth -gt $MaxDepth -or $null -eq $Node) {
            return
        }

        try {
            $indent = ("  " * $Depth)
            $name = $Node.Current.Name
            $automationId = $Node.Current.AutomationId
            $className = $Node.Current.ClassName
            $typeName = Convert-ControlTypeName -Element $Node
            $text = Get-ElementTextValue -Element $Node
            $lines.Add("$indent[$typeName] Name='$name' AutomationId='$automationId' Class='$className' Text='$text'")
        } catch {
            $lines.Add("$indent[Unavailable]")
        }

        $child = $walker.GetFirstChild($Node)
        while ($null -ne $child) {
            Visit-Node -Node $child -Depth ($Depth + 1)
            $child = $walker.GetNextSibling($child)
        }
    }

    Visit-Node -Node $Root -Depth 0
    return $lines
}

function Find-CandidateElementsByName {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Root,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $descendants = Get-Descendants -Root $Root
    $candidateElements = @()
    foreach ($item in $descendants) {
        try {
            $currentName = $item.Current.Name
            if ([string]::IsNullOrWhiteSpace($currentName)) {
                continue
            }
            if ($currentName -eq $Name -or $currentName -like "*$Name*") {
                $candidateElements += $item
            }
        } catch {
            continue
        }
    }

    return $candidateElements | Sort-Object {
        $current = $_.Current.Name
        if ($current -eq $Name) { 0 } else { 1 }
    }
}

function Invoke-Element {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    $invokePattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
        ([System.Windows.Automation.InvokePattern]$invokePattern).Invoke()
        return $true
    }

    $selectionPattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$selectionPattern)) {
        ([System.Windows.Automation.SelectionItemPattern]$selectionPattern).Select()
        return $true
    }

    return $false
}

function Open-ConversationByName {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window,
        [Parameter(Mandatory = $true)]
        [string]$ConversationName
    )

    $candidates = Find-CandidateElementsByName -Root $Window -Name $ConversationName
    foreach ($candidate in $candidates) {
        if (Invoke-Element -Element $candidate) {
            Start-Sleep -Milliseconds 500
            return $candidate
        }
    }

    return $null
}

function Find-BestMessageContainer {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window
    )

    $descendants = Get-Descendants -Root $Window
    $scored = @()

    foreach ($item in $descendants) {
        try {
            $typeName = Convert-ControlTypeName -Element $item
            if ($typeName -notin @("Pane", "List", "Document", "Custom")) {
                continue
            }

            $textChildren = $item.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                    [System.Windows.Automation.ControlType]::Text
                ))
            )

            if ($textChildren.Count -lt 3) {
                continue
            }

            $scored += [pscustomobject]@{
                Element = $item
                Score = $textChildren.Count
            }
        } catch {
            continue
        }
    }

    if ($scored.Count -eq 0) {
        return $null
    }

    return ($scored | Sort-Object Score -Descending | Select-Object -First 1)
}

function Get-RecentMessagesFromWindow {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window,
        [int]$Last = 20
    )

    $containerInfo = Find-BestMessageContainer -Window $Window
    if ($null -eq $containerInfo) {
        return $null
    }

    $texts = $containerInfo.Element.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Text
        ))
    )

    $messages = @()
    foreach ($textElement in $texts) {
        try {
            $value = Get-ElementTextValue -Element $textElement
            if ($value -and $value.Trim()) {
                $messages += $value.Trim()
            }
        } catch {
            continue
        }
    }

    $messages = $messages | Select-Object -Unique
    if ($Last -gt 0 -and $messages.Count -gt $Last) {
        $messages = $messages[($messages.Count - $Last)..($messages.Count - 1)]
    }

    return [pscustomobject]@{
        ContainerScore = $containerInfo.Score
        Messages = $messages
    }
}

function Find-BestInputElement {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window
    )

    $descendants = Get-Descendants -Root $Window
    $candidates = @()
    foreach ($item in $descendants) {
        try {
            $typeName = Convert-ControlTypeName -Element $item
            if ($typeName -ne "Edit" -and $typeName -ne "Document") {
                continue
            }

            $rect = $item.Current.BoundingRectangle
            $name = $item.Current.Name
            $score = 0
            if ($rect.Height -gt 40) {
                $score += 2
            }
            if ($rect.Bottom -gt 400) {
                $score += 2
            }
            if ($name -match "input|message|chat") {
                $score += 3
            }

            $candidates += [pscustomobject]@{
                Element = $item
                Score = $score
            }
        } catch {
            continue
        }
    }

    if ($candidates.Count -eq 0) {
        return $null
    }

    return ($candidates | Sort-Object Score -Descending | Select-Object -First 1)
}

function Set-ClipboardMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    Set-Clipboard -Value $Text
}

function Set-ClipboardImageFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )

    if (-not [System.IO.Path]::IsPathRooted($ImagePath)) {
        throw "ImagePath must be an absolute path."
    }
    if (-not (Test-Path -LiteralPath $ImagePath -PathType Leaf)) {
        throw "Image file not found: $ImagePath"
    }

    $ext = [System.IO.Path]::GetExtension($ImagePath).ToLowerInvariant()
    if ($ext -notin @(".png", ".jpg", ".jpeg", ".bmp", ".gif")) {
        throw "Unsupported image extension '$ext'."
    }

    $stream = [System.IO.File]::OpenRead($ImagePath)
    try {
        $source = [System.Drawing.Image]::FromStream($stream)
        try {
            $bitmap = New-Object System.Drawing.Bitmap($source)
            try {
                [System.Windows.Forms.Clipboard]::SetImage($bitmap)
            } finally {
                $bitmap.Dispose()
            }
        } finally {
            $source.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function Invoke-ClickAtPoint {
    param(
        [Parameter(Mandatory = $true)]
        [int]$X,
        [Parameter(Mandatory = $true)]
        [int]$Y
    )

    [QQUia.NativeMethods]::SetCursorPos($X, $Y) | Out-Null
    Start-Sleep -Milliseconds 80
    [QQUia.NativeMethods]::mouse_event($script:QqMouseLeftDown, [uint32]$X, [uint32]$Y, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 40
    [QQUia.NativeMethods]::mouse_event($script:QqMouseLeftUp, [uint32]$X, [uint32]$Y, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 150
}

function Invoke-PasteAndSend {
    param(
        [switch]$SendEnter
    )

    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 150
    if ($SendEnter) {
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    }
}

function Send-QQTextByWindow {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $inputInfo = Find-BestInputElement -Window $Window
    if ($null -ne $inputInfo) {
        try {
            $valuePattern = $null
            if ($inputInfo.Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
                ([System.Windows.Automation.ValuePattern]$valuePattern).SetValue($Message)
                Start-Sleep -Milliseconds 150
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                return [pscustomobject]@{
                    Method = "uia_value_pattern"
                    InputScore = $inputInfo.Score
                }
            }
        } catch {
        }
    }

    Set-QQForeground -Element $Window
    Set-ClipboardMessage -Text $Message
    Invoke-PasteAndSend -SendEnter
    return [pscustomobject]@{
        Method = "clipboard_keyboard_fallback"
        InputScore = if ($null -ne $inputInfo) { $inputInfo.Score } else { -1 }
    }
}

function Test-ConversationTitleVisibleByUIA {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window,
        [Parameter(Mandatory = $true)]
        [string]$ConversationName
    )

    $candidateElements = Find-CandidateElementsByName -Root $Window -Name $ConversationName
    return (@($candidateElements).Count -gt 0)
}
