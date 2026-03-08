param(
    [Parameter(Mandatory = $true)]
    [string]$ConversationName,
    [int]$Last = 20,
    [string]$WindowNameRegex = "QQ",
    [ValidateSet("auto", "adapter", "uia", "vision")]
    [string]$Mode = "auto",
    [string]$OcrLanguage = "chi_sim+eng"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. "$PSScriptRoot\qq_uia_common.ps1"
. "$PSScriptRoot\qq_visual_common.ps1"
. "$PSScriptRoot\qq_napcat_common.ps1"

function Invoke-AdapterRead {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetConversation,
        [Parameter(Mandatory = $true)]
        [int]$TakeLast
    )

    if (-not (Test-NapCatAvailable)) {
        return $null
    }

    $target = Resolve-NapCatTargetByName -ConversationName $TargetConversation
    if ($null -eq $target) {
        return $null
    }

    $history = Get-NapCatMessageHistory -Target $target -Count $TakeLast
    $messages = @(
        $history |
            Sort-Object time, message_seq |
            Select-Object -Last $TakeLast |
            ForEach-Object { Convert-NapCatMessageToText -MessageObject $_ }
    )

    return [pscustomobject]@{
        mode_used = "adapter"
        conversation_name = $target.matched_name
        messages = $messages
        confidence = 0.98
        details = @{
            adapter = "napcat"
            target_type = $target.target_type
            target_id = $target.target_id
        }
    }
}

function Invoke-UiaRead {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$TargetConversation,
        [Parameter(Mandatory = $true)]
        [int]$TakeLast
    )

    $opened = Open-ConversationByName -Window $WindowContext.Element -ConversationName $TargetConversation
    if ($null -eq $opened) {
        return $null
    }

    Start-Sleep -Milliseconds 500
    $recent = Get-RecentMessagesFromWindow -Window $WindowContext.Element -Last $TakeLast
    if ($null -eq $recent) {
        return $null
    }

    return [pscustomobject]@{
        mode_used = "uia"
        conversation_name = $TargetConversation
        messages = @($recent.Messages)
        confidence = [math]::Min(1, [math]::Round(($recent.ContainerScore / 20), 2))
        details = @{
            container_score = $recent.ContainerScore
        }
    }
}

function Invoke-VisionRead {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$TargetConversation,
        [Parameter(Mandatory = $true)]
        [int]$TakeLast,
        [Parameter(Mandatory = $true)]
        [string]$Language
    )

    if (-not (Test-TesseractAvailable)) {
        throw "Tesseract executable not found in PATH."
    }

    $opened = Open-ConversationByVision -WindowContext $WindowContext -ConversationName $TargetConversation -Language $Language
    if ($null -eq $opened) {
        return $null
    }

    if (-not (Test-ConversationTitleVisibleByVision -WindowContext $WindowContext -ConversationName $TargetConversation -Language $Language)) {
        return $null
    }

    $visible = Get-VisibleMessagesByVision -WindowContext $WindowContext -Last $TakeLast -Language $Language
    return [pscustomobject]@{
        mode_used = "vision"
        conversation_name = $TargetConversation
        messages = @($visible.Lines | ForEach-Object { $_.Text })
        confidence = [math]::Round(($visible.Confidence / 100), 2)
        details = @{
            screenshot = $visible.Screenshot
        }
    }
}

try {
    $readResult = $null
    $attempts = @()
    $windowContext = $null

    if ($Mode -in @("auto", "adapter")) {
        try {
            $attempts += "adapter"
            $readResult = Invoke-AdapterRead -TargetConversation $ConversationName -TakeLast $Last
        } catch {
            if ($Mode -eq "adapter") {
                throw
            }
        }
    }

    if ($null -eq $readResult -and $Mode -ne "adapter") {
        $windowContext = Get-QQWindowContext -WindowNameRegex $WindowNameRegex
        if ($null -eq $windowContext) {
            $result = New-QQFailureResult -Operation "read" -FailureCode "window_not_found" -Message "Could not find a QQ main window matching regex '$WindowNameRegex', and no adapter result was available." -Data @{
                conversation_name = $ConversationName
                mode_requested = $Mode
                attempted_modes = $attempts
            }
            ConvertTo-QQJson -InputObject $result
            exit 1
        }
    }

    if ($null -eq $readResult -and $Mode -in @("auto", "uia")) {
        try {
            $attempts += "uia"
            $readResult = Invoke-UiaRead -WindowContext $windowContext -TargetConversation $ConversationName -TakeLast $Last
        } catch {
        }
    }

    if ($null -eq $readResult -and $Mode -in @("auto", "vision")) {
        try {
            $attempts += "vision"
            $readResult = Invoke-VisionRead -WindowContext $windowContext -TargetConversation $ConversationName -TakeLast $Last -Language $OcrLanguage
        } catch {
            if ($Mode -eq "vision") {
                throw
            }
        }
    }

    if ($null -eq $readResult) {
        $failureCode = if ($attempts -contains "vision") { "message_region_not_found" } else { "conversation_not_found" }
        $result = New-QQFailureResult -Operation "read" -FailureCode $failureCode -Message "Failed to open '$ConversationName' or extract visible messages." -Data @{
            conversation_name = $ConversationName
            mode_requested = $Mode
            attempted_modes = $attempts
        }
        ConvertTo-QQJson -InputObject $result
        exit 1
    }

    $result = New-QQSuccessResult -Operation "read" -Data @{
        mode_used = $readResult.mode_used
        conversation_name = $readResult.conversation_name
        messages = $readResult.messages
        confidence = $readResult.confidence
        attempted_modes = $attempts
        details = $readResult.details
    }
    ConvertTo-QQJson -InputObject $result
} catch {
    $failureCode = if ($_.Exception.Message -like "*Tesseract*") { "message_region_not_found" } else { "conversation_not_found" }
    $result = New-QQFailureResult -Operation "read" -FailureCode $failureCode -Message $_.Exception.Message -Data @{
        conversation_name = $ConversationName
        mode_requested = $Mode
    }
    ConvertTo-QQJson -InputObject $result
    exit 1
}
