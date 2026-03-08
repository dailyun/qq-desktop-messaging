param(
    [Parameter(Mandatory = $true)]
    [string]$ConversationName,
    [string]$Message,
    [string]$ImagePath,
    [string]$WindowNameRegex = "QQ",
    [ValidateSet("auto", "adapter", "uia", "vision")]
    [string]$Mode = "auto",
    [ValidateSet("text", "image")]
    [string]$ContentType = "text",
    [string]$OcrLanguage = "chi_sim+eng"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. "$PSScriptRoot\qq_uia_common.ps1"
. "$PSScriptRoot\qq_visual_common.ps1"
. "$PSScriptRoot\qq_napcat_common.ps1"

function Open-TargetConversationByAdapter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetConversation
    )

    if (-not (Test-NapCatAvailable)) {
        return $null
    }

    return Resolve-NapCatTargetByName -ConversationName $TargetConversation
}

function Open-TargetConversationAuto {
    param(
        [AllowNull()]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$TargetConversation,
        [Parameter(Mandatory = $true)]
        [string]$SelectedMode,
        [Parameter(Mandatory = $true)]
        [string]$Language
    )

    $attempted = @()

    if ($SelectedMode -in @("auto", "adapter")) {
        $attempted += "adapter"
        $resolved = Open-TargetConversationByAdapter -TargetConversation $TargetConversation
        if ($null -ne $resolved) {
            return [pscustomobject]@{
                mode_used = "adapter"
                attempted = $attempted
                method = "napcat_name_lookup"
                target = $resolved
            }
        }
    }

    if ($null -ne $WindowContext -and $SelectedMode -in @("auto", "uia")) {
        $attempted += "uia"
        $opened = Open-ConversationByName -Window $WindowContext.Element -ConversationName $TargetConversation
        if ($null -ne $opened) {
            return [pscustomobject]@{
                mode_used = "uia"
                attempted = $attempted
                method = "uia_name_match"
                target = $null
            }
        }
    }

    if ($null -ne $WindowContext -and $SelectedMode -in @("auto", "vision")) {
        if (-not (Test-TesseractAvailable)) {
            if ($SelectedMode -eq "vision") {
                throw "Tesseract executable not found in PATH."
            }
        } else {
            $attempted += "vision"
            $opened = Open-ConversationByVision -WindowContext $WindowContext -ConversationName $TargetConversation -Language $Language
            if ($null -ne $opened) {
                return [pscustomobject]@{
                    mode_used = "vision"
                    attempted = $attempted
                    method = $opened.Method
                    target = $null
                }
            }
        }
    }

    return [pscustomobject]@{
        mode_used = $null
        attempted = $attempted
        method = $null
        target = $null
    }
}

function Confirm-TargetConversation {
    param(
        [Parameter(Mandatory = $true)]
        [object]$WindowContext,
        [Parameter(Mandatory = $true)]
        [string]$TargetConversation,
        [Parameter(Mandatory = $true)]
        [string]$PreferredMode,
        [Parameter(Mandatory = $true)]
        [string]$Language
    )

    if ($PreferredMode -eq "uia" -and (Test-ConversationTitleVisibleByUIA -Window $WindowContext.Element -ConversationName $TargetConversation)) {
        return $true
    }

    if ($PreferredMode -eq "adapter") {
        return $true
    }

    if (Test-TesseractAvailable) {
        return (Test-ConversationTitleVisibleByVision -WindowContext $WindowContext -ConversationName $TargetConversation -Language $Language)
    }

    return (Test-ConversationTitleVisibleByUIA -Window $WindowContext.Element -ConversationName $TargetConversation)
}

try {
    if ($ContentType -eq "text" -and [string]::IsNullOrWhiteSpace($Message)) {
        throw "Message is required when ContentType is 'text'."
    }

    if ($ContentType -eq "image" -and [string]::IsNullOrWhiteSpace($ImagePath)) {
        throw "ImagePath is required when ContentType is 'image'."
    }

    $windowContext = $null
    $openResult = Open-TargetConversationAuto -WindowContext $windowContext -TargetConversation $ConversationName -SelectedMode $Mode -Language $OcrLanguage
    if (-not $openResult.mode_used -and $Mode -ne "adapter") {
        $windowContext = Get-QQWindowContext -WindowNameRegex $WindowNameRegex
        if ($null -ne $windowContext) {
            $openResult = Open-TargetConversationAuto -WindowContext $windowContext -TargetConversation $ConversationName -SelectedMode $Mode -Language $OcrLanguage
        }
    }

    if (-not $openResult.mode_used) {
        $result = New-QQFailureResult -Operation "send" -FailureCode "conversation_not_found" -Message "Failed to open the target conversation '$ConversationName'." -Data @{
            conversation_name = $ConversationName
            mode_requested = $Mode
            content_type = $ContentType
            attempted_modes = $openResult.attempted
        }
        ConvertTo-QQJson -InputObject $result
        exit 1
    }

    if ($openResult.mode_used -ne "adapter") {
        if ($null -eq $windowContext) {
            $result = New-QQFailureResult -Operation "send" -FailureCode "window_not_found" -Message "Could not find a QQ main window matching regex '$WindowNameRegex', and no adapter-only send was available." -Data @{
                conversation_name = $ConversationName
                mode_requested = $Mode
                content_type = $ContentType
                attempted_modes = $openResult.attempted
            }
            ConvertTo-QQJson -InputObject $result
            exit 1
        }
    }

    if ($openResult.mode_used -ne "adapter") {
        Start-Sleep -Milliseconds 500
        if (-not (Test-QQForeground -Element $windowContext.Element)) {
            Set-QQForeground -Element $windowContext.Element
        }
    }

    if ($openResult.mode_used -ne "adapter" -and -not (Confirm-TargetConversation -WindowContext $windowContext -TargetConversation $ConversationName -PreferredMode $openResult.mode_used -Language $OcrLanguage)) {
        $result = New-QQFailureResult -Operation "send" -FailureCode "send_not_confirmed" -Message "QQ is foreground but the target conversation title could not be confirmed." -Data @{
            conversation_name = $ConversationName
            mode_used = $openResult.mode_used
            content_type = $ContentType
        }
        ConvertTo-QQJson -InputObject $result
        exit 1
    }

    if ($ContentType -eq "text") {
        $sendResult = $null
        if ($openResult.mode_used -eq "adapter") {
            $adapterResponse = Send-NapCatTextMessage -Target $openResult.target -Message $Message
            $sendResult = [pscustomobject]@{
                Method = "napcat_onebot_http"
                Response = $adapterResponse
            }
        } elseif ($openResult.mode_used -eq "uia" -or $Mode -eq "uia") {
            $sendResult = Send-QQTextByWindow -Window $windowContext.Element -Message $Message
        } else {
            $focus = Focus-InputByVision -WindowContext $windowContext
            Set-ClipboardMessage -Text $Message
            Invoke-PasteAndSend -SendEnter
            $sendResult = [pscustomobject]@{
                Method = "vision_clipboard_keyboard"
                InputScore = 0
                Focus = $focus
            }
        }

        $result = New-QQSuccessResult -Operation "send" -Data @{
            mode_used = $openResult.mode_used
            content_type = "text"
            target = if ($openResult.mode_used -eq "adapter") { $openResult.target.matched_name } else { $ConversationName }
            send_method = $sendResult.Method
            confidence = if ($openResult.mode_used -eq "adapter") { 0.99 } elseif ($openResult.mode_used -eq "uia") { 0.9 } else { 0.7 }
            attempted_modes = $openResult.attempted
            details = if ($openResult.mode_used -eq "adapter") {
                @{
                    adapter = "napcat"
                    target_type = $openResult.target.target_type
                    target_id = $openResult.target.target_id
                    message_id = $sendResult.Response.data.message_id
                }
            } else {
                $null
            }
        }
        ConvertTo-QQJson -InputObject $result
        exit 0
    }

    if (-not [System.IO.Path]::IsPathRooted($ImagePath)) {
        throw "ImagePath must be an absolute path."
    }

    if ($openResult.mode_used -eq "adapter") {
        try {
            $adapterImageResponse = Send-NapCatImageMessage -Target $openResult.target -ImagePath $ImagePath
            $result = New-QQSuccessResult -Operation "send" -Data @{
                mode_used = "adapter"
                content_type = "image"
                target = $openResult.target.matched_name
                send_method = "napcat_onebot_http"
                confidence = 0.95
                attempted_modes = $openResult.attempted
                image_path = $ImagePath
                details = @{
                    adapter = "napcat"
                    target_type = $openResult.target.target_type
                    target_id = $openResult.target.target_id
                    message_id = $adapterImageResponse.data.message_id
                }
            }
            ConvertTo-QQJson -InputObject $result
            exit 0
        } catch {
            if ($Mode -eq "adapter") {
                throw
            }
            $windowContext = Get-QQWindowContext -WindowNameRegex $WindowNameRegex
            if ($null -eq $windowContext) {
                throw
            }
            $openResult = Open-TargetConversationAuto -WindowContext $windowContext -TargetConversation $ConversationName -SelectedMode "vision" -Language $OcrLanguage
        }
    }

    $focus = if ($openResult.mode_used -eq "uia") {
        $bestInput = Find-BestInputElement -Window $windowContext.Element
        if ($null -ne $bestInput) {
            Invoke-Element -Element $bestInput.Element | Out-Null
            [pscustomobject]@{ Method = "uia_input_focus" }
        } else {
            Focus-InputByVision -WindowContext $windowContext
        }
    } else {
        Focus-InputByVision -WindowContext $windowContext
    }

    if ($null -eq $focus) {
        $result = New-QQFailureResult -Operation "send" -FailureCode "input_region_not_found" -Message "Could not focus the QQ input region." -Data @{
            conversation_name = $ConversationName
            mode_used = $openResult.mode_used
            content_type = "image"
        }
        ConvertTo-QQJson -InputObject $result
        exit 1
    }

    Set-ClipboardImageFile -ImagePath $ImagePath
    Start-Sleep -Milliseconds 150
    Invoke-PasteAndSend -SendEnter

    $result = New-QQSuccessResult -Operation "send" -Data @{
        mode_used = $openResult.mode_used
        content_type = "image"
        target = $ConversationName
        send_method = "clipboard_image_paste"
        confidence = if ($openResult.mode_used -eq "uia") { 0.85 } else { 0.65 }
        attempted_modes = $openResult.attempted
        image_path = $ImagePath
    }
    ConvertTo-QQJson -InputObject $result
} catch {
    $message = $_.Exception.Message
    $failureCode = "send_not_confirmed"
    if ($message -like "*ImagePath must be an absolute path*") {
        $failureCode = "input_region_not_found"
    } elseif ($message -like "*Tesseract*") {
        $failureCode = "input_region_not_found"
    } elseif ($message -like "*image*not found*" -or $message -like "*Unsupported image extension*") {
        $failureCode = "input_region_not_found"
    }

    $result = New-QQFailureResult -Operation "send" -FailureCode $failureCode -Message $message -Data @{
        conversation_name = $ConversationName
        mode_requested = $Mode
        content_type = $ContentType
    }
    ConvertTo-QQJson -InputObject $result
    exit 1
}
