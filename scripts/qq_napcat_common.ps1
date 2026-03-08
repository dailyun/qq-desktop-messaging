Set-StrictMode -Version Latest

function Get-NapCatEndpoint {
    param(
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    return [pscustomobject]@{
        Host = $ApiHost
        Port = $Port
        Token = $Token
        BaseUri = "http://$ApiHost`:$Port"
    }
}

function Invoke-NapCatAction {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action,
        [hashtable]$Payload = @{},
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    $endpoint = Get-NapCatEndpoint -ApiHost $ApiHost -Port $Port -Token $Token
    $uri = "$($endpoint.BaseUri)/$Action"
    $headers = @{ Authorization = "Bearer $($endpoint.Token)" }
    $body = ($Payload | ConvertTo-Json -Depth 12 -Compress)
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $bodyBytes -ContentType "application/json; charset=utf-8"
}

function Test-NapCatAvailable {
    param(
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    try {
        $status = Invoke-NapCatAction -Action "get_status" -ApiHost $ApiHost -Port $Port -Token $Token
        return ($status.status -eq "ok" -and $status.data.online)
    } catch {
        return $false
    }
}

function Get-NapCatFriendList {
    param(
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    $response = Invoke-NapCatAction -Action "get_friend_list" -ApiHost $ApiHost -Port $Port -Token $Token
    return @($response.data)
}

function Get-NapCatGroupList {
    param(
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    $response = Invoke-NapCatAction -Action "get_group_list" -ApiHost $ApiHost -Port $Port -Token $Token
    return @($response.data)
}

function Resolve-NapCatTargetByName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConversationName,
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    $needle = $ConversationName.Trim()
    $friends = Get-NapCatFriendList -ApiHost $ApiHost -Port $Port -Token $Token
    $groups = Get-NapCatGroupList -ApiHost $ApiHost -Port $Port -Token $Token

    $friendCandidates = @(
        $friends | Where-Object { $_.remark -eq $needle -or $_.nickname -eq $needle }
    )
    if ($friendCandidates.Count -eq 0) {
        $friendCandidates = @(
            $friends | Where-Object {
                ($_.remark -and $_.remark.Contains($needle)) -or
                ($_.nickname -and $_.nickname.Contains($needle))
            }
        )
    }
    if ($friendCandidates.Count -gt 0) {
        $friend = $friendCandidates | Select-Object -First 1
        $displayName = if ([string]::IsNullOrWhiteSpace($friend.remark)) { $friend.nickname } else { $friend.remark }
        return [pscustomobject]@{
            target_type = "private"
            target_id = [string]$friend.user_id
            matched_name = $displayName
            raw = $friend
        }
    }

    $groupCandidates = @(
        $groups | Where-Object { $_.group_name -eq $needle -or $_.group_remark -eq $needle }
    )
    if ($groupCandidates.Count -eq 0) {
        $groupCandidates = @(
            $groups | Where-Object {
                ($_.group_name -and $_.group_name.Contains($needle)) -or
                ($_.group_remark -and $_.group_remark.Contains($needle))
            }
        )
    }
    if ($groupCandidates.Count -gt 0) {
        $group = $groupCandidates | Select-Object -First 1
        $displayName = if ([string]::IsNullOrWhiteSpace($group.group_remark)) { $group.group_name } else { $group.group_remark }
        return [pscustomobject]@{
            target_type = "group"
            target_id = [string]$group.group_id
            matched_name = $displayName
            raw = $group
        }
    }

    return $null
}

function Convert-NapCatMessageToText {
    param(
        [Parameter(Mandatory = $true)]
        [object]$MessageObject
    )

    if (-not [string]::IsNullOrWhiteSpace($MessageObject.raw_message)) {
        return [string]$MessageObject.raw_message
    }

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($segment in @($MessageObject.message)) {
        if ($segment.type -eq "text") {
            $parts.Add([string]$segment.data.text)
            continue
        }
        $parts.Add("[$($segment.type)]")
    }
    return ($parts -join "")
}

function Get-NapCatMessageHistory {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Target,
        [int]$Count = 20,
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    $action = if ($Target.target_type -eq "group") { "get_group_msg_history" } else { "get_friend_msg_history" }
    $idKey = if ($Target.target_type -eq "group") { "group_id" } else { "user_id" }
    $payload = @{
        count = $Count
        message_seq = 0
    }
    $payload[$idKey] = [int64]$Target.target_id

    $response = Invoke-NapCatAction -Action $action -Payload $payload -ApiHost $ApiHost -Port $Port -Token $Token
    return @($response.data.messages)
}

function Send-NapCatTextMessage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Target,
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    $action = if ($Target.target_type -eq "group") { "send_group_msg" } else { "send_private_msg" }
    $idKey = if ($Target.target_type -eq "group") { "group_id" } else { "user_id" }
    $payload = @{
        message = $Message
    }
    $payload[$idKey] = [int64]$Target.target_id

    return Invoke-NapCatAction -Action $action -Payload $payload -ApiHost $ApiHost -Port $Port -Token $Token
}

function Send-NapCatImageMessage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Target,
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,
        [string]$ApiHost = $(if ($env:NAPCAT_HOST) { $env:NAPCAT_HOST } else { "127.0.0.1" }),
        [int]$Port = $(if ($env:NAPCAT_PORT) { [int]$env:NAPCAT_PORT } else { 3004 }),
        [string]$Token = $(if ($env:NAPCAT_TOKEN) { $env:NAPCAT_TOKEN } else { "napcat-local-token" })
    )

    $uriPath = [System.Uri]::New($ImagePath).AbsoluteUri
    $action = if ($Target.target_type -eq "group") { "send_group_msg" } else { "send_private_msg" }
    $idKey = if ($Target.target_type -eq "group") { "group_id" } else { "user_id" }
    $payload = @{
        message = @(
            @{
                type = "image"
                data = @{
                    file = $uriPath
                }
            }
        )
    }
    $payload[$idKey] = [int64]$Target.target_id

    return Invoke-NapCatAction -Action $action -Payload $payload -ApiHost $ApiHost -Port $Port -Token $Token
}
