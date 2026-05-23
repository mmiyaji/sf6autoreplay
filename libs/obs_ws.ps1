param(
    [string]$HostName = "127.0.0.1",
    [int]$Port = 4455,
    [string]$Password = "",
    [ValidateSet("status", "start", "stop", "split")]
    [string]$Action = "status",
    [int]$TimeoutMs = 5000
)

$ErrorActionPreference = "Stop"

function ConvertTo-Base64Sha256([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToBase64String($sha.ComputeHash($bytes))
    } finally {
        $sha.Dispose()
    }
}

function Receive-WebSocketText($Socket, [System.Threading.CancellationToken]$Token) {
    $buffer = New-Object byte[] 8192
    $segments = New-Object System.Collections.Generic.List[byte]

    do {
        $seg = [System.ArraySegment[byte]]::new($buffer)
        $result = $Socket.ReceiveAsync($seg, $Token).GetAwaiter().GetResult()
        if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
            throw "OBS WebSocket closed the connection"
        }
        for ($i = 0; $i -lt $result.Count; $i++) {
            $segments.Add($buffer[$i])
        }
    } while (-not $result.EndOfMessage)

    return [System.Text.Encoding]::UTF8.GetString($segments.ToArray())
}

function Send-WebSocketJson($Socket, $Payload, [System.Threading.CancellationToken]$Token) {
    $json = $Payload | ConvertTo-Json -Depth 20 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $seg = [System.ArraySegment[byte]]::new($bytes)
    $null = $Socket.SendAsync($seg, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $Token).GetAwaiter().GetResult()
}

function Invoke-ObsRequest($Socket, [string]$RequestType, $RequestData, [System.Threading.CancellationToken]$Token) {
    $requestId = [guid]::NewGuid().ToString()
    $payload = @{
        op = 6
        d = @{
            requestType = $RequestType
            requestId = $requestId
        }
    }
    if ($null -ne $RequestData) {
        $payload.d.requestData = $RequestData
    }

    Send-WebSocketJson $Socket $payload $Token

    while ($true) {
        $message = Receive-WebSocketText $Socket $Token
        $obj = $message | ConvertFrom-Json
        if ($obj.op -ne 7) {
            continue
        }
        if ($obj.d.requestId -ne $requestId) {
            continue
        }
        return $obj.d
    }
}

$cts = [System.Threading.CancellationTokenSource]::new($TimeoutMs)
$socket = [System.Net.WebSockets.ClientWebSocket]::new()

try {
    $uri = [Uri]::new("ws://$HostName`:$Port")
    $null = $socket.ConnectAsync($uri, $cts.Token).GetAwaiter().GetResult()

    $hello = Receive-WebSocketText $socket $cts.Token | ConvertFrom-Json
    if ($hello.op -ne 0) {
        throw "Unexpected OBS WebSocket hello opcode: $($hello.op)"
    }

    $identify = @{
        op = 1
        d = @{
            rpcVersion = 1
        }
    }

    if ($hello.d.authentication) {
        if ([string]::IsNullOrEmpty($Password)) {
            throw "OBS WebSocket requires a password"
        }
        $secret = ConvertTo-Base64Sha256 ($Password + $hello.d.authentication.salt)
        $auth = ConvertTo-Base64Sha256 ($secret + $hello.d.authentication.challenge)
        $identify.d.authentication = $auth
    }

    Send-WebSocketJson $socket $identify $cts.Token

    while ($true) {
        $identified = Receive-WebSocketText $socket $cts.Token | ConvertFrom-Json
        if ($identified.op -eq 2) {
            break
        }
    }

    $requestType = switch ($Action) {
        "status" { "GetRecordStatus" }
        "start"  { "StartRecord" }
        "stop"   { "StopRecord" }
        "split"  { "SplitRecordFile" }
    }

    $response = Invoke-ObsRequest $socket $requestType $null $cts.Token
    $ok = [bool]$response.requestStatus.result
    $result = [ordered]@{
        ok = $ok
        action = $Action
        requestType = $requestType
        code = $response.requestStatus.code
        comment = $response.requestStatus.comment
        data = $response.responseData
    }

    $result | ConvertTo-Json -Depth 20 -Compress
    if (-not $ok) {
        exit 2
    }
} catch {
    ([ordered]@{
        ok = $false
        action = $Action
        error = $_.Exception.Message
    } | ConvertTo-Json -Depth 10 -Compress)
    exit 1
} finally {
    if ($socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
        try {
            $null = $socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "done", [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
        } catch {}
    }
    $socket.Dispose()
    $cts.Dispose()
}
