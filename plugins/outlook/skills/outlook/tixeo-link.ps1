# The client can use a private server. Read only the selected server and its
# address; never include the profile or invitation URL in a receipt.
function Get-TixeoOrigin {
    $path = Join-Path $env:LOCALAPPDATA 'tixeoclient/config.ini'
    if (!(Test-Path -LiteralPath $path)) { return 'https://cloud.tixeo.com' }
    $servers = @{}
    $section = $selected = ''
    foreach ($line in [IO.File]::ReadLines($path)) {
        if ($line -match '^\s*\[([^\]]+)\]') { $section = $Matches[1]; continue }
        if ($section -eq 'Server' -and $line -match '^\s*Selected\s*=\s*(\d+)\s*$') { $selected = $Matches[1] }
        if ($section -match '^Server\.\d+$' -and $line -match '^\s*Address\s*=\s*(\S+)\s*$') { $servers[$section] = $Matches[1] }
    }
    $address = $servers["Server.$selected"]
    if (!$selected) { return 'https://cloud.tixeo.com' }
    $uri = $null
    if ([uri]::TryCreate($address,[UriKind]::Absolute,[ref]$uri) -and $uri.Scheme -eq 'https' -and !$uri.UserInfo) {
        return $uri.GetLeftPart([UriPartial]::Authority)
    }
}

function Get-TixeoLinks([string]$text,[string]$origin) {
    if (!$origin) { return }
    $pattern = '(?i)(?:^|[\s<("\x27\[])(https://[^\s<>"\x27)\]]+)'
    foreach ($match in [regex]::Matches($text,$pattern)) {
        $url = $match.Groups[1].Value.TrimEnd('.',',')
        $uri = $null
        if (![uri]::TryCreate($url,[UriKind]::Absolute,[ref]$uri) -or $uri.UserInfo -or $uri.Fragment -or
            $uri.GetLeftPart([UriPartial]::Authority) -ne $origin -or $uri.AbsolutePath -cne '/meet/meeting/join') { continue }
        $keys = @($uri.Query.TrimStart('?').Split('&') | Where-Object { $_ -cmatch '^umi(=|$)' })
        if ($keys.Count -eq 1 -and $keys[0] -cmatch '^umi=[A-Za-z0-9._~%+=/-]+$') { $url }
    }
}
