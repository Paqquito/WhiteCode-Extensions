[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('status','folders','latest','search','read','draft','send-draft','availability','meeting','send-meeting')][string]$Action,
    [ValidateSet('inbox','sent')][string]$Folder = 'inbox',
    [string[]]$Subfolder = @(),
    [string]$FromDate, [string]$UntilDate, [string]$Query = '',
    [ValidateRange(1,30)][int]$Limit = 15,
    [ValidateRange(0,2147483646)][int]$Offset = 0,
    [string]$Id, [string]$Subject, [string]$Body,
    [string[]]$To = @(), [string[]]$Cc = @(), [string[]]$Bcc = @(),
    [string]$Start, [string]$End, [string]$Location, [string]$VideoUrl,
    [ValidateSet('Zoom','Tixeo')][string]$VideoProvider,
    [switch]$Open, [switch]$SendInvitation, [switch]$ToSelf
)
$ErrorActionPreference = 'Stop'
$culture = [Globalization.CultureInfo]::InvariantCulture
. (Join-Path $PSScriptRoot 'meeting-send.ps1')
. (Join-Path $PSScriptRoot 'mail-send.ps1')
. (Join-Path $PSScriptRoot 'zoom-time-alert.ps1')

function Get-Day([string]$value, [datetime]$fallback) {
    if (!$value) { return $fallback }
    return [datetime]::ParseExact($value, 'yyyy-MM-dd', $culture)
}
function Get-Instant([string]$value) {
    if ($value -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?(Z|[+-]\d{2}:\d{2})$') {
        throw 'Specify an ISO date and time with a UTC offset.'
    }
    return [datetimeoffset]::Parse($value, $culture)
}
function Add-Recipients($item, [string[]]$addresses, [int]$kind) {
    foreach ($address in $addresses) {
        if ([string]::IsNullOrWhiteSpace($address) -or $address -match '[;\r\n]') {
            throw 'Supply one recipient per array entry.'
        }
        $recipient = $item.Recipients.Add($address.Trim())
        $recipient.Type = $kind
        if (!$recipient.Resolve()) { throw "Unresolved or ambiguous recipient: $address" }
    }
}
function Clip([string]$text, [int]$length) {
    return $text.Substring(0,[Math]::Min($text.Length,$length))
}
function Invoke-Outlook {
    if ($VideoProvider -and ($Action -ne 'meeting' -or $VideoUrl)) {
        throw 'Use VideoProvider only for a new meeting without a supplied link.'
    }
    if ($Action -eq 'meeting') {
        $starts = Get-Instant $Start
        $ends = Get-Instant $End
        if ($ends -le $starts) { throw 'End must be later than Start.' }
        if ($VideoProvider -eq 'Zoom' -and $ends -le [datetimeoffset]::Now) {
            $result = @{ action='meeting'; saved=$false; opened=$false; videoProvider='Zoom'; videoStatus='meeting-ended'; verifiedZoomLink=$false
                detail='The time slot has ended. No appointment or Zoom link was created. No invitation was sent. Ask only for a new time.' }
            if ($SendInvitation) { $result.submission = @{status='not-submitted'} }
            return $result
        }
    }
    try { $app = New-Object -ComObject Outlook.Application }
    catch { throw 'Open classic Outlook for Windows with a configured profile.' }
    $session = $app.Session
    if ($session.Accounts.Count -lt 1) { throw 'No configured Outlook account.' }
    if ($SendInvitation -and ($Action -ne 'meeting' -or !$VideoProvider)) { throw 'SendInvitation requires a new Zoom or Tixeo meeting.' }
    if ($ToSelf) {
        if ($Action -notin 'meeting','send-meeting','draft','send-draft' -or $To.Count -or $Cc.Count -or $Bcc.Count) { throw 'ToSelf requires a mail draft or meeting with no other recipients.' }
        $self = Get-AttendeeKey $session.CurrentUser 1
        $To = @($self.Substring(2))
    }
    if ($Bcc.Count -and $Action -notin 'draft','send-draft') { throw 'Bcc applies only to mail drafts.' }
    if ($Action -eq 'send-draft') { return Submit-OutlookDraft $session $Id $Subject $To $Cc $Bcc }
    if ($Action -eq 'send-meeting') { return Submit-VideoMeeting $session $Id $To $Cc }
    if ($SendInvitation -and !$To.Count -and !$Cc.Count) { throw 'Specify the intended recipients before scheduling.' }
    if ($Action -eq 'status') {
        $video = @($app.COMAddIns | Where-Object { $_.Connect -and $_.ProgId -match 'zoom|teams|tixeo|webex' } | ForEach-Object { $_.ProgId })
        return @{ outlookVersion=$app.Version; accountCount=$session.Accounts.Count; videoAddIns=$video }
    }
    if ($Action -in 'folders','latest','search','read') {
        $folderType = if ($Folder -eq 'inbox') { 6 } else { 5 }
        $source = $session.GetDefaultFolder($folderType)
        foreach ($part in $Subfolder) { $source = $source.Folders.Item($part) }
        if ($Action -eq 'latest') {
            $property = if ($Folder -eq 'sent') { 'SentOn' } else { 'ReceivedTime' }
            $items = $source.Items
            $items.Sort("[$property]",$true)
            for ($item=$items.GetFirst(); $null -ne $item; $item=$items.GetNext()) {
                if ($item.Class -eq 43) {
                    return @{ found=$true; id=$item.EntryID; subject=$item.Subject; sender=$item.SenderName; at=([datetime]$item.$property).ToString('o') }
                }
            }
            return @{ found=$false }
        }
        if ($Action -eq 'folders') {
            return @{ folder=$Folder; subfolder=$Subfolder; children=@($source.Folders | ForEach-Object { $_.Name }) }
        }
        if ($Action -eq 'read') {
            if ($Id -notmatch '^[0-9a-fA-F]{8,4096}$') { throw 'Invalid Outlook item ID.' }
            $item = $session.GetItemFromID($Id,$source.StoreID)
            if ($item.Class -ne 43 -or $item.Parent.EntryID -ne $source.EntryID) { throw 'Message does not belong to the requested folder.' }
            $text = [string]$item.Body
            $pos = [Math]::Min($Offset,$text.Length)
            $chunk = Clip ($text.Substring($pos)) 8000
            $next = if ($pos+$chunk.Length -lt $text.Length) { $pos+$chunk.Length }
            return @{ id=$Id; subject=$item.Subject; sender=$item.SenderName; body=$chunk; nextOffset=$next }
        }
        $today = [datetime]::Today
        $first = Get-Day $FromDate $today
        $last = Get-Day $UntilDate ($today.AddDays(1))
        if ($last -le $first) { throw 'UntilDate must be later than FromDate (exclusive).' }
        $property = if ($Folder -eq 'sent') { 'SentOn' } else { 'ReceivedTime' }
        # Jet date literals use the Outlook user's regional short date and time.
        $filter = "[$property] >= '" + $first.ToString('g') + "' AND [$property] < '" + $last.ToString('g') + "'"
        $items = $source.Items.Restrict($filter)
        $items.Sort("[$property]",$true)
        $count = $items.Count
        $cursor = $Offset
        $hits = @()
        while ($cursor -lt $count -and $cursor-$Offset -lt 100 -and $hits.Count -lt $Limit) {
            $cursor++
            $item = $items.Item($cursor)
            if ($item.Class -ne 43) { continue }
            $text = [string]$item.Body
            $match = $text.IndexOf($Query,[StringComparison]::OrdinalIgnoreCase)
            if ($Query -and $match -lt 0 -and
                ("$($item.Subject) $($item.SenderName) $($item.To) $($item.CC)").IndexOf($Query,[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            $excerpt = Clip ($text.Substring([Math]::Max(0,$match-60))) 500
            $hits += @{ id=$item.EntryID; subject=$item.Subject; sender=$item.SenderName; at=([datetime]$item.$property).ToString('o'); excerpt=$excerpt }
        }
        $next = if ($cursor -lt $count) { $cursor }
        return @{ folder=$Folder; subfolder=$Subfolder; from=$first.ToString('yyyy-MM-dd'); until=$last.ToString('yyyy-MM-dd'); results=$hits; nextOffset=$next }
    }
    if ($Action -eq 'availability') {
        $begin = (Get-Instant $Start).LocalDateTime
        $finish = (Get-Instant $End).LocalDateTime
        if ($begin.Date -ne $finish.Date -or $finish -le $begin -or
            [TimeZoneInfo]::Local.GetUtcOffset($begin.Date) -ne [TimeZoneInfo]::Local.GetUtcOffset($finish)) {
            throw 'Choose a positive interval within one local day with a stable UTC offset.'
        }
        if ($To.Count -gt 20) { throw 'Check at most 20 people per request.' }
        $firstSlot = [int][math]::Floor($begin.TimeOfDay.TotalMinutes / 30)
        $after = [int][math]::Ceiling($finish.TimeOfDay.TotalMinutes / 30)
        # The personal calendar is locally available even when Exchange free/busy is offline.
        $ownState = 'unknown'
        try {
            $calendar = $session.GetDefaultFolder(9).Items
            $calendar.Sort('[Start]')
            $calendar.IncludeRecurrences = $true
            $overlap = $calendar.Restrict("[Start] < '" + $finish.ToString('g') + "' AND [End] > '" + $begin.ToString('g') + "'")
            $ownState = 'free'
            for ($entry=$overlap.GetFirst(); $null -ne $entry; $entry=$overlap.GetNext()) {
                if ($entry.BusyStatus -ne 0) { $ownState = 'occupied-or-tentative'; break }
            }
        } catch { $ownState = 'unknown' }
        $people = @()
        foreach ($address in $To) { $people += $session.CreateRecipient($address) }
        $states = foreach ($person in $people) {
            $state = 'unknown'
            $reason = 'unresolved-recipient'
            try {
                if ($person.Resolve()) {
                    $reason = 'exchange-unavailable'
                    try { $slots = [string]$person.FreeBusy($begin.Date,30,$true) }
                    catch { $slots = [string]$person.AddressEntry.GetFreeBusy($begin.Date,30,$true) }
                    $reason = 'no-published-data'
                    if ($slots.Length -ge $after) {
                        $window = $slots.Substring($firstSlot,$after-$firstSlot)
                        if ($window -match '^[0]+$') { $state = 'free' }
                        elseif ($window -match '[123]') { $state = 'occupied-or-tentative' }
                    }
                }
            } catch { $state = 'unknown' }
            $row = @{ person=$person.Name; availability=$state }
            if ($state -eq 'unknown') { $row.reason = $reason }
            $row
        }
        $organizer = @{ person=$session.CurrentUser.Name; availability=$ownState }
        return @{ start=$Start; end=$End; resolutionMinutes=30; people=@($organizer) + @($states | Where-Object { $_ }) }
    }
    if ([string]::IsNullOrWhiteSpace($Subject)) { throw 'A subject is required.' }
    $drafts = $session.GetDefaultFolder(16)
    if ($Action -eq 'meeting') {
        if ($VideoUrl) {
            $uri = [uri]$VideoUrl
            if (!$uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.UserInfo -or $VideoUrl -match '\s') { throw 'Supply an HTTPS meeting link.' }
        }
        # Outlook converts appointments moved to the mail Drafts folder into plain messages.
        $item = $app.CreateItem(1)
        $item.MeetingStatus = 1
        $item.BusyStatus = 0
        $item.StartUTC = $starts.UtcDateTime
        $item.EndUTC = $ends.UtcDateTime
        $item.Location = if ($Location) { $Location } else { $VideoUrl }
        if ($VideoUrl) { $Body = "$Body" + [Environment]::NewLine + $VideoUrl }
    } else {
        $item = $drafts.Items.Add('IPM.Note')
    }
    $item.Subject = $Subject
    if ($Action -eq 'draft') {
        $original = $item.UserProperties.Add('WhiteCodeDraftSubjectHash',1,$false)
        $original.Value = Get-MeetingDigest $Subject
    }
    $item.Body = [string]$Body
    Add-Recipients $item $To 1
    Add-Recipients $item $Cc 2
    if ($Bcc.Count) { Add-Recipients $item $Bcc 3 }
    $item.Save()
    $storeId = $item.Parent.StoreID
    $result = @{ id=$item.EntryID; folder=$item.Parent.Name; action=$Action; saved=$true; opened=$false; recipients=$item.Recipients.Count }
    if ($Open -or $VideoProvider) {
        try { $item.Display($false); $result.opened=$true }
        catch { $result.openError='Saved. Open the item in Outlook.' }
    }
    $handle = 0
    if ($VideoProvider) {
        $origin = if ($VideoProvider -eq 'Tixeo') { Get-TixeoOrigin }
        $originalLinks = @(if ($VideoProvider -eq 'Zoom') { Get-ZoomLinks $Body $Location } else { Get-TixeoLinks "$Body`n$Location" $origin })
        $result.videoProvider = $VideoProvider
        $result.videoStatus = 'manual-action-required'
        if ($result.opened) {
            try {
                if (!('MeetingWindow' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'MeetingWindow.cs') }
                $handle = [MeetingWindow]::Handle($item.GetInspector)
                $worker = Join-Path $PSScriptRoot 'video-ribbon.ps1'
                $ps = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
                $process = Start-Process -FilePath $ps -ArgumentList @('-NoProfile','-NonInteractive','-File',('"' + $worker + '"'),'-Window',$handle,'-Provider',$VideoProvider) -PassThru -WindowStyle Hidden
                try {
                    if (!$process.WaitForExit(20000)) {
                        $process.Kill()
                        $result.videoStatus = 'check-open-invitation'
                    } elseif ($process.ExitCode -eq 4) {
                        $result.videoStatus = 'zoom-review-open'
                    } elseif ($process.ExitCode -eq 5) {
                        $result.videoStatus = 'addin-command-disabled'
                    } elseif ($process.ExitCode -eq 0) {
                        $result.videoStatus = 'check-open-invitation'
                        if ($VideoProvider -eq 'Tixeo') { $item.Save() }
                        if (Save-MeetingLinkProof $session $item $VideoProvider $origin $originalLinks) {
                            $result.videoStatus = $VideoProvider.ToLowerInvariant() + '-link-verified'
                        }
                    }
                } finally { $process.Dispose() }
            } catch { $result.videoStatus = 'check-open-invitation' }
        }
    }
    if ($VideoProvider -eq 'Zoom') {
        $result.verifiedZoomLink = $result.videoStatus -eq 'zoom-link-verified'
        if ($handle -and !$result.verifiedZoomLink) {
            try {
                if (Close-ZoomTimeAlert $handle) { $result.videoStatus = 'meeting-ended' }
            } catch { Write-Verbose $_.Exception.Message }
        }
        if ($result.opened -and (!$Open -or $SendInvitation -or $result.videoStatus -eq 'meeting-ended')) {
            # Flush the add-in's edits. Submission below reopens the calendar
            # item by ID and checks the persisted link and participants.
            $result.windowClosed = Close-MeetingWindow $app $storeId $result.id -SaveDraftEdits
            $result.opened = !$result.windowClosed
        }
    }
    if ($SendInvitation) {
        $result.submission = @{ status='not-submitted' }
        if ($result.videoStatus -in 'zoom-link-verified','tixeo-link-verified') {
            try { $result.submission = Submit-VideoMeeting $session $result.id $To $Cc }
            catch { $result.submission = @{status='check-outlook'; detail=$_.Exception.Message} }
        }
    }
    if ($SendInvitation -and $result.opened) {
        if ($result.submission -and $result.submission.ContainsKey('windowClosed')) {
            $result.windowClosed = $result.submission.windowClosed
        } elseif (!$result.ContainsKey('windowClosed')) {
            $result.windowClosed = Close-MeetingWindow $app $storeId $result.id
        }
        if ($result.windowClosed) { $result.opened = $false }
    }
    return $result
}
Invoke-Outlook | ConvertTo-Json -Compress -Depth 5
