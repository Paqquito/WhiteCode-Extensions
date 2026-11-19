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
    [switch]$Open, [switch]$SendInvitation, [switch]$ToSelf,
    [string]$KeepSlot
)
$ErrorActionPreference = 'Stop'
$culture = [Globalization.CultureInfo]::InvariantCulture
. (Join-Path $PSScriptRoot 'meeting-send.ps1')
. (Join-Path $PSScriptRoot 'availability.ps1')
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
        $starts = Get-Instant $Start
        $ends = Get-Instant $End
        $begin = $starts.LocalDateTime
        $finish = $ends.LocalDateTime
        if ($begin.Date -ne $finish.Date -or $finish -le $begin -or
            [TimeZoneInfo]::Local.GetUtcOffset($begin.Date) -ne [TimeZoneInfo]::Local.GetUtcOffset($finish)) {
            throw 'Choose a positive interval within one local day with a stable UTC offset.'
        }
        if ($To.Count -gt 20) { throw 'Check at most 20 people per request.' }
        $calendars = Read-MeetingCalendars $session $To $begin.Date $begin.Date.AddDays(1)
        return @{ start=$Start; end=$End; resolutionMinutes=30; people=@(Get-MeetingStates $calendars $starts $ends) }
    }
    if ([string]::IsNullOrWhiteSpace($Subject)) { throw 'A subject is required.' }
    $drafts = $session.GetDefaultFolder(16)
    if ($Action -eq 'meeting') {
        if ($VideoUrl) {
            $uri = [uri]$VideoUrl
            if (!$uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.UserInfo -or $VideoUrl -match '\s') { throw 'Supply an HTTPS meeting link.' }
        }
        $selectedLabel = $null
        if ($env:TOOL_TURN_ID -and $env:TOOL_CHOICE_PIPE) {
            $requestScope = ConvertTo-Json -Compress -InputObject ([ordered]@{turn=$env:TOOL_TURN_ID; subject=$Subject; duration=($ends-$starts).Ticks
                required=@($To | Sort-Object); optional=@($Cc | Sort-Object); provider=$VideoProvider; url=$VideoUrl; send=$SendInvitation.IsPresent})
            $originalKey = Get-MeetingDigest ($requestScope + '|' + $starts.UtcTicks)
            $stored = Find-PreparedMeeting $session $Subject $starts $originalKey
            if ($stored) {
                $receipt = @{action='meeting'; status='already-saved'; saved=$true; id=$stored.EntryID; start=([datetimeoffset]$stored.Start).ToString('o')
                    end=([datetimeoffset]$stored.End).ToString('o'); recipients=$stored.Recipients.Count}
                if ($SendInvitation) {
                    $submitted = $stored.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8229000B')
                    $receipt.submission = @{status=$(if ($submitted) { 'already-submitted' } else { 'not-submitted' })}
                }
                return $receipt
            }
        }
        $confirmedSlot = if (!$env:TOOL_CHOICE_PIPE) { $KeepSlot }
        while ($true) {
            $firstDay = $starts.LocalDateTime.Date
            $until = $firstDay.AddDays(7)
            if ($ends.LocalDateTime -ge $until) { $until = $ends.LocalDateTime.Date.AddDays(1) }
            $calendars = Read-MeetingCalendars $session ($To + $Cc) $firstDay $until
            $states = @(Get-MeetingStates $calendars $starts $ends)
            if ('occupied-or-tentative' -notin $states.availability) { break }
            # A reply for one time, attendee list or set of conflicts cannot
            # silently override a newer conflict. Alternatives are checked anew.
            $decision = [ordered]@{ subject=$Subject; start=$starts.UtcTicks; end=$ends.UtcTicks; to=$To; cc=$Cc
                provider=$VideoProvider; send=$SendInvitation.IsPresent; organizer=$session.CurrentUser.AddressEntry.ID
                availability=@($states | ForEach-Object { @($_.person,$_.availability) -join '|' }) }
            $token = (Get-MeetingDigest (ConvertTo-Json -InputObject $decision -Compress)).Substring(0,16)
            if ($confirmedSlot -ceq $token) { break }
            $choice = New-MeetingConflictChoice $calendars $states $starts $ends $token
            if (!$env:TOOL_CHOICE_PIPE) { return $choice }
            $reply = Read-MeetingChoice $choice.questions[0]
            if ($reply.other) { return @{status='alternative-entered'; requested=$reply.other; saved=$false; opened=$false} }
            if ($reply.none -or @($reply.labels).Count -ne 1) { return @{status='choice-skipped'; saved=$false; opened=$false} }
            $selectedLabel = $reply.labels[0]
            if ($selectedLabel -ceq $choice.keepLabel) { $confirmedSlot = $token; continue }
            $slot = $choice.suggestions | Where-Object { $_.label -ceq $selectedLabel } | Select-Object -First 1
            if (!$slot) { return @{status='ask-for-new-time'; saved=$false; opened=$false} }
            $starts = Get-Instant $slot.start
            $ends = Get-Instant $slot.end
            $confirmedSlot = $null
        }
        # Outlook converts appointments moved to the mail Drafts folder into plain messages.
        $item = $app.CreateItem(1)
        $item.MeetingStatus = 1
        $item.BusyStatus = 0
        $item.StartUTC = $starts.UtcDateTime
        $item.EndUTC = $ends.UtcDateTime
        if ($selectedLabel -and $originalKey) {
            $selectedKey = Get-MeetingDigest ($requestScope + '|' + $starts.UtcTicks)
            $marker = $item.UserProperties.Add('WhiteCodeScheduleChoice',1,$false)
            $marker.Value = "$originalKey $selectedKey"
        }
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
    if ($Action -eq 'meeting') {
        $result.start = $starts.ToString('yyyy-MM-ddTHH:mm:sszzz')
        $result.end = $ends.ToString('yyyy-MM-ddTHH:mm:sszzz')
        if ($selectedLabel) { $result.userSelected = $selectedLabel }
        $result.unknownAvailability = @($states | Where-Object availability -eq 'unknown' | ForEach-Object { $_.person })
    }
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
