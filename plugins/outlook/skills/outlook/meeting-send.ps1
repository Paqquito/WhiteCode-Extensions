# Send only a Zoom appointment prepared and verified by this extension.
function Get-MeetingDigest([string]$value) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','') }
    finally { $sha.Dispose() }
}
function Get-ZoomLinks([string]$text) {
    $pattern = '(?i)(?:^|[\s<("\x27\[])(https://(?:[a-z0-9-]+\.)*(?:zoom\.us|zoom\.com|zoomgov\.com|zoom\.com\.cn)' +
        '/j/[0-9]{9,11}(?:[?#][^\s<>")\]\x27]*)?)(?=$|[\s<>"\x27)\]]|[.,](?=\s|$))'
    foreach ($match in [regex]::Matches($text,$pattern)) { $match.Groups[1].Value.TrimEnd('.',',') }
}

# Zoom can write the join URL to the body, the location, or both. A bare
# copy and a copy with a passcode may identify the same meeting.
function Get-ZoomJoinUrl($appointment) {
    $urls = @(Get-ZoomLinks "$($appointment.Body)`n$($appointment.Location)" | Sort-Object -Unique -CaseSensitive)
    if ($urls.Count -eq 1) { return $urls[0] }
    $targets = @($urls | ForEach-Object { ([uri]$_).GetLeftPart([UriPartial]::Path) } | Sort-Object -Unique)
    $qualified = @($urls | Where-Object { $_ -match '[?#]' })
    if ($targets.Count -eq 1 -and $qualified.Count -eq 1) { return $qualified[0] }
}

function Save-ZoomLinkProof($session,$item,[string[]]$originalLinks) {
    $until = [datetime]::UtcNow.AddSeconds(20)
    do {
        $url = Get-ZoomJoinUrl $item
        if ($url -and $url -cnotin $originalLinks) {
            $item.Save()
            $stored = $session.GetItemFromID($item.EntryID,$item.Parent.StoreID)
            if ($stored.Class -eq 26 -and $stored.MeetingStatus -eq 1 -and (Get-ZoomJoinUrl $stored) -ceq $url) {
                $proof = $stored.UserProperties.Find('WhiteCodeZoomProof')
                if (!$proof) { $proof = $stored.UserProperties.Add('WhiteCodeZoomProof',1,$false) }
                $proof.Value = Get-MeetingDigest ($stored.EntryID + '|' + $url)
                $stored.Save()
                return $true
            }
        }
        Start-Sleep -Milliseconds 250
    } while ([datetime]::UtcNow -lt $until)
    return $false
}

function Get-AttendeeKey($recipient, [int]$role) {
    if (!$recipient.Resolve()) { throw 'Unresolved or ambiguous attendee.' }
    $entry = $recipient.AddressEntry
    $address = if ($entry.Type -eq 'EX') {
        $user = $entry.GetExchangeUser()
        if ($user) { $user.PrimarySmtpAddress }
    } elseif ($entry.Type -eq 'SMTP') { $entry.Address }
    if (!$address -or $address -notmatch '^[^\s@;<>]+@[^\s@;<>]+$') {
        throw 'A unique personal SMTP address is required for each attendee.'
    }
    return "$role|$($address.ToLowerInvariant())"
}
function Submit-ZoomMeeting($session, [string]$entryId, [string[]]$required, [string[]]$optional) {
    if ($entryId -notmatch '^[0-9a-fA-F]{8,4096}$' -or !(@($required).Count + @($optional).Count)) {
        throw 'Supply the saved meeting ID and its intended attendees.'
    }
    $mutex = New-Object Threading.Mutex($false,('Local\WhiteCodeOutlook-' + (Get-MeetingDigest $entryId.ToUpperInvariant())))
    $held = $false
    try {
        try { $held = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $held = $true }
        if (!$held) { throw 'Another submission is in progress for this appointment.' }
        $calendar = $session.GetDefaultFolder(9)
        $meeting = $session.GetItemFromID($entryId,$calendar.StoreID)
        if ($meeting.Class -ne 26 -or $meeting.MeetingStatus -ne 1 -or $meeting.Parent.EntryID -ne $calendar.EntryID) {
            throw 'Expected an organizer appointment in the default calendar.'
        }
        if (!$meeting.Saved) { throw 'Save your Outlook edits before submission.' }
        $proof = $meeting.UserProperties.Find('WhiteCodeZoomProof')
        if (!$proof) { throw 'No verified Zoom link receipt for this appointment.' }
        $url = Get-ZoomJoinUrl $meeting
        if (!$url -or (Get-MeetingDigest ($meeting.EntryID + '|' + $url)) -ne $proof.Value) { throw 'The verified Zoom link is missing, changed or ambiguous.' }

        $expected = @(
            foreach ($name in $required) { Get-AttendeeKey ($session.CreateRecipient($name)) 1 }
            foreach ($name in $optional) { Get-AttendeeKey ($session.CreateRecipient($name)) 2 }
        )
        $actual = @(
            foreach ($recipient in $meeting.Recipients) {
                # After Zoom saves the appointment Outlook can expose the organizer as
                # olRequired too. PidTagRecipientFlags identifies the organizer row.
                if ($recipient.AddressEntry.ID -eq $session.CurrentUser.AddressEntry.ID) {
                    $flags = $recipient.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x5FFD0003')
                    if ($recipient.Type -eq 0 -or ($flags -band 2)) { continue }
                }
                Get-AttendeeKey $recipient $recipient.Type
            }
        )
        if (!$actual.Count -or (Compare-Object ($expected | Sort-Object) ($actual | Sort-Object))) {
            throw 'Outlook attendees or roles differ from the requested list. Invitation retained.'
        }
        # FInvited also detects manual sends through the Outlook inspector.
        if ($meeting.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8229000B')) {
            return @{status='already-submitted'; id=$entryId}
        }
        $attempt = $meeting.UserProperties.Find('WhiteCodeSubmitAttempt')
        if ($attempt) { throw 'Submission already attempted. Check Outbox and Sent Items before any manual retry.' }
        $attempt = $meeting.UserProperties.Add('WhiteCodeSubmitAttempt',1,$false)
        $attempt.Value = [datetime]::UtcNow.ToString('o')
        $meeting.Save()
        $meeting.Send()
        return @{status='submitted'; id=$entryId; verifiedParticipants=$actual.Count; verifiedZoomLink=$true}
    } finally {
        if ($held) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}
