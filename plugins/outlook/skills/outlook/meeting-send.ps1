# Send only a Zoom appointment prepared and verified by this extension.
function Get-MeetingDigest([string]$value) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','') }
    finally { $sha.Dispose() }
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
        $urls = @(Get-ZoomLinks $meeting.Body | Sort-Object -Unique)
        $valid = @($urls | Where-Object { (Get-MeetingDigest ($meeting.EntryID + '|' + $_)) -eq $proof.Value })
        $targets = @($urls | ForEach-Object { ([uri]$_).GetLeftPart([UriPartial]::Path) } | Sort-Object -Unique)
        if ($valid.Count -ne 1 -or $targets.Count -ne 1) { throw 'The verified Zoom link is missing, changed or ambiguous.' }

        $expected = @(
            foreach ($name in $required) { Get-AttendeeKey ($session.CreateRecipient($name)) 1 }
            foreach ($name in $optional) { Get-AttendeeKey ($session.CreateRecipient($name)) 2 }
        )
        $actual = @(
            foreach ($recipient in $meeting.Recipients) {
                if ($recipient.Type -eq 0 -and $recipient.AddressEntry.ID -eq $session.CurrentUser.AddressEntry.ID) { continue }
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
