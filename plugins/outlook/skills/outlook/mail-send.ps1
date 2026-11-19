# Submit the existing saved message only. The caller supplies the recipients
# authorized for this send; never derive that list from the draft itself.
function Submit-OutlookDraft($session, [string]$entryId, [string]$expectedSubject,
                            [string[]]$toList, [string[]]$ccList, [string[]]$bccList) {
    if ($entryId -notmatch '^[0-9a-fA-F]{8,4096}$') {
        throw 'Supply the saved draft ID.'
    }
    $expected = @(
        foreach ($name in $toList) { Get-AttendeeKey ($session.CreateRecipient($name)) 1 }
        foreach ($name in $ccList) { Get-AttendeeKey ($session.CreateRecipient($name)) 2 }
        foreach ($name in $bccList) { Get-AttendeeKey ($session.CreateRecipient($name)) 3 }
    )
    if (!$expected.Count) { throw 'Specify the intended mail recipients before sending.' }
    $mutex = New-Object Threading.Mutex($false,('Local\WhiteCodeOutlook-' + (Get-MeetingDigest $entryId.ToUpperInvariant())))
    $held = $false
    try {
        try { $held = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $held = $true }
        if (!$held) { throw 'Another submission is in progress for this draft.' }
        $drafts = $session.GetDefaultFolder(16)
        try { $mail = $session.GetItemFromID($entryId,$drafts.StoreID) }
        catch { throw 'Draft ID unavailable. Check Drafts, Outbox and Sent Items; do not recreate the message.' }
        if ($mail.Class -ne 43 -or $mail.Parent.EntryID -ne $drafts.EntryID -or $mail.Sent -or $mail.Submitted) {
            throw 'This item is not an unsent message in Drafts. Check Outbox and Sent Items.'
        }
        $subjectMatches = $mail.Subject -ceq $expectedSubject
        if ([string]::IsNullOrWhiteSpace($expectedSubject)) {
            $original = $mail.UserProperties.Find('WhiteCodeDraftSubjectHash')
            $subjectMatches = $null -ne $original -and $original.Value -eq (Get-MeetingDigest $mail.Subject)
        }
        if (!$mail.Saved -or !$subjectMatches) {
            throw 'Unsaved changes or subject mismatch. Review the draft before sending.'
        }
        if ($mail.UserProperties.Find('WhiteCodeSubmitAttempt')) {
            throw 'Submission already attempted. Check Outbox and Sent Items before any retry.'
        }
        $actual = @(foreach ($recipient in $mail.Recipients) { Get-AttendeeKey $recipient $recipient.Type })
        if (!$actual.Count -or (Compare-Object ($expected | Sort-Object) ($actual | Sort-Object))) {
            throw 'Outlook mail recipients or roles differ from the requested list. Draft retained.'
        }
        $attempt = $mail.UserProperties.Add('WhiteCodeSubmitAttempt',1,$false)
        $attempt.Value = [datetime]::UtcNow.ToString('o')
        $mail.Save()
        $sentSubject = $mail.Subject
        $mail.Send()
        return @{ status='submitted'; draftId=$entryId; subject=$sentSubject; verifiedRecipients=$actual.Count }
    } finally {
        if ($held) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}
