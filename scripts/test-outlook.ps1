# Opt-in integration check for classic Outlook. Creates only uniquely named, recipient-free test items.
$ErrorActionPreference = 'Stop'
$script = Join-Path $PSScriptRoot '../plugins/outlook/skills/outlook/outlook.ps1'
$session = (New-Object -ComObject Outlook.Application).Session
$drafts = $session.GetDefaultFolder(16)
foreach ($operation in 'draft','meeting') {
    $subject = 'WhiteCode Outlook test ' + [guid]::NewGuid()
    $parameters = @{ Action=$operation; Subject=$subject; Body="First line. O'Brien."; Start='2027-01-06T10:00:00+01:00'; End='2027-01-06T10:30:00+01:00'; VideoUrl='https://example.org/meeting' }
    $receipt = & $script @parameters | ConvertFrom-Json
    $item = $session.GetItemFromID($receipt.id, $drafts.StoreID)
    if ($item.Subject -ne $subject -or $item.Recipients.Count -ne 0) { throw 'Unexpected test item.' }
    try {
        $folder = if ($operation -eq 'meeting') { $session.GetDefaultFolder(9) } else { $drafts }
        if (!$receipt.saved -or $item.Parent.EntryID -ne $folder.EntryID) { throw 'Unexpected Outlook folder.' }
        if (!$item.Body.Contains($parameters.Body)) { throw 'Draft body mismatch.' }
        if ($operation -eq 'meeting') {
            if ($item.StartUTC.ToString('s') -ne '2027-01-06T09:00:00' -or $item.EndUTC.ToString('s') -ne '2027-01-06T09:30:00') { throw 'Incorrect meeting times.' }
            if ($item.Class -ne 26 -or $item.MeetingStatus -ne 1 -or $item.BusyStatus -ne 0 -or !$item.Body.Contains($parameters.VideoUrl)) { throw 'Incorrect invitation.' }
            $refusal = ''
            try { $null = & $script -Action send-meeting -Id $receipt.id -ToSelf }
            catch { $refusal = $_.Exception.Message }
            if ($refusal -notlike '*No verified Zoom link*') { throw 'Expected an unverified meeting to be refused.' }
        }
    } finally { $item.Delete() }
}
# Read-only check: do not print mailbox content to the test log.
$latest = & $script -Action latest | ConvertFrom-Json
$inbox = $session.GetDefaultFolder(6).Items
$inbox.Sort('[ReceivedTime]', $true)
$firstMail = $inbox.GetFirst()
while ($null -ne $firstMail -and $firstMail.Class -ne 43) { $firstMail = $inbox.GetNext() }
if ($latest.found -ne ($null -ne $firstMail) -or ($latest.found -and $latest.id -ne $firstMail.EntryID)) {
    throw 'Latest mail is not the most recently received mail.'
}
$page = & $script -Action search -FromDate ([datetime]::Today.AddDays(-7).ToString('yyyy-MM-dd')) -Limit 1 | ConvertFrom-Json
if ($page.results.Count -gt 0) {
    $message = & $script -Action read -Id $page.results[0].id | ConvertFrom-Json
    if ($message.id -ne $page.results[0].id) { throw 'Search/read mismatch.' }
}
Write-Output 'Outlook search, mail draft and meeting draft verified.'
