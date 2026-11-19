# Opt-in integration check for classic Outlook. Creates only uniquely named, recipient-free test items.
$ErrorActionPreference = 'Stop'
$script = Join-Path $PSScriptRoot '../plugins/outlook/skills/outlook/outlook.ps1'
. (Join-Path $PSScriptRoot '../plugins/outlook/skills/outlook/meeting-send.ps1')
$joinUrl = 'https://example.zoom.us/j/12345678901?pwd=AbCd'
$bare = $joinUrl.Split('?')[0]
foreach ($sample in @(
    @{ Body='Agenda'; Location=$joinUrl; Expected=$joinUrl },
    @{ Body="Join ($joinUrl)"; Location=$bare; Expected=$joinUrl },
    @{ Body=$joinUrl; Location=$joinUrl; Expected=$joinUrl },
    @{ Body=$joinUrl; Location=$joinUrl.Replace('AbCd','abcd'); Expected=$null },
    @{ Body=$joinUrl; Location=$joinUrl.Replace('12345678901','22345678901'); Expected=$null },
    @{ Body="https://example.org/?next=$joinUrl"; Location=''; Expected=$null },
    @{ Body='https://zoom.us.example.org/j/12345678901'; Location=''; Expected=$null }
)) {
    if ((Get-ZoomJoinUrl $sample) -cne $sample.Expected) { throw 'Unexpected Zoom join URL selection.' }
}
$origin = 'https://video.example.org'
$tixeo = "$origin/meet/meeting/join?umi=Example123"
foreach ($sample in @(
    @{ Body=$tixeo; Location=$tixeo; Expected=$tixeo },
    @{ Body='Agenda'; Location=$tixeo; Expected=$tixeo },
    @{ Body=$tixeo; Location=$tixeo.Replace('Example123','Example456'); Expected=$null },
    @{ Body=$tixeo.Replace('Example123',''); Location=''; Expected=$null },
    @{ Body="$tixeo&umi=Second"; Location=''; Expected=$null },
    @{ Body=$tixeo.Replace('video.example.org','video.example.org.invalid'); Location=''; Expected=$null },
    @{ Body=$tixeo.Replace('https://','https://person@'); Location=''; Expected=$null },
    @{ Body="https://example.net/?next=$tixeo"; Location=''; Expected=$null }
)) {
    if ((Get-MeetingJoinUrl $sample Tixeo $origin) -cne $sample.Expected) { throw 'Unexpected Tixeo join URL selection.' }
}
$session = (New-Object -ComObject Outlook.Application).Session
$drafts = $session.GetDefaultFolder(16)
foreach ($operation in 'draft','meeting') {
    $subject = 'WhiteCode Outlook test ' + [guid]::NewGuid()
    $parameters = @{ Action=$operation; Subject=$subject; Body="First line. O'Brien."; Start='2027-01-06T10:00:00+01:00'; End='2027-01-06T10:30:00+01:00'; VideoUrl='https://example.org/meeting' }
    if ($operation -eq 'meeting') {
        $receipt = & (Join-Path $PSScriptRoot '../plugins/outlook/skills/visio/meeting.ps1') -Action meeting -Subject $subject -Body $parameters.Body -Start $parameters.Start -End $parameters.End -VideoUrl $parameters.VideoUrl | ConvertFrom-Json
    } else { $receipt = & $script @parameters | ConvertFrom-Json }
    $item = $session.GetItemFromID($receipt.id, $drafts.StoreID)
    if ($item.Subject -ne $subject -or $item.Recipients.Count -ne 0) { throw 'Unexpected test item.' }
    try {
        $folder = if ($operation -eq 'meeting') { $session.GetDefaultFolder(9) } else { $drafts }
        if (!$receipt.saved -or $item.Parent.EntryID -ne $folder.EntryID) { throw 'Unexpected Outlook folder.' }
        if (!$item.Body.Contains($parameters.Body)) { throw 'Draft body mismatch.' }
        if ($operation -eq 'draft') {
            $recipient = $item.Recipients.Add($session.CurrentUser.Name)
            $recipient.Type = 3
            if (!$recipient.Resolve()) { throw 'Could not resolve the local test account.' }
            $item.Save()
            $refusal = ''
            try { $null = & $script -Action send-draft -Id $receipt.id -Subject $subject -ToSelf }
            catch { $refusal = $_.Exception.Message }
            if ($refusal -notlike '*mail recipients or roles differ*' -or $item.UserProperties.Find('WhiteCodeSubmitAttempt')) {
                throw 'Unexpected handling of a changed Bcc recipient.'
            }
        }
        if ($operation -eq 'meeting') {
            if ($item.StartUTC.ToString('s') -ne '2027-01-06T09:00:00' -or $item.EndUTC.ToString('s') -ne '2027-01-06T09:30:00') { throw 'Incorrect meeting times.' }
            if ($item.Class -ne 26 -or $item.MeetingStatus -ne 1 -or $item.BusyStatus -ne 0 -or !$item.Body.Contains($parameters.VideoUrl)) { throw 'Incorrect invitation.' }
            $refusal = ''
            try { $null = & $script -Action send-meeting -Id $receipt.id -ToSelf }
            catch { $refusal = $_.Exception.Message }
            if ($refusal -notlike '*verified video link receipt*') { throw 'Expected an unverified meeting to be refused.' }
            # Exercise the real Outlook inspector on this test appointment only.
            $item.Display($false)
            $store = $folder.StoreID
            if (!(Find-MeetingInspector $session.Application $store $receipt.id)) { throw 'Test inspector not found.' }
            if (!(Close-SubmittedMeeting $session.Application $store $receipt.id)) { throw 'Test inspector stayed open.' }
            # Tixeo link in the persisted item, then changed underneath its receipt.
            $item = $session.GetItemFromID($receipt.id,$store)
            $item.Body = $tixeo
            $item.Save()
            if (!(Save-MeetingLinkProof $session $item Tixeo $origin @())) { throw 'Tixeo receipt missing.' }
            $item = $session.GetItemFromID($receipt.id,$store)
            $item.Body = ''
            $item.Save()
            $refusal = ''
            try { $null = & $script -Action send-meeting -Id $receipt.id -ToSelf } catch { $refusal = $_.Exception.Message }
            if ($refusal -notlike '*verified Tixeo link is missing*' -or $item.UserProperties.Find('WhiteCodeSubmitAttempt')) { throw 'Expected changed link to block submission.' }
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
