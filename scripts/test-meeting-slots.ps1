$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../plugins/outlook/skills/outlook/availability.ps1')
foreach ($case in @(
    @{meeting=1; response=1; busy=0; conflict=$true}, # Legacy organizer meeting
    @{meeting=3; response=3; busy=0; conflict=$true}, # Accepted meeting marked free
    @{meeting=0; response=0; busy=0; conflict=$false},
    @{meeting=0; response=0; busy=1; conflict=$true},
    @{meeting=0; response=0; busy=2; conflict=$true},
    @{meeting=5; response=1; busy=2; conflict=$false},
    @{meeting=7; response=3; busy=2; conflict=$false},
    @{meeting=3; response=4; busy=2; conflict=$false}
)) {
    $appointment = @{MeetingStatus=$case.meeting; ResponseStatus=$case.response; BusyStatus=$case.busy}
    if ((Test-LocalCalendarConflict $appointment) -ne $case.conflict) { throw "Incorrect local conflict: $($case | ConvertTo-Json -Compress)" }
}
# Use a future Monday independent of the workstation date.
$monday = [datetime]::Today.AddDays(1)
while ($monday.DayOfWeek -ne 'Monday') { $monday = $monday.AddDays(1) }
$requested = [datetimeoffset]$monday.AddHours(10)
$finish = $requested.AddMinutes(45)
$recipient = [pscustomobject]@{ calls=0 }
$recipient | Add-Member ScriptMethod FreeBusy {
    param($day,$resolution,$complete)
    $this.calls++
    if ($resolution -ne 30 -or !$complete) { throw 'Unexpected remote availability request.' }
    # Busy in the second half-hour of the requested 45-minute meeting.
    return ('0' * 21) + '2' + ('0' * 314)
}
$remote = @{ person='Optional attendee'; recipient=$recipient; resolved=$true; days=@{}; unavailable=$false }
$calendars = @{ local=@{ person='Organizer'; known=$true; intervals=@(@{start=$finish.UtcDateTime; end=$finish.AddMinutes(45).UtcDateTime}) }; others=@($remote) }
$states = @(Get-MeetingStates $calendars $requested $finish)
if ($states[0].availability -ne 'free' -or $states[1].availability -ne 'occupied-or-tentative') { throw 'Incorrect interval boundaries or remote conflict.' }
$choice = New-MeetingConflictChoice $calendars $states $requested $finish 'test'
if ($choice.saved -or $choice.questions[0].options.Count -ne 4 -or $choice.questions[0].question -notlike '*Optional attendee*') { throw 'Missing question or attendees.' }
foreach ($slot in $choice.suggestions) {
    $start = [datetimeoffset]::Parse($slot.start)
    $end = [datetimeoffset]::Parse($slot.end)
    if ($start -lt $finish.AddMinutes(45) -or $end-$start -ne $finish-$requested -or $slot.unknown.Count) { throw 'Invalid shared slot.' }
}
if ($recipient.calls -ne 1) { throw 'The Exchange grid should be reused for every candidate.' }
if ($choice.suggestions[0].label -notmatch '^\d{2}:\d{2}' -or $choice.questions[0].options[0].label -cne $choice.suggestions[0].label) { throw 'Same-day choice lost its selectable time label.' }
$lateStart = [datetimeoffset]$monday.AddHours(17)
$nextDay = @(Find-MeetingAlternatives $calendars $lateStart $lateStart.AddHours(1))[0]
if ($nextDay.label -notmatch ([regex]::Escape($monday.AddDays(1).ToString('ddd d MMM yyyy',[cultureinfo]'fr-FR')))) { throw 'Missing date on a later-day choice.' }
if (!$choice.questions[0].require_answer -or $choice.questions[0].options[$choice.suggestions.Count].description -notmatch '10:00') { throw 'Missing explicit choice to keep the requested time.' }
$remote.days.Clear()
$remote.unavailable = $true
$calendars.local.intervals = @(@{start=$requested.UtcDateTime; end=$finish.UtcDateTime})
$states = @(Get-MeetingStates $calendars $requested $finish)
if ($states[1].availability -ne 'unknown') { throw 'Unavailable Exchange data is not free.' }
$choice = New-MeetingConflictChoice $calendars $states $requested $finish 'test'
if (!$choice.suggestions.Count -or $choice.suggestions[0].unknown -notcontains 'Optional attendee') { throw 'Missing provisional slot label.' }
$calendars.local.intervals = @(@{start=$requested.UtcDateTime; end=$finish.AddDays(8).UtcDateTime})
$choice = New-MeetingConflictChoice $calendars $states $requested $finish 'test'
if ($choice.suggestions.Count -or $choice.questions[0].options.Count -ne 2) { throw 'Expected keep or choose another time.' }
'Meeting availability, shared suggestions, unknown calendars and no-alternative choices passed.'
