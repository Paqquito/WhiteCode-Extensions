# One local recurrence query and reusable Exchange free/busy grids per meeting.
function Read-MeetingCalendars($session,[string[]]$addresses,[datetime]$firstDay,[datetime]$until) {
    $local = @{ person=$session.CurrentUser.Name; intervals=@(); known=$false }
    try {
        $items = $session.GetDefaultFolder(9).Items
        $items.Sort('[Start]')
        $items.IncludeRecurrences = $true
        $range = $items.Restrict("[Start] < '" + $until.ToString('g') + "' AND [End] > '" + $firstDay.ToString('g') + "'")
        $local.intervals = @(for ($item=$range.GetFirst(); $null -ne $item; $item=$range.GetNext()) {
            if ($item.BusyStatus -ne 0) { @{ start=$item.StartUTC; end=$item.EndUTC } }
        })
        $local.known = $true
    } catch { Write-Verbose 'Personal calendar unavailable.' }
    $seen = @{}
    $organizerId = $session.CurrentUser.AddressEntry.ID
    if ($organizerId) { $seen[$organizerId] = $true }
    $others = @(foreach ($address in $addresses) {
        $person = $null
        $resolved = $false
        try {
            $person = $session.CreateRecipient($address)
            $resolved = $person.Resolve()
            $id = if ($resolved) { $person.AddressEntry.ID }
            if ($id) {
                if ($seen.ContainsKey($id)) { continue }
                $seen[$id] = $true
            }
        } catch { $resolved = $false }
        $name = if ($person.Name) { $person.Name } else { $address }
        @{ person=$name; recipient=$person; resolved=$resolved; days=@{}; unavailable=$false }
    })
    return @{ local=$local; others=$others }
}

function Get-ExchangeDay($person,[datetime]$day) {
    $key = $day.ToString('yyyy-MM-dd')
    if ($person.days.ContainsKey($key)) { return $person.days[$key] }
    if (!$person.resolved -or $person.unavailable) { return '' }
    if ([TimeZoneInfo]::Local.GetUtcOffset($day) -ne [TimeZoneInfo]::Local.GetUtcOffset($day.AddDays(1))) {
        $person.days[$key] = ''
        return ''
    }
    try {
        try { $grid = [string]$person.recipient.FreeBusy($day,30,$true) }
        catch { $grid = '' }
        if (!$grid) { $grid = [string]$person.recipient.AddressEntry.GetFreeBusy($day,30,$true) }
    } catch { $grid = '' }
    if (!$grid) { $person.unavailable = $true; return '' }
    # Exchange normally returns one month. Re-anchor after a local clock change
    # instead of projecting 48 half-hours onto a 23- or 25-hour day.
    for ($offset=0; $offset -lt $grid.Length; $offset+=48) {
        $date = $day.AddDays($offset / 48)
        if ([TimeZoneInfo]::Local.GetUtcOffset($date) -ne [TimeZoneInfo]::Local.GetUtcOffset($date.AddDays(1))) { break }
        $person.days[$date.ToString('yyyy-MM-dd')] = $grid.Substring($offset,[math]::Min(48,$grid.Length-$offset))
    }
    return [string]$person.days[$key]
}

function Get-MeetingStates($calendars,[datetimeoffset]$start,[datetimeoffset]$end) {
    $local = $calendars.local
    $state = if ($local.known) { 'free' } else { 'unknown' }
    foreach ($interval in $local.intervals) {
        if ($interval.start -lt $end.UtcDateTime -and $interval.end -gt $start.UtcDateTime) { $state='occupied-or-tentative'; break }
    }
    @{ person=$local.person; availability=$state }
    foreach ($person in $calendars.others) {
        $state = 'free'
        $cursor = $start
        while ($cursor -lt $end) {
            $begin = $cursor.LocalDateTime
            $day = $begin.Date
            $midnight = [datetimeoffset]$day.AddDays(1)
            $slots = Get-ExchangeDay $person $day
            $first = [int][math]::Floor($begin.TimeOfDay.TotalMinutes / 30)
            $after = if ($end -ge $midnight) { 48 } else { [int][math]::Ceiling($end.LocalDateTime.TimeOfDay.TotalMinutes / 30) }
            if ($slots -and $after -gt $first) {
                $available = [math]::Max(0,[math]::Min($slots.Length,$after)-$first)
                $slice = if ($available) { $slots.Substring($first,$available) } else { '' }
                if ($slice -match '[123]') { $state='occupied-or-tentative'; break }
                if ($available -lt $after-$first -or $slice -notmatch '^0+$') { $state='unknown' }
            } else { $state='unknown' }
            if ($person.unavailable -or !$person.resolved) { break }
            $cursor = $midnight
        }
        $row = @{ person=$person.person; availability=$state }
        if ($state -eq 'unknown') {
            $row.reason = if (!$person.resolved) { 'unresolved-recipient' } elseif ($person.unavailable) { 'exchange-unavailable' } else { 'no-published-data' }
        }
        $row
    }
}

function Find-MeetingAlternatives($calendars,[datetimeoffset]$start,[datetimeoffset]$end) {
    $duration = $end - $start
    $confirmed = @()
    $provisional = @()
    $now = [datetimeoffset]::Now
    for ($offset=0; $offset -lt 7; $offset++) {
        $day = $start.LocalDateTime.Date.AddDays($offset)
        if ($day.DayOfWeek -in 'Saturday','Sunday') { continue }
        $candidate = [datetimeoffset]$day.AddHours(9)
        $close = [datetimeoffset]$day.AddHours(18)
        while ($candidate + $duration -le $close) {
            $finish = $candidate + $duration
            if ($candidate -gt $start -and $candidate -gt $now) {
                $states = @(Get-MeetingStates $calendars $candidate $finish)
                if ('occupied-or-tentative' -notin $states.availability) {
                    $unknown = @($states | Where-Object availability -eq 'unknown' | ForEach-Object { $_.person })
                    $slot = @{ start=$candidate.ToString('yyyy-MM-ddTHH:mm:sszzz'); end=$finish.ToString('yyyy-MM-ddTHH:mm:sszzz')
                        label=($candidate.ToString('dd/MM/yyyy HH:mm') + ' - ' + $finish.ToString('HH:mm')); unknown=$unknown }
                    if (!$unknown.Count) { $confirmed += $slot } elseif ($provisional.Count -lt 3) { $provisional += $slot }
                    if ($confirmed.Count -eq 3) { return $confirmed }
                }
            }
            $candidate = $candidate.AddMinutes(30)
        }
    }
    @($confirmed + $provisional) | Select-Object -First 3
}

function New-MeetingConflictChoice($calendars,$states,[datetimeoffset]$start,[datetimeoffset]$end,[string]$keepToken) {
    $busy = @($states | Where-Object availability -eq 'occupied-or-tentative' | ForEach-Object { $_.person })
    $suggestions = @(Find-MeetingAlternatives $calendars $start $end)
    $options = @(foreach ($slot in $suggestions) {
        $description = if ($slot.unknown.Count) { 'Disponibilité à confirmer pour ' + ($slot.unknown -join ', ') } else { 'Libre pour tous les participants.' }
        @{ label=$slot.label; description=$description }
    })
    $requested = $start.ToLocalTime().ToString('dd/MM/yyyy HH:mm') + ' - ' + $end.ToLocalTime().ToString('dd/MM/yyyy HH:mm')
    $keepLabel = 'Conserver le créneau demandé'
    $options += @{ label=$keepLabel; description=$requested }
    if (!$suggestions.Count) { $options += @{ label='Choisir un autre horaire'; description='Indiquer une autre date ou heure.' } }
    $question = @{ header='Horaire'; require_answer=$true
        question="Sur le créneau du $requested, Outlook indique une indisponibilité pour $($busy -join ', '). Quel horaire souhaitez-vous retenir ?"; options=$options }
    $originalStart = $start.ToString('yyyy-MM-ddTHH:mm:sszzz')
    $originalEnd = $end.ToString('yyyy-MM-ddTHH:mm:sszzz')
    return @{ status='availability-choice-required'; action='meeting'; saved=$false; opened=$false; start=$originalStart; end=$originalEnd
        people=$states; suggestions=$suggestions; keepLabel=$keepLabel; keepSlot=$keepToken; questions=@($question)
    }
}

function Find-PreparedMeeting($session,[string]$subject,[datetimeoffset]$start,[string]$key) {
    $day = $start.LocalDateTime.Date
    $items = $session.GetDefaultFolder(9).Items.Restrict("[Subject] = '" + $subject.Replace("'","''") +
        "' AND [Start] >= '" + $day.ToString('g') + "' AND [Start] < '" + $day.AddDays(7).ToString('g') + "'")
    for ($item=$items.GetFirst(); $null -ne $item; $item=$items.GetNext()) {
        $marker = $item.UserProperties.Find('WhiteCodeScheduleChoice')
        if ($marker -and $key -cin ($marker.Value -split ' ')) { return $item }
    }
}

function Read-MeetingChoice($question) {
    $pipe = [IO.Pipes.NamedPipeClientStream]::new('.', $env:TOOL_CHOICE_PIPE, [IO.Pipes.PipeDirection]::InOut)
    try {
        $pipe.Connect(3000)
        $utf8 = [Text.UTF8Encoding]::new($false)
        $writer = [IO.StreamWriter]::new($pipe,$utf8,1024,$true)
        $reader = [IO.StreamReader]::new($pipe,$utf8,$false,1024,$true)
        $writer.WriteLine((ConvertTo-Json $question -Depth 5 -Compress))
        $writer.Flush()
        $line = $reader.ReadLine()
        if ($line) { return ($line | ConvertFrom-Json) }
    } catch { Write-Verbose 'The decision panel is unavailable.' }
    finally { $pipe.Dispose() }
    return @{none=$true}
}
