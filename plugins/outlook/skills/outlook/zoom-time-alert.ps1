# Handle just the scheduling error belonging to this meeting's inspector.
function Close-ZoomTimeAlert([long]$handle) {
    Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
    $element = [Windows.Automation.AutomationElement]
    $root = $element::FromHandle([intptr]$handle)
    $tree = [Windows.Automation.TreeScope]::Descendants
    $dialogs = New-Object Windows.Automation.PropertyCondition($element::ClassNameProperty,'#32770')
    foreach ($dialog in $root.FindAll($tree,$dialogs)) {
        if ($dialog.Current.ProcessId -ne $root.Current.ProcessId -or
            $dialog.Current.Name -notmatch '^Zoom\s*[-\u2013]\s*(Planifier une r.union|Schedule a Meeting)$') { continue }
        $controls = $dialog.FindAll($tree,[Windows.Automation.Condition]::TrueCondition)
        $message = @($controls | Where-Object { $_.Current.ClassName -eq 'Static' -and $_.Current.Name -match
            '(?i)(heure de fin soit post.rieure . l.heure actuelle|end time (is |will be |to be )?(later than|after) the current time)' })
        $buttons = @($controls | Where-Object { $_.Current.ClassName -eq 'Button' })
        if ($message.Count -ne 1 -or $buttons.Count -ne 1 -or $buttons[0].Current.Name -ne 'OK') { continue }
        return [MeetingWindow]::AcknowledgeButton([intptr]$dialog.Current.NativeWindowHandle,[intptr]$buttons[0].Current.NativeWindowHandle)
    }
    return $false
}
