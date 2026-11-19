# Windows PowerShell worker. The caller supplies the HWND of the newly created Outlook inspector.
param([Parameter(Mandatory)][long]$Window, [Parameter(Mandatory)][ValidateSet('Zoom','Tixeo')][string]$Provider)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
$element = [Windows.Automation.AutomationElement]
$tree = [Windows.Automation.TreeScope]::Descendants
$root = $element::FromHandle([IntPtr]$Window)
if ((Get-Process -Id $root.Current.ProcessId).ProcessName -ne 'OUTLOOK') { exit 2 }
$groupName = New-Object Windows.Automation.PropertyCondition($element::NameProperty,$Provider)
$buttonType = New-Object Windows.Automation.PropertyCondition($element::ControlTypeProperty,[Windows.Automation.ControlType]::Button)
$until = [datetime]::UtcNow.AddSeconds(5)
do {
    $buttons = @(
        foreach ($group in $root.FindAll($tree,$groupName)) {
            if ($group.Current.ControlType -ne [Windows.Automation.ControlType]::Group) { continue }
            foreach ($button in $group.FindAll($tree,$buttonType)) {
                if ($button.Current.IsOffscreen) { continue }
                # Accent-independent for Windows PowerShell 5.1 installations reading UTF-8 without BOM.
                $name = $button.Current.Name.Normalize([Text.NormalizationForm]::FormD) -replace '\p{M}',''
                if (($Provider -eq 'Zoom' -and $name -in 'Programmer une reunion','Schedule a Meeting') -or
                    ($Provider -eq 'Tixeo' -and $name -in 'Reunion en ligne','Online Meeting')) { $button }
            }
        }
    )
    if ($buttons.Count -eq 1) {
        if (!$buttons[0].Current.IsEnabled) {
            Start-Sleep -Milliseconds 150
            continue
        }
        if ($Provider -eq 'Zoom') {
            $buttons[0].GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
            $dialogs = New-Object Windows.Automation.PropertyCondition($element::ControlTypeProperty,[Windows.Automation.ControlType]::Window)
            for ($attempt=0; $attempt -lt 12; $attempt++) {
                foreach ($dialog in $root.FindAll($tree,$dialogs)) {
                    if ($dialog.Current.Name -match '^Zoom\s*[-\u2013]\s*(Planifier une r.union|Schedule a Meeting)$') { exit 4 }
                }
                Start-Sleep -Milliseconds 150
            }
        } else {
            $toggle = $buttons[0].GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
            if ($toggle.Current.ToggleState -eq [Windows.Automation.ToggleState]::Off) { $toggle.Toggle() }
            if ($toggle.Current.ToggleState -ne [Windows.Automation.ToggleState]::On) { exit 3 }
        }
        exit 0
    }
    if ($buttons.Count -gt 1) { exit 2 }
    Start-Sleep -Milliseconds 150
} while ([datetime]::UtcNow -lt $until)
if ($buttons.Count -eq 1 -and !$buttons[0].Current.IsEnabled) { exit 5 }
exit 2
