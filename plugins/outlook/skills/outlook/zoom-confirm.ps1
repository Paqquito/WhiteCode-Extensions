# The Outlook inspector/ribbon worker passes only its own Zoom scheduling dialog.
function Confirm-ZoomDialog($dialog) {
    $handle = [IntPtr]$dialog.Current.NativeWindowHandle
    if ($dialog.Current.ClassName -ne 'DUIMainFrame' -or $handle -eq [IntPtr]::Zero) { return $false }
    $buttons = @($dialog.FindAll($tree,$buttonType) | Where-Object {
        $_.Current.IsEnabled -and !$_.Current.IsOffscreen -and $_.Current.Name -in 'Continuez','Continuer','Continue'
    })
    if ($buttons.Count -eq 1) {
        $buttons[0].GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
        return $true
    }
    # Older Zoom clients draw the controls themselves. Read only the owned dialog,
    # in memory, and require a unique Continue/Cancel button pair on the same row.
    Add-Type -AssemblyName System.Runtime.WindowsRuntime,System.Drawing
    Add-Type -Path (Join-Path $PSScriptRoot 'ZoomDialog.cs') -ReferencedAssemblies System.Drawing
    $null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrResult,Windows.Foundation,ContentType=WindowsRuntime]
    $asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetGenericArguments().Count -eq 1 -and
        $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation' + [char]96 + '1'
    } | Select-Object -First 1
    function Read-Operation($op,$type) {
        $task = $asTask.MakeGenericMethod($type).Invoke($null,@($op))
        if (!$task.Wait(3000)) { throw 'Zoom dialog recognition timed out.' }
        $task.Result
    }
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if (!$engine) { return $false }
    $stream = [ZoomDialog]::Capture($handle)
    try {
        $source = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($stream)
        $decoder = Read-Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($source)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Read-Operation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        try {
            $page = Read-Operation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            $words = @($page.Lines | ForEach-Object { $_.Words } | Where-Object { $_.BoundingRect.Y -gt $bitmap.PixelHeight * 0.75 })
            $ok = @($words | Where-Object { $_.Text -in 'Continuez','Continuer','Continue' })
            $cancel = @($words | Where-Object { $_.Text -in 'Annuler','Cancel' })
            if ($ok.Count -ne 1 -or $cancel.Count -ne 1) { return $false }
            $a=$ok[0].BoundingRect; $b=$cancel[0].BoundingRect
            if ($a.X + $a.Width -ge $b.X -or [math]::Abs($a.Y-$b.Y) -gt $a.Height) { return $false }
            return [ZoomDialog]::ClickClient($handle,$bitmap.PixelWidth,$bitmap.PixelHeight,
                [int]($a.X+$a.Width/2),[int]($a.Y+$a.Height/2))
        } finally { $bitmap.Dispose() }
    } finally { $stream.Dispose() }
}
