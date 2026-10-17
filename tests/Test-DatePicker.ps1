param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot '..\DatePicker.accdb'),
    [string]$ScreenshotDirectory = '',
    [switch]$VerifyInstalled
)

$ErrorActionPreference = 'Stop'
$encoding = [Text.Encoding]::GetEncoding(936, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
$source = $encoding.GetString([IO.File]::ReadAllBytes((Join-Path $PSScriptRoot '..\Module_DatePicker.bas')))
if ($source.Length -lt 10000 -or $source.Contains([char]0xFFFD)) { throw 'Missing or corrupted date picker source' }
if ($source -match '(?<!\r)\n') { throw 'VBA source must use CRLF' }
$source = $source -replace '(?m)^Attribute VB_Name[^\r\n]*\r?\n', ''
$continuations = 0
$longest = 0
foreach ($line in ($source -split '\r?\n')) {
    if ($line -match ' _\s*$') {
        $continuations++
        $longest = [Math]::Max($longest, $continuations)
    } else {
        $continuations = 0
    }
}
if ($longest -ge 24) { throw 'Too many VBA line continuations' }

$testSource = @'
Option Compare Database
Option Explicit

Private m_checks As Long

Private Sub Check(ByVal passed As Boolean, ByVal message As String)
    If Not passed Then Err.Raise vbObjectError + 1905, "DatePickerTests", message
    m_checks = m_checks + 1
End Sub

Public Function OpenDateSizePreview(ByVal showTime As Boolean) As Boolean
    DoCmd.OpenForm "Demo"
    Forms("Demo").Controls("Text2").Value = DateSerial(2026, 10, 17) + TimeSerial(13, 24, 0)
    PickDateFor Forms("Demo").Controls("Text2"), showTime
    Forms("frmDatePicker").TimerInterval = 0
    OpenDateSizePreview = True
End Function

Public Function RunDatePickerTests() As String
    On Error GoTo Failed
    Dim mode As Variant, showTime As Boolean, ctl As Control, frm As Form
    Dim expectedDate As Date, originalDate As Date, savedError As String
    m_checks = 0
    originalDate = DateSerial(2026, 10, 17) + TimeSerial(13, 24, 0)
    For Each mode In Array(False, True, False, True)
        showTime = CBool(mode)
        Check OpenDateSizePreview(showTime), "dropdown opens"
        Set frm = Forms("frmDatePicker")
        Check frm.Controls("lstHour").Visible = showTime, "hour visibility"
        Check frm.Controls("lstMin").Visible = showTime, "minute visibility"
        Check frm.Controls("lblHour").Visible = showTime, "hour label visibility"
        Check frm.Controls("lblMin").Visible = showTime, "minute label visibility"
        If showTime Then
            Check frm.InsideWidth > 3450, "date-time window retains full width"
            Check frm.Controls("lstHour").ListCount = 24, "hours retained"
            Check frm.Controls("lstMin").ListCount = 60, "minutes retained"
            Check frm.Controls("lstHour").Value = "13", "hour restored"
            Check frm.Controls("lstMin").Value = "24", "minute restored"
        Else
            Check Abs(frm.InsideWidth - 3450) <= 15, "actual date window narrows"
            Check frm.Controls("txtDisplay").Width = 3270, "date display narrows"
            Check frm.Controls("rctBorder").Width = 3400, "date border narrows"
        End If
        For Each ctl In frm.Controls
            If ctl.Visible Then
                Check ctl.Left + ctl.Width <= frm.InsideWidth + 15, "visible control fits width: " & ctl.Name
                Check ctl.Top + ctl.Height <= frm.InsideHeight + 15, "visible control fits height: " & ctl.Name
            End If
        Next ctl
        Set ctl = Nothing
        Set frm = Nothing
        Check DatePicker_Cancel(), "cancel succeeds"
        Check Forms("Demo").Controls("Text2").Value = originalDate, "cancel keeps input"
        Check OpenDateSizePreview(showTime), "reopen for writeback"
        Check DatePicker_DayClick(21), "day selection succeeds"
        expectedDate = DateSerial(2026, 10, 18)
        If showTime Then expectedDate = expectedDate + TimeSerial(13, 24, 0)
        Check Forms("Demo").Controls("Text2").Value = expectedDate, "date/time writeback retained"
        Check Not CurrentProject.AllForms("frmDatePicker").IsLoaded, "selection closes dropdown"
    Next mode
    Check OpenDateSizePreview(False), "open for clear"
    Check DatePicker_Clear(), "clear succeeds"
    Check IsNull(Forms("Demo").Controls("Text2").Value), "clear writes Null"
    Forms("Demo").Controls("Text2").Value = originalDate
    PickDateFor Forms("Demo").Controls("Text2"), False, "yyyy-mm-dd"
    Check DatePicker_DayClick(21), "formatted selection succeeds"
    Check Forms("Demo").Controls("Text2").Value = "2026-10-18", "formatted date writeback"
    RunDatePickerTests = "PASS: " & m_checks & " date picker assertions"
    Exit Function
Failed:
    savedError = "FAIL: " & Err.Number & " " & Err.Description
    On Error Resume Next
    DatePicker_Cancel
    RunDatePickerTests = savedError
End Function
'@

$captureTokens = $null
$captureErrors = $null
$captureAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Test-DateTimeRangePicker.ps1'), [ref]$captureTokens, [ref]$captureErrors)
if ($captureErrors.Count -gt 0) { throw 'Cannot parse existing native capture helper' }
$captureLiteral = $captureAst.Find({ param($node) $node -is [Management.Automation.Language.StringConstantExpressionAst] -and $node.Value -like '*public static class DateTimeRangeWindowCapture*' }, $true)
if ($null -eq $captureLiteral) { throw 'Missing native capture helper' }
if (-not ('DateTimeRangeWindowCapture' -as [type])) { Add-Type -TypeDefinition $captureLiteral.Value }
if ($ScreenshotDirectory) {
    if (-not (Test-Path -LiteralPath $ScreenshotDirectory -PathType Container)) { throw 'Screenshot directory does not exist' }
    Add-Type -AssemblyName System.Drawing
}

$temporaryDatabase = Join-Path $env:TEMP ('DatePickerTests-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
Copy-Item -LiteralPath $DatabasePath -Destination $temporaryDatabase
$access = $null
$project = $null
$ownedProcess = $null
$existingAccessIds = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    $windowHandle = [long]$access.hWndAccessApp()
    $ownedProcess = Get-Process -Name MSACCESS | Where-Object { $_.MainWindowHandle.ToInt64() -eq $windowHandle } | Select-Object -First 1
    if ($null -eq $ownedProcess -or $ownedProcess.Id -in $existingAccessIds) { throw 'Cannot identify a new, test-owned Access process' }
    $access.OpenCurrentDatabase($temporaryDatabase)
    $project = $access.VBE.ActiveVBProject
    $entries = @(@('Module_DatePickerTests', $testSource), @('Module_DatePicker', $source))
    foreach ($entry in $entries) {
        if ($VerifyInstalled -and $entry[0] -eq 'Module_DatePicker') { continue }
        Write-Output ('Importing ' + $entry[0])
        $component = $null
        foreach ($candidate in $project.VBComponents) {
            if ($candidate.Name -eq $entry[0]) { $component = $candidate; break }
        }
        if ($null -eq $component) { $component = $project.VBComponents.Add(1); $component.Name = $entry[0] }
        $codeModule = $component.CodeModule
        if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
        $codeModule.AddFromString($entry[1])
        $access.DoCmd.Save(5, $entry[0])
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($codeModule)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($component)
        $candidate = $null
    }
    Write-Output 'Compiling date picker test project'
    $access.DoCmd.RunCommand(126)
    $result = [string]$access.Eval('RunDatePickerTests()')
    Write-Output $result
    if ($result -notlike 'PASS:*') { throw $result }
    $pixelWidths = @()
    foreach ($mode in @('False', 'True', 'False', 'True')) {
        if (-not $access.Eval("OpenDateSizePreview($mode)")) { throw 'Cannot open size preview' }
        $pickerHandle = [IntPtr][long]$access.Eval('Forms("frmDatePicker").hWnd')
        $bounds = New-Object DateTimeRangeWindowCapture+Rect
        if ([DateTimeRangeWindowCapture]::GetWindowRect($pickerHandle, [ref]$bounds) -eq 0) { throw 'Cannot inspect popup window' }
        $pixelWidths += $bounds.Right - $bounds.Left
        if ($ScreenshotDirectory) {
            $fileName = 'access-date-only.png'
            if ($mode -eq 'True') { $fileName = 'access-date-time.png' }
            $bitmap = New-Object Drawing.Bitmap(($bounds.Right - $bounds.Left), ($bounds.Bottom - $bounds.Top))
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            try {
                $deviceContext = $graphics.GetHdc()
                try {
                    if ([DateTimeRangeWindowCapture]::PrintWindow($pickerHandle, $deviceContext, 2) -eq 0) { throw 'Cannot capture popup' }
                } finally {
                    $graphics.ReleaseHdc($deviceContext)
                }
                $bitmap.Save((Join-Path $ScreenshotDirectory $fileName), [Drawing.Imaging.ImageFormat]::Png)
            } finally {
                $graphics.Dispose()
                $bitmap.Dispose()
            }
        }
        [void]$access.Eval('DatePicker_Cancel()')
    }
    if ($pixelWidths[0] -ge $pixelWidths[1] -or $pixelWidths[0] -ne $pixelWidths[2] -or $pixelWidths[1] -ne $pixelWidths[3]) { throw 'Actual window sizes do not follow date/time mode' }
    Write-Output ("Native window widths: date=$($pixelWidths[0])px; date-time=$($pixelWidths[1])px; repeated mode switch: PASS")
    Write-Output "VBA compile: PASS; longest continuation run: $longest"
} finally {
    if ($null -ne $project) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($project) }
    if ($null -ne $access) {
        $access.CloseCurrentDatabase()
        if ($null -ne $ownedProcess -and $ownedProcess.Id -notin $existingAccessIds) { Stop-Process -InputObject $ownedProcess -ErrorAction SilentlyContinue }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($access)
    }
}