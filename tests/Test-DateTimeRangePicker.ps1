param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot '..\DatePicker.accdb'),
    [string]$ScreenshotPath = '',
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot '..\Module_DateTimeRangePicker.bas'
$sourceBytes = [IO.File]::ReadAllBytes($modulePath)
try {
    $sourceText = (New-Object Text.UTF8Encoding($false, $true)).GetString($sourceBytes)
} catch [Text.DecoderFallbackException] {
    $sourceText = [Text.Encoding]::GetEncoding(936).GetString($sourceBytes)
}
if ([string]::IsNullOrWhiteSpace($sourceText) -or $sourceText -notmatch '(?m)^Public Function CreateDateTimeRangePickerForm') { throw 'Missing picker source; refusing to compile an empty module' }
if ($sourceText.Contains([char]0xFFFD)) { throw 'Corrupted source text' }
$continuations = 0
$longest = 0
foreach ($line in ($sourceText -split '\r?\n')) {
    if ($line -match ' _\s*$') {
        $continuations++
        $longest = [Math]::Max($longest, $continuations)
    } else {
        $continuations = 0
    }
}
if ($longest -ge 24) { throw "Too many VBA line continuations: $longest" }

$testSource = @'
Option Compare Database
Option Explicit

Private Const PICKER As String = "frmDateTimeRangePicker"
Private Const HOST As String = "frmDateTimeRangeTestHost"
Private Const INITIAL As String = "2026-10-07 10:10:36 ~ 2026-11-18 10:11:37"
Private m_checks As Long
Private m_modalAction As Integer

Private Sub Check(ByVal passed As Boolean, ByVal message As String)
    If Not passed Then
        Err.Raise vbObjectError + 1903, "DateTimeRangeTests", message
    End If
    m_checks = m_checks + 1
End Sub

Private Sub MakeHost()
    Dim frm As Form, ctl As Control, temporaryName As String
    Set frm = CreateForm
    temporaryName = frm.Name
    frm.Width = 11000
    frm.Section(acDetail).Height = 1800
    frm.PopUp = True
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 180, 180, 10300, 420)
    ctl.Name = "txtRange"
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 180, 900, 4800, 420)
    ctl.Name = "txtStart"
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 5400, 900, 4800, 420)
    ctl.Name = "txtEnd"
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename HOST, acForm, temporaryName
    DoCmd.OpenForm HOST
End Sub

Private Sub OpenRange(ByVal value As Variant)
    Forms(HOST).Controls("txtRange").Value = value
    Check PickDateTimeRangeFor(Forms(HOST).Controls("txtRange")), "dropdown opens"
End Sub

Private Function FindCell(ByVal panel As Integer, ByVal dayNumber As Integer) As Integer
    Dim cell As Integer, ctl As Control
    For cell = panel * 42 + 1 To panel * 42 + 42
        Set ctl = Forms(PICKER).Controls("lblD" & Format$(cell, "00"))
        If ctl.Caption = CStr(dayNumber) And ctl.ForeColor <> 10526880 Then
            FindCell = cell
            Exit Function
        End If
    Next cell
    Err.Raise 5, , "Cannot find day"
End Function

Private Sub SetTime(ByVal endpoint As Integer, ByVal hourNumber As Integer, ByVal minuteNumber As Integer, ByVal secondNumber As Integer)
    Dim prefix As String
    prefix = "cboStart"
    If endpoint = 1 Then prefix = "cboEnd"
    Forms(PICKER).Controls(prefix & "Hour").Value = Format$(hourNumber, "00")
    Forms(PICKER).Controls(prefix & "Min").Value = Format$(minuteNumber, "00")
    Forms(PICKER).Controls(prefix & "Sec").Value = Format$(secondNumber, "00")
    Check DateTimeRangePicker_TimeChanged(endpoint), "time updated"
End Sub

Public Function RangeModalTimer() As Boolean
    On Error GoTo Failed
    Forms(PICKER).TimerInterval = 0
    If m_modalAction = 1 Then
        RangeModalTimer = DateTimeRangePicker_Cancel()
    ElseIf m_modalAction = 2 Then
        RangeModalTimer = DateTimeRangePicker_Clear()
    Else
        RangeModalTimer = DateTimeRangePicker_OK()
    End If
    Exit Function
Failed:
    DateTimeRangePicker_Cancel
End Function

Public Function OpenRangePreview() As Boolean
    DoCmd.OpenForm HOST
    Forms(HOST).Controls("txtRange").Value = INITIAL
    OpenRangePreview = PickDateTimeRangeForCtl(HOST, "txtRange")
End Function

Public Function RunDateTimeRangeTests() As String
    On Error GoTo Failed
    Dim ctl As Control, endpoint As Integer, prefix As String, item As Integer, chosen As Variant
    Check CreateDateTimeRangePickerForm(True), "form build"
    Check CreateDateTimeRangePickerForm(True), "idempotent rebuild"
    DoCmd.OpenForm PICKER
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "empty range cannot confirm"
    Check Forms(PICKER).Controls("cmdOK").Default, "Enter confirms"
    Check Forms(PICKER).Controls("cmdCancel").Cancel, "Escape cancels"
    Check Forms(PICKER).Caption = ChrW(&H65F6) & ChrW(&H95F4) & ChrW(&H8303) & ChrW(&H56F4), "Chinese caption intact"
    For Each ctl In Forms(PICKER).Controls
        Check ctl.Left >= 0 And ctl.Left + ctl.Width <= Forms(PICKER).Width, "horizontal bounds: " & ctl.Name
        Check ctl.Top >= 0 And ctl.Top + ctl.Height <= Forms(PICKER).Section(acDetail).Height, "vertical bounds: " & ctl.Name
    Next ctl
    For endpoint = 0 To 1
        prefix = "cboStart"
        If endpoint = 1 Then prefix = "cboEnd"
        Check Forms(PICKER).Controls(prefix & "Hour").ListCount = 24, "24 hours"
        Check Forms(PICKER).Controls(prefix & "Min").ListCount = 60, "60 minutes"
        Check Forms(PICKER).Controls(prefix & "Sec").ListCount = 60, "60 seconds"
        Check Forms(PICKER).Controls(prefix & "Sec").AfterUpdate = "=DateTimeRangePicker_TimeChanged(" & endpoint & ")", "time binding"
    Next endpoint
    DateTimeRangePicker_Cancel
    MakeHost
    OpenRange INITIAL
    Check Forms(PICKER).Controls("lblDisplay").Caption = INITIAL, "default dates and seconds restored"
    DateTimeRangePicker_DayClick FindCell(0, 7)
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "partial range cannot confirm"
    DateTimeRangePicker_DayMouseMove FindCell(1, 18)
    Check Forms(PICKER).Controls("lblD" & Format$(FindCell(1, 10), "00")).BackColor = 16576233, "hover range"
    DateTimeRangePicker_DayClick FindCell(1, 18)
    Check CurrentProject.AllForms(PICKER).IsLoaded, "second date click does not close"
    Check Forms(HOST).Controls("txtRange").Value = INITIAL, "dates only preview"
    For item = 0 To 59
        SetTime 0, 10, 10, item
        Check Forms(PICKER).Controls("cboStartSec").Value = Format$(item, "00"), "each start second"
        SetTime 1, 10, 11, item
        Check Forms(PICKER).Controls("cboEndSec").Value = Format$(item, "00"), "each end second"
    Next item
    SetTime 0, 0, 0, 0
    SetTime 1, 23, 59, 59
    Check Forms(HOST).Controls("txtRange").Value = INITIAL, "time edits only preview"
    DateTimeRangePicker_OK
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-07 00:00:00 ~ 2026-11-18 23:59:59", "combined seconds output"

    OpenRange INITIAL
    SetTime 0, 22, 33, 44
    DateTimeRangePicker_Cancel
    Check Forms(HOST).Controls("txtRange").Value = INITIAL, "cancel unchanged"
    OpenRange INITIAL
    Forms(PICKER).Controls("cboStartHour").Value = "99"
    Check Not DateTimeRangePicker_TimeChanged(0), "invalid hour rejected"
    Check Forms(PICKER).Controls("cboStartHour").Value = "10", "invalid hour restored"
    Forms(PICKER).Controls("cboEndSec").Value = Null
    Check Not DateTimeRangePicker_TimeChanged(1), "empty seconds rejected"
    Check Forms(PICKER).Controls("cboEndSec").Value = "37", "empty seconds restored"
    Forms(PICKER).Controls("txtStartDate").Value = "not a date"
    Check Not DateTimeRangePicker_DateChanged(0), "invalid date rejected"
    Check Forms(PICKER).Controls("txtStartDate").Value = #10/7/2026#, "invalid date restored"
    DateTimeRangePicker_Cancel
    OpenRange INITIAL
    DateTimeRangePicker_Clear
    Check IsNull(Forms(HOST).Controls("txtRange").Value), "clear writes Null"
    OpenRange Null
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "Null starts unselected"
    Check Forms(PICKER).Controls("cboStartHour").Value = "00", "default start midnight"
    Check Forms(PICKER).Controls("cboEndSec").Value = "59", "default end day"
    DateTimeRangePicker_Cancel
    OpenRange "not a date range"
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "invalid input fallback"
    DateTimeRangePicker_Cancel

    OpenRange INITIAL
    Forms(PICKER).Controls("txtStartDate").Value = #12/31/2026#
    Check DateTimeRangePicker_DateChanged(0), "manual start date"
    Forms(PICKER).Controls("txtEndDate").Value = #1/1/2027#
    Check DateTimeRangePicker_DateChanged(1), "manual end date"
    DateTimeRangePicker_OK
    Check Forms(HOST).Controls("txtRange").Value = "2026-12-31 10:10:36 ~ 2027-01-01 10:11:37", "cross year retains time"

    OpenRange "2028-02-29 23:59:58 ~ 2028-03-01 00:00:01"
    DateTimeRangePicker_OK
    Check Forms(HOST).Controls("txtRange").Value = "2028-02-29 23:59:58 ~ 2028-03-01 00:00:01", "leap day"
    OpenRange "2026-10-17 23:59:59 ~ 2026-10-17 00:00:00"
    DateTimeRangePicker_OK
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-17 00:00:00 ~ 2026-10-17 23:59:59", "same-day reversal sorted"
    OpenRange "2026-10-17 12:25:30 ~ 2026-10-17 12:25:31"
    SetTime 0, 12, 25, 32
    DateTimeRangePicker_OK
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-17 12:25:31 ~ 2026-10-17 12:25:32", "edited same-day reversal sorted at confirmation"
    OpenRange INITIAL
    DateTimeRangePicker_DayClick FindCell(1, 18)
    DateTimeRangePicker_DayClick FindCell(0, 7)
    DateTimeRangePicker_OK
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-07 10:11:37 ~ 2026-11-18 10:10:36", "reversed endpoints move with time"

    Forms(HOST).Controls("txtStart").Value = #10/17/2026 12:25:30#
    Forms(HOST).Controls("txtEnd").Value = #10/17/2026 12:25:31#
    Check PickDateTimeRangeFor(Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd")), "two controls open"
    SetTime 1, 12, 25, 32
    DateTimeRangePicker_OK
    Check VarType(Forms(HOST).Controls("txtStart").Value) = vbDate, "start is Date"
    Check VarType(Forms(HOST).Controls("txtEnd").Value) = vbDate, "end is Date"
    Check Forms(HOST).Controls("txtEnd").Value = #10/17/2026 12:25:32#, "end second retained"
    Check PickDateTimeRangeFor(Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd"), "yyyy/mm/dd hh:nn:ss"), "custom format opens"
    DateTimeRangePicker_OK
    Check VarType(Forms(HOST).Controls("txtStart").Value) = vbString, "custom format is String"
    Check Forms(HOST).Controls("txtEnd").Value = "2026/10/17 12:25:32", "custom format seconds"
    Check PickDateTimeRangeFor(Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd")), "clear two controls opens"
    DateTimeRangePicker_Clear
    Check IsNull(Forms(HOST).Controls("txtStart").Value) And IsNull(Forms(HOST).Controls("txtEnd").Value), "clear both controls"

    Forms(HOST).Controls("txtStart").Value = DateAdd("s", 3661, DateSerial(100, 1, 1))
    Forms(HOST).Controls("txtEnd").Value = DateAdd("s", 3662, DateSerial(100, 1, 1))
    Check PickDateTimeRangeFor(Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd")), "minimum year opens"
    Check Not DateTimeRangePicker_Navigate(-1), "minimum navigation clamped"
    DateTimeRangePicker_OK
    Check Year(Forms(HOST).Controls("txtStart").Value) = 100 And Hour(Forms(HOST).Controls("txtStart").Value) = 1, "negative OLE Date retains date and hour"
    Check Minute(Forms(HOST).Controls("txtStart").Value) = 1 And Second(Forms(HOST).Controls("txtStart").Value) = 1, "negative OLE Date retains seconds: actual=" & Format$(Forms(HOST).Controls("txtStart").Value, "yyyy-mm-dd hh:nn:ss") & "; source=" & Format$(DateAdd("s", 3661, DateSerial(100, 1, 1)), "yyyy-mm-dd hh:nn:ss")
    Forms(HOST).Controls("txtStart").Value = DateSerial(9999, 12, 31)
    Forms(HOST).Controls("txtEnd").Value = DateAdd("s", 86399, DateSerial(9999, 12, 31))
    Check PickDateTimeRangeFor(Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd")), "maximum date opens"
    Check Not DateTimeRangePicker_Navigate(1), "maximum navigation clamped"
    DateTimeRangePicker_OK
    Check Forms(HOST).Controls("txtEnd").Value = DateAdd("s", 86399, DateSerial(9999, 12, 31)), "maximum second retained"

    AttachDateTimeRangePicker Forms(HOST), "txtStart", "txtEnd"
    Check Forms(HOST).Controls("txtStart").OnDblClick = Forms(HOST).Controls("txtEnd").OnDblClick, "attach both events"
    Forms(HOST).Controls("txtRange").Value = INITIAL
    Check PickDateTimeRangeForCtl(HOST, "txtRange"), "expression entry"
    DoCmd.SelectObject acForm, HOST
    DateTimeRangePicker_Deactivate
    DateTimeRangePicker_Timer
    Check Not CurrentProject.AllForms(PICKER).IsLoaded, "outside click closes"
    Check Forms(HOST).Controls("txtRange").Value = INITIAL, "outside click unchanged"

    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=RangeModalTimer()"
    Forms(PICKER).TimerInterval = 100
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowDateTimeRangePicker(#10/7/2026 10:10:36#, #11/18/2026 10:11:37#)
    Check IsArray(chosen), "modal result array"
    Check VarType(chosen(0)) = vbDate And VarType(chosen(1)) = vbDate, "modal result Date values"
    Check chosen(0) = #10/7/2026 10:10:36# And chosen(1) = #11/18/2026 10:11:37#, "modal seconds"
    m_modalAction = 1
    chosen = ShowDateTimeRangePicker()
    Check IsNull(chosen), "modal no defaults and cancel"
    m_modalAction = 2
    chosen = ShowDateTimeRangePicker(#10/17/2026#, #10/18/2026#)
    Check IsNull(chosen), "modal clear returns Null"
    Check CreateDateTimeRangePickerForm(True), "restore production events"
    DoCmd.Close acForm, HOST, acSaveNo
    RunDateTimeRangeTests = "PASS: " & m_checks & " assertions"
    Exit Function
Failed:
    Dim savedError As String
    savedError = "FAIL: " & Err.Number & " " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, PICKER, acSaveNo
    RunDateTimeRangeTests = savedError
End Function
'@

$temporaryDatabase = Join-Path $env:TEMP ('DateTimeRangeTests-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
Copy-Item -LiteralPath $DatabasePath -Destination $temporaryDatabase
$access = $null
$project = $null
$ownedProcess = $null
$existingAccessIds = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$previewKeptOpen = $false
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    $windowHandle = [long]$access.hWndAccessApp()
    $ownedProcess = Get-Process -Name MSACCESS | Where-Object { $_.MainWindowHandle.ToInt64() -eq $windowHandle } | Select-Object -First 1
    if ($null -eq $ownedProcess -or $ownedProcess.Id -in $existingAccessIds) { throw 'Cannot identify a new, test-owned Access process' }
    $access.OpenCurrentDatabase($temporaryDatabase)
    Write-Output 'Temporary database opened'
    $project = $access.VBE.ActiveVBProject
    foreach ($entry in @(@('Module_DateTimeRangePicker', $sourceText), @('Module_DateTimeRangeTests', $testSource))) {
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
    $access.DoCmd.RunCommand(126)
    Write-Output 'VBA compilation completed; running range assertions'
    $result = $access.Eval('RunDateTimeRangeTests()')
    Write-Output $result
    if ($result -notlike 'PASS:*') { throw $result }
    Write-Output "VBA compile: PASS; longest continuation run: $longest"
    if ($KeepOpen -or $ScreenshotPath) {
        [void]$access.Eval('OpenRangePreview()')
        if ($ScreenshotPath) {
            Add-Type -AssemblyName System.Drawing
            if (-not ('DateTimeRangeWindowCapture' -as [type])) {
                Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DateTimeRangeWindowCapture {
    [StructLayout(LayoutKind.Sequential)]
    public struct Rect { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll", ExactSpelling=true, SetLastError=true, CallingConvention=CallingConvention.Winapi)]
    public static extern int GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll", ExactSpelling=true, CallingConvention=CallingConvention.Winapi)]
    public static extern int PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);
}
'@
            }
            $pickerHandle = [IntPtr][long]$access.Eval('[Forms]![frmDateTimeRangePicker].[hWnd]')
            $bounds = New-Object DateTimeRangeWindowCapture+Rect
            if ([DateTimeRangeWindowCapture]::GetWindowRect($pickerHandle, [ref]$bounds) -eq 0) { throw 'Cannot read picker bounds' }
            $bitmap = New-Object Drawing.Bitmap(($bounds.Right - $bounds.Left), ($bounds.Bottom - $bounds.Top))
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            try {
                $deviceContext = $graphics.GetHdc()
                try {
                    if ([DateTimeRangeWindowCapture]::PrintWindow($pickerHandle, $deviceContext, 2) -eq 0) { throw 'Cannot capture picker window' }
                } finally {
                    $graphics.ReleaseHdc($deviceContext)
                }
                $bitmap.Save($ScreenshotPath, [Drawing.Imaging.ImageFormat]::Png)
            } finally {
                $graphics.Dispose()
                $bitmap.Dispose()
            }
            Write-Output "Screenshot: $ScreenshotPath"
        }
        if ($KeepOpen) {
            $access.UserControl = $true
            $previewKeptOpen = $true
            Write-Output "Preview database: $temporaryDatabase"
        }
    }
} finally {
    if ($null -ne $project) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($project) }
    if ($null -ne $access) {
        if (-not $previewKeptOpen) {
            Write-Output 'Closing temporary database'
            if ($ScreenshotPath) { [void]$access.Eval('DateTimeRangePicker_Cancel()') }
            $access.CloseCurrentDatabase()
            Write-Output 'Temporary database closed'
            if ($null -ne $ownedProcess -and $ownedProcess.Id -notin $existingAccessIds) {
                Stop-Process -InputObject $ownedProcess -ErrorAction SilentlyContinue
            }
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($access)
    }
    if (-not $previewKeptOpen) { Write-Output 'Temporary Access instance cleaned up' }
}