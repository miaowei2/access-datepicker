param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot '..\DatePicker.accdb'),
    [string]$ScreenshotPath = ''
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot '..\Module_TimePicker.bas'
$sourceBytes = [IO.File]::ReadAllBytes($modulePath)
try {
    $sourceText = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($sourceBytes)
} catch [System.Text.DecoderFallbackException] {
    $sourceText = [Text.Encoding]::GetEncoding(936).GetString($sourceBytes)
}
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

Private Const PICKER As String = "frmTimePicker"
Private Const HOST As String = "frmTimeTestHost"
Private m_checks As Long

Private Sub Check(ByVal passed As Boolean, ByVal message As String)
    If Not passed Then
        Err.Raise vbObjectError + 1902, "TimePickerTests", message
    End If
    m_checks = m_checks + 1
End Sub

Private Sub MakeHost()
    Dim frm As Form, ctl As Control, temporaryName As String
    Set frm = CreateForm
    temporaryName = frm.Name
    frm.Width = 7000
    frm.Section(acDetail).Height = 1500
    frm.PopUp = True
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 180, 180, 6000, 420)
    ctl.Name = "txtTime"
    ctl.Format = "hh:nn:ss"
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename HOST, acForm, temporaryName
    DoCmd.OpenForm HOST
End Sub

Private Sub OpenTime(ByVal value As Variant, Optional ByVal outputFormat As String = "")
    Forms(HOST).Controls("txtTime").Value = value
    Check PickTimeFor(Forms(HOST).Controls("txtTime"), outputFormat), "dropdown opens"
End Sub

Private Sub SetTime(ByVal hourNumber As Integer, ByVal minuteNumber As Integer, ByVal secondNumber As Integer)
    Forms(PICKER).Controls("lstHour").Value = Format$(hourNumber, "00")
    Forms(PICKER).Controls("lstMin").Value = Format$(minuteNumber, "00")
    Forms(PICKER).Controls("lstSec").Value = Format$(secondNumber, "00")
    Check TimePicker_TimeChanged(), "list change handled"
End Sub

Public Function TimeTestModalTimer() As Boolean
    On Error GoTo Failed
    Forms(PICKER).TimerInterval = 0
    Check Forms(PICKER).Controls("lblDisplay").Caption = "12:25:30", "modal default retains seconds"
    SetTime 23, 59, 59
    TimeTestModalTimer = TimePicker_OK()
    Exit Function
Failed:
    TimePicker_Cancel
End Function

Public Function OpenTimeTestPreview() As Boolean
    DoCmd.OpenForm HOST
    Forms(HOST).Controls("txtTime").Value = TimeSerial(12, 25, 30)
    OpenTimeTestPreview = PickTimeForCtl(HOST, "txtTime")
    Forms(PICKER).Repaint
    DoEvents
End Function

Public Function RunTimeTests() As String
    On Error GoTo Failed
    Dim chosen As Variant, ctl As Control, item As Integer, oldValue As Variant
    Check CreateTimePickerForm(True), "form build"
    Check CreateTimePickerForm(True), "idempotent rebuild"
    DoCmd.OpenForm PICKER
    Check Forms(PICKER).Controls("lstHour").ListCount = 24, "24 hours"
    Check Forms(PICKER).Controls("lstMin").ListCount = 60, "60 minutes"
    Check Forms(PICKER).Controls("lstSec").ListCount = 60, "60 seconds"
    Check Not IsNull(Forms(PICKER).Controls("lstHour").Value), "direct open initializes hour"
    Check Forms(PICKER).Controls("lstHour").AfterUpdate = "=TimePicker_TimeChanged()", "hour event"
    Check Forms(PICKER).Controls("lstMin").AfterUpdate = "=TimePicker_TimeChanged()", "minute event"
    Check Forms(PICKER).Controls("lstSec").AfterUpdate = "=TimePicker_TimeChanged()", "second event"
    Check Forms(PICKER).Controls("cmdOK").Default, "Enter confirms"
    Check Forms(PICKER).Controls("cmdCancel").Cancel, "Escape cancels"
    For Each ctl In Forms(PICKER).Controls
        Check ctl.Left + ctl.Width <= Forms(PICKER).Width, "horizontal bounds: " & ctl.Name
        Check ctl.Top + ctl.Height <= Forms(PICKER).Section(acDetail).Height, "vertical bounds: " & ctl.Name
    Next ctl
    TimePicker_Cancel
    MakeHost

    OpenTime #10/17/2026 12:25:30#
    Check Forms(PICKER).Controls("lblDisplay").Caption = "12:25:30", "date portion ignored and seconds retained"
    Check Forms(PICKER).Controls("lstHour").Value = "12", "default hour selection"
    Check Forms(PICKER).Controls("lstMin").Value = "25", "default minute selection"
    Check Forms(PICKER).Controls("lstSec").Value = "30", "default second selection"
    For item = 0 To 23
        SetTime item, 25, 30
        Check Forms(PICKER).Controls("lblDisplay").Caption = Format$(item, "00") & ":25:30", "each hour can be selected"
    Next item
    For item = 0 To 59
        SetTime 12, item, 30
        Check Forms(PICKER).Controls("lblDisplay").Caption = "12:" & Format$(item, "00") & ":30", "each minute can be selected"
        SetTime 12, 25, item
        Check Forms(PICKER).Controls("lblDisplay").Caption = "12:25:" & Format$(item, "00"), "each second can be selected"
    Next item
    Check Forms(HOST).Controls("txtTime").Value = #10/17/2026 12:25:30#, "preview does not write"
    SetTime 23, 59, 59
    Check TimePicker_OK(), "confirm succeeds"
    Check Not CurrentProject.AllForms(PICKER).IsLoaded, "confirm closes"
    Check VarType(Forms(HOST).Controls("txtTime").Value) = vbDate, "default output is Date"
    Check Forms(HOST).Controls("txtTime").Value = TimeSerial(23, 59, 59), "maximum time without date"

    OpenTime #12:25:30 PM#
    SetTime 0, 0, 0
    TimePicker_OK
    Check Forms(HOST).Controls("txtTime").Value = TimeSerial(0, 0, 0), "midnight is valid"
    OpenTime TimeSerial(0, 0, 0)
    Check Forms(PICKER).Controls("lblDisplay").Caption = "00:00:00", "midnight reopening does not use current time"
    TimePicker_Cancel
    OpenTime "12:25:30", "hh:nn:ss"
    SetTime 1, 2, 3
    TimePicker_OK
    Check VarType(Forms(HOST).Controls("txtTime").Value) = vbString, "custom format writes text"
    Check Forms(HOST).Controls("txtTime").Value = "01:02:03", "custom format with leading zeroes"
    OpenTime "01:02:03"
    Check Forms(PICKER).Controls("lblDisplay").Caption = "01:02:03", "text default restored"
    SetTime 4, 5, 6
    TimePicker_Cancel
    Check Forms(HOST).Controls("txtTime").Value = "01:02:03", "cancel keeps input"
    OpenTime Null
    Check Not IsNull(Forms(PICKER).Controls("lstSec").Value), "Null defaults to current time"
    TimePicker_Cancel
    OpenTime "invalid"
    Check Forms(PICKER).Controls("lblDisplay").Caption Like "##:##:##", "invalid defaults to current time"
    TimePicker_Now
    Check Abs(DateDiff("s", CDate(Forms(PICKER).Controls("lblDisplay").Caption), Time)) <= 2, "now button sets current time"
    TimePicker_Clear
    Check IsNull(Forms(HOST).Controls("txtTime").Value), "clear writes Null"

    AttachTimePicker Forms(HOST), "txtTime", "hh:nn:ss"
    Check Forms(HOST).Controls("txtTime").OnDblClick = "=PickTimeForCtl(""" & HOST & """,""txtTime"",""hh:nn:ss"")", "attach expression"
    Check PickTimeForCtl(HOST, "txtTime"), "expression API opens"
    TimePicker_Cancel
    OpenTime "12:25:30"
    SetTime 4, 5, 6
    TimePicker_Timer
    Check CurrentProject.AllForms(PICKER).IsLoaded, "timer does not close active popup"
    Forms(HOST).SetFocus
    TimePicker_Timer
    Check Not CurrentProject.AllForms(PICKER).IsLoaded, "outside click closes"
    Check Forms(HOST).Controls("txtTime").Value = "12:25:30", "outside preserves old value"

    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=TimeTestModalTimer()"
    Forms(PICKER).TimerInterval = 50
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowTimePicker(#10/17/2026 12:25:30#)
    Check VarType(chosen) = vbDate, "modal returns Date"
    Check chosen = TimeSerial(23, 59, 59), "modal result without date"
    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=TimePicker_Cancel()"
    Forms(PICKER).TimerInterval = 50
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowTimePicker()
    Check IsNull(chosen), "modal cancel returns Null"
    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=TimePicker_Clear()"
    Forms(PICKER).TimerInterval = 50
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowTimePicker(#12:00:00 PM#)
    Check IsNull(chosen), "modal clear returns Null"
    Check CreateTimePickerForm(True), "restore production events"
    DoCmd.Close acForm, HOST, acSaveNo
    RunTimeTests = "PASS: " & m_checks & " Access time assertions"
    Exit Function
Failed:
    Dim errorText As String
    errorText = "FAIL after " & m_checks & " checks: " & Err.Number & " " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, PICKER, acSaveNo
    DoCmd.Close acForm, HOST, acSaveNo
    RunTimeTests = errorText
End Function
'@

$temporaryDatabase = Join-Path $env:TEMP ('TimeTests-' + [guid]::NewGuid().ToString('N') + '.accdb')
Copy-Item -LiteralPath $DatabasePath -Destination $temporaryDatabase
$access = $null
$databaseOpened = $false
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($temporaryDatabase)
    $databaseOpened = $true
    $project = $access.VBE.ActiveVBProject
    foreach ($entry in @(
        @{ Name = 'Module_TimePicker'; Source = $sourceText },
        @{ Name = 'Module_TimePickerTests'; Source = $testSource }
    )) {
        $component = $null
        foreach ($candidate in $project.VBComponents) {
            if ($candidate.Name -eq $entry.Name) { $component = $candidate; break }
        }
        if ($null -eq $component) {
            $component = $project.VBComponents.Add(1)
            $component.Name = $entry.Name
        }
        if ($component.CodeModule.CountOfLines -gt 0) {
            $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
        }
        $component.CodeModule.AddFromString($entry.Source)
        $access.DoCmd.Save(5, $entry.Name)
    }
    $access.DoCmd.RunCommand(126)
    $result = $access.Eval('RunTimeTests()')
    Write-Output $result
    if ($result -notlike 'PASS:*') { throw $result }
    Write-Output "VBA compile: PASS; maximum consecutive continuations: $longest"
    if ($ScreenshotPath) {
        Add-Type -AssemblyName System.Drawing
        if (-not ('TimeWindowCapture' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class TimeWindowCapture {
    [StructLayout(LayoutKind.Sequential)]
    public struct Rect { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll", ExactSpelling=true, SetLastError=true, CallingConvention=CallingConvention.Winapi)]
    public static extern int GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll", ExactSpelling=true, CallingConvention=CallingConvention.Winapi)]
    public static extern int PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);
}
'@
        }
        $access.Visible = $true
        if (-not $access.Eval('OpenTimeTestPreview()')) { throw 'Preview failed to open' }
        $windowHandle = [IntPtr][long]$access.Eval('[Forms]![frmTimePicker].[hWnd]')
        $bounds = New-Object TimeWindowCapture+Rect
        if ([TimeWindowCapture]::GetWindowRect($windowHandle, [ref]$bounds) -eq 0) {
            throw 'Cannot read time window bounds'
        }
        $bitmap = New-Object Drawing.Bitmap(($bounds.Right - $bounds.Left), ($bounds.Bottom - $bounds.Top))
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $deviceContext = $graphics.GetHdc()
            try {
                if ([TimeWindowCapture]::PrintWindow($windowHandle, $deviceContext, 2) -eq 0) {
                    throw 'Cannot capture time window'
                }
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
} finally {
    $candidate = $null
    if ($null -ne $component) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
        $component = $null
    }
    if ($null -ne $project) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($project)
        $project = $null
    }
    if ($null -ne $access) {
        try {
            if ($databaseOpened) { $access.CloseCurrentDatabase() }
            $access.Quit(2)
        } catch {
            Write-Warning "Access cleanup: $($_.Exception.Message)"
        } finally {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($access)
            $access = $null
        }
    }
    [GC]::Collect()
    Remove-Item -LiteralPath $temporaryDatabase -ErrorAction SilentlyContinue
}
Write-Output 'Temporary database cleanup: PASS'