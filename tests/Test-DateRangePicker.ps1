param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot '..\DatePicker.accdb')
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot '..\Module_DateRangePicker.bas'
$sourceBytes = [System.IO.File]::ReadAllBytes($modulePath)
try {
    $sourceText = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($sourceBytes)
} catch [System.Text.DecoderFallbackException] {
    $sourceText = [System.Text.Encoding]::GetEncoding(936).GetString($sourceBytes)
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

Private Const PICKER As String = "frmDateRangePicker"
Private Const HOST As String = "frmRangeTestHost"
Private m_checks As Long

Private Sub Check(ByVal passed As Boolean, ByVal message As String)
    If Not passed Then
        Err.Raise vbObjectError + 1900, "DateRangePickerTests", message
    End If
    m_checks = m_checks + 1
End Sub

Private Sub MakeHost()
    Dim frm As Form, ctl As Control, temporaryName As String
    Set frm = CreateForm
    temporaryName = frm.Name
    frm.Width = 10500
    frm.Section(acDetail).Height = 1800
    frm.PopUp = True
    frm.AutoCenter = True
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 180, 180, 9000, 420)
    ctl.Name = "txtRange"
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 180, 800, 4200, 420)
    ctl.Name = "txtStart"
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 4800, 800, 4200, 420)
    ctl.Name = "txtEnd"
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename HOST, acForm, temporaryName
    DoCmd.OpenForm HOST
End Sub

Private Sub OpenRange(ByVal value As Variant)
    Forms(HOST).Controls("txtRange").Value = value
    PickDateRangeFor Forms(HOST).Controls("txtRange")
End Sub

Private Function FindCell(ByVal panel As Integer, ByVal dayNumber As Integer) As Integer
    Dim cell As Integer, ctl As Control
    For cell = panel * 42 + 1 To panel * 42 + 42
        Set ctl = Forms(PICKER).Controls("lblD" & Format(cell, "00"))
        If ctl.Caption = CStr(dayNumber) And ctl.ForeColor <> 10526880 Then
            FindCell = cell
            Exit Function
        End If
    Next cell
    Err.Raise 5, , "Cannot find day " & dayNumber
End Function

Public Function RangeTestModalTimer() As Boolean
    On Error GoTo Failed
    Forms(PICKER).TimerInterval = 0
    DateRangePicker_DayClick FindCell(0, 1)
    DateRangePicker_DayClick FindCell(1, 30)
    RangeTestModalTimer = True
    Exit Function
Failed:
    DateRangePicker_Cancel
End Function

Public Function RunDateRangeTests() As String
    On Error GoTo Failed
    Dim ctl As Control, cell As Integer, otherCell As Integer
    Dim chosen As Variant, oldText As String
    Check CreateDateRangePickerForm(True), "form build"
    Check CreateDateRangePickerForm(True), "idempotent rebuild"
    DoCmd.OpenForm PICKER
    Check Left(Forms(PICKER).Controls("lblMonth0").Caption, 4) = CStr(Year(Date)), "direct open initializes month"
    Check Forms(PICKER).Controls("lblD01").OnClick = "=DateRangePicker_DayClick(1)", "first event binding"
    Check Forms(PICKER).Controls("lblD84").OnClick = "=DateRangePicker_DayClick(84)", "last event binding"
    Check Forms(PICKER).Controls("cmdCancel").Cancel, "Escape cancels"
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "empty range cannot confirm"
    Dim names As Variant, name As Variant
    names = Array("lblDisplay", "lblMonth0", "lblMonth1", "lblD01", "lblD84", "cmdOK", "cmdCancel")
    For Each name In names
        Set ctl = Forms(PICKER).Controls(CStr(name))
        Check ctl.Left + ctl.Width <= Forms(PICKER).Width, "horizontal bounds: " & name
        Check ctl.Top + ctl.Height <= Forms(PICKER).Section(acDetail).Height, "vertical bounds: " & name
    Next name
    DateRangePicker_Cancel
    MakeHost

    OpenRange "2026-10-01 ~ 2026-11-30"
    Check Left(Forms(PICKER).Controls("lblMonth0").Caption, 4) = "2026", "default year"
    Check Forms(PICKER).Controls("lblDisplay").Caption = "2026-10-01 ~ 2026-11-30", "defaults loaded"
    Check Forms(PICKER).Controls("cmdOK").Enabled, "complete defaults can confirm"
    cell = FindCell(0, 1)
    Check Forms(PICKER).Controls("lblD" & Format(cell, "00")).BackColor = 16024898, "start endpoint color"
    otherCell = FindCell(1, 30)
    Check Forms(PICKER).Controls("lblD" & Format(otherCell, "00")).BackColor = 16024898, "end endpoint color"
    Check Forms(PICKER).Controls("lblD" & Format(FindCell(1, 15), "00")).BackColor = 16576233, "interval color"
    Check Not DateRangePicker_DayClick(0), "invalid cell ignored"
    DateRangePicker_DayClick cell
    Check CurrentProject.AllForms(PICKER).IsLoaded, "first click stays open"
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "first click cannot confirm"
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-01 ~ 2026-11-30", "first click does not write back"
    DateRangePicker_DayMouseMove FindCell(1, 20)
    Check Forms(PICKER).Controls("lblD" & Format(FindCell(1, 15), "00")).BackColor = 16576233, "hover interval preview"
    DateRangePicker_DayMouseMove 0
    Check Forms(PICKER).Controls("lblD" & Format(FindCell(1, 15), "00")).BackColor = 16777215, "hover preview resets"
    DateRangePicker_DayClick otherCell
    Check Not CurrentProject.AllForms(PICKER).IsLoaded, "second click closes"
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-01 ~ 2026-11-30", "cross-month result"

    OpenRange "2026-10-01 ~ 2026-11-30"
    DateRangePicker_DayClick FindCell(1, 30)
    DateRangePicker_DayMouseMove FindCell(0, 1)
    Check Forms(PICKER).Controls("lblD" & Format(FindCell(0, 15), "00")).BackColor = 16576233, "reverse hover preview"
    DateRangePicker_DayClick FindCell(0, 1)
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-01 ~ 2026-11-30", "reversed endpoints sorted"

    OpenRange "2026-10-01 ~ 2026-11-30"
    cell = FindCell(0, 16)
    DateRangePicker_DayClick cell
    DateRangePicker_DayClick cell
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-16 ~ 2026-10-16", "same-day range"

    OpenRange "2026-12-01 ~ 2027-01-31"
    DateRangePicker_DayClick FindCell(0, 31)
    DateRangePicker_DayClick FindCell(1, 1)
    Check Forms(HOST).Controls("txtRange").Value = "2026-12-31 ~ 2027-01-01", "year boundary"

    OpenRange "2028-02-01 ~ 2028-03-01"
    DateRangePicker_DayClick FindCell(0, 29)
    DateRangePicker_DayClick FindCell(1, 1)
    Check Forms(HOST).Controls("txtRange").Value = "2028-02-29 ~ 2028-03-01", "leap day"

    OpenRange "2026-10-01 ~ 2026-11-30"
    DateRangePicker_DayClick FindCell(0, 15)
    DateRangePicker_Navigate 2
    DateRangePicker_DayClick FindCell(1, 15)
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-15 ~ 2027-01-15", "navigation preserves start"

    OpenRange "2026-10-01 ~ 2026-11-30"
    DateRangePicker_DayClick FindCell(0, 15)
    DateRangePicker_Cancel
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-01 ~ 2026-11-30", "cancel keeps old value"
    OpenRange "2026-10-01 ~ 2026-11-30"
    DateRangePicker_Clear
    Check IsNull(Forms(HOST).Controls("txtRange").Value), "clear single textbox"
    OpenRange "invalid"
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "invalid defaults ignored"
    DateRangePicker_Cancel
    OpenRange Null
    Check Not Forms(PICKER).Controls("cmdOK").Enabled, "Null defaults ignored"
    DateRangePicker_Cancel

    Forms(HOST).Controls("txtStart").Value = #11/30/2026 13:00:00#
    Forms(HOST).Controls("txtEnd").Value = #10/1/2026 19:00:00#
    PickDateRangeFor Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd")
    Check Forms(PICKER).Controls("lblDisplay").Caption = "2026-10-01 ~ 2026-11-30", "two defaults sorted and time stripped"
    DateRangePicker_DayClick FindCell(0, 2)
    DateRangePicker_DayClick FindCell(1, 28)
    Check VarType(Forms(HOST).Controls("txtStart").Value) = vbDate, "start is Date"
    Check VarType(Forms(HOST).Controls("txtEnd").Value) = vbDate, "end is Date"
    Check Forms(HOST).Controls("txtStart").Value = #10/2/2026#, "start written"
    Check Forms(HOST).Controls("txtEnd").Value = #11/28/2026#, "end written"
    PickDateRangeFor Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd"), "yyyy/mm/dd"
    DateRangePicker_OK
    Check Forms(HOST).Controls("txtStart").Value = "2026/10/02", "custom output format"
    PickDateRangeFor Forms(HOST).Controls("txtStart"), Forms(HOST).Controls("txtEnd")
    DateRangePicker_Clear
    Check IsNull(Forms(HOST).Controls("txtStart").Value), "clear start"
    Check IsNull(Forms(HOST).Controls("txtEnd").Value), "clear end"

    AttachDateRangePicker Forms(HOST), "txtRange"
    Check InStr(Forms(HOST).Controls("txtRange").OnDblClick, "PickDateRangeForCtl") > 0, "single attach"
    AttachDateRangePicker Forms(HOST), "txtStart", "txtEnd"
    Check Forms(HOST).Controls("txtStart").OnDblClick = Forms(HOST).Controls("txtEnd").OnDblClick, "both ends attached"
    Check PickDateRangeForCtl(HOST, "txtRange"), "expression API opens"
    DateRangePicker_Cancel

    OpenRange "2026-10-01 ~ 2026-11-30"
    DateRangePicker_DayClick FindCell(0, 15)
    Forms(HOST).SetFocus
    DateRangePicker_Timer
    Check Not CurrentProject.AllForms(PICKER).IsLoaded, "outside click closes"
    Check Forms(HOST).Controls("txtRange").Value = "2026-10-01 ~ 2026-11-30", "outside click does not write"

    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=RangeTestModalTimer()"
    Forms(PICKER).TimerInterval = 50
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowDateRangePicker(#10/1/2026#, #11/30/2026#)
    Check IsArray(chosen), "modal returns array"
    Check VarType(chosen(0)) = vbDate And VarType(chosen(1)) = vbDate, "modal returns Date endpoints"
    Check chosen(0) = #10/1/2026# And chosen(1) = #11/30/2026#, "modal endpoints"
    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=DateRangePicker_Cancel()"
    Forms(PICKER).TimerInterval = 50
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowDateRangePicker()
    Check IsNull(chosen), "modal cancel returns Null"
    Check CreateDateRangePickerForm(True), "restore production events"
    DoCmd.Close acForm, HOST, acSaveNo
    RunDateRangeTests = "PASS: " & m_checks & " Access assertions"
    Exit Function
Failed:
    Dim errorText As String
    errorText = "FAIL after " & m_checks & " checks: " & Err.Number & " " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, PICKER, acSaveNo
    DoCmd.Close acForm, HOST, acSaveNo
    RunDateRangeTests = errorText
End Function
'@

$temporaryDatabase = Join-Path $env:TEMP ('DateRangeTests-' + [guid]::NewGuid().ToString('N') + '.accdb')
Copy-Item -LiteralPath $DatabasePath -Destination $temporaryDatabase
$access = $null
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($temporaryDatabase)
    $project = $access.VBE.ActiveVBProject
    foreach ($entry in @(
        @{ Name = 'Module_DateRangePicker'; Source = $sourceText },
        @{ Name = 'Module_DateRangePickerTests'; Source = $testSource }
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
    $result = $access.Eval('RunDateRangeTests()')
    Write-Output $result
    if ($result -notlike 'PASS:*') { throw $result }
    Write-Output "VBA compile: PASS; maximum consecutive continuations: $longest"
} finally {
    if ($null -ne $access) {
        $access.Quit(2)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($access)
    }
    Remove-Item -LiteralPath $temporaryDatabase -ErrorAction SilentlyContinue
}