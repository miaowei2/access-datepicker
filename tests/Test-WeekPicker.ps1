param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot '..\DatePicker.accdb')
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot '..\Module_WeekPicker.bas'
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

Private Const PICKER As String = "frmWeekPicker"
Private Const HOST As String = "frmWeekTestHost"
Private m_checks As Long

Private Sub Check(ByVal passed As Boolean, ByVal message As String)
    If Not passed Then
        Err.Raise vbObjectError + 1901, "WeekPickerTests", message
    End If
    m_checks = m_checks + 1
End Sub

Private Function WeekText(ByVal text As String) As String
    WeekText = text & ChrW(&H5468)
End Function

Private Sub MakeHost()
    Dim frm As Form, ctl As Control, temporaryName As String
    Set frm = CreateForm
    temporaryName = frm.Name
    frm.Width = 7000
    frm.Section(acDetail).Height = 1600
    frm.PopUp = True
    frm.AutoCenter = True
    Set ctl = CreateControl(temporaryName, acTextBox, acDetail, , , 180, 180, 6000, 420)
    ctl.Name = "txtWeek"
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename HOST, acForm, temporaryName
    DoCmd.OpenForm HOST
End Sub

Private Sub OpenWeek(ByVal value As Variant)
    Forms(HOST).Controls("txtWeek").Value = value
    Check PickWeekFor(Forms(HOST).Controls("txtWeek")), "open dropdown"
End Sub

Public Function WeekTestModalTimer() As Boolean
    On Error GoTo Failed
    Forms(PICKER).TimerInterval = 0
    WeekTestModalTimer = WeekPicker_DayClick(32)
    Exit Function
Failed:
    WeekPicker_Cancel
End Function

Public Function RunWeekTests() As String
    On Error GoTo Failed
    Dim chosen As Variant, cell As Integer, ctl As Control
    Check WeekPicker_WeekStart(#9/27/2026#) = #9/27/2026#, "Sunday starts week"
    Check WeekPicker_WeekStart(#10/3/2026 23:59:59#) = #9/27/2026#, "Saturday belongs to same week and time stripped"
    Check WeekPicker_WeekEnd(#9/30/2026#) = #10/3/2026#, "week ends on Saturday"
    Check WeekPicker_WeekText(#9/27/2026#) = WeekText("2026-40"), "screenshot label"
    Check WeekPicker_WeekStart(WeekText("2026-40")) = #9/27/2026#, "label parsing"
    Check WeekPicker_WeekStart(" " & WeekText("2026-40") & " ") = #9/27/2026#, "whitespace parsing"
    Check WeekPicker_WeekText(#12/27/2020#) = WeekText("2021-01"), "December week belongs to new year"
    Check WeekPicker_WeekText(#1/1/2021#) = WeekText("2021-01"), "January matches December week"
    Check WeekPicker_WeekText(#1/3/2021#) = WeekText("2021-02"), "next Sunday advances week"
    Check WeekPicker_WeekText(#1/1/2023#) = WeekText("2023-01"), "Sunday January first"
    Check WeekPicker_WeekText(#2/29/2028#) = WeekText("2028-10"), "leap day numbering"
    Check WeekPicker_WeekText(#12/24/2023#) = WeekText("2023-52"), "year before new week year"
    Check WeekPicker_WeekText(#12/31/2023#) = WeekText("2024-01"), "year rolls on Sunday"
    Check WeekPicker_WeekStart(WeekText("2022-53")) = #12/25/2022#, "valid week 53"
    Check IsNull(WeekPicker_WeekStart(WeekText("2026-53"))), "invalid week 53 rejected"
    Check IsNull(WeekPicker_WeekStart(WeekText("2026-00"))), "week zero rejected"
    Check IsNull(WeekPicker_WeekStart(WeekText("2026-54"))), "week 54 rejected"
    Check IsNull(WeekPicker_WeekStart(WeekText("2026-4"))), "malformed label rejected"
    Check IsNull(WeekPicker_WeekStart("invalid")), "invalid default is Null"
    Check IsNull(WeekPicker_WeekStart(Null)), "Null helper result"
    Check Len(WeekPicker_WeekText(Null)) = 0, "invalid week text is empty string"
    Check IsNull(WeekPicker_WeekEnd(Null)), "invalid week end is Null"
    Dim testYear As Integer, offset As Integer, dayValue As Date, startValue As Variant
    For testYear = 2019 To 2030
        For offset = 0 To 14
            dayValue = DateSerial(testYear, 12, 25) + offset
            startValue = WeekPicker_WeekStart(dayValue)
            Check WeekPicker_WeekStart(WeekPicker_WeekText(dayValue)) = startValue, "cross-year label round trip"
            Check WeekPicker_WeekText(startValue) = WeekPicker_WeekText(startValue + 6), "whole week shares label"
        Next offset
    Next testYear

    Check CreateWeekPickerForm(True), "form build"
    Check CreateWeekPickerForm(True), "idempotent rebuild"
    DoCmd.OpenForm PICKER
    Check Left(Forms(PICKER).Controls("lblMonth").Caption, 4) = CStr(Year(Date)), "direct open initializes year"
    Check Forms(PICKER).Controls("lblD01").OnClick = "=WeekPicker_DayClick(1)", "first event binding"
    Check Forms(PICKER).Controls("lblD42").OnClick = "=WeekPicker_DayClick(42)", "last event binding"
    Check Forms(PICKER).Controls("cmdCancel").Cancel, "Escape cancels"
    Check Not WeekPicker_DayClick(0), "invalid cell ignored"
    Check Not WeekPicker_DayClick(43), "cell out of range ignored"
    For Each ctl In Forms(PICKER).Controls
        Check ctl.Left + ctl.Width <= Forms(PICKER).Width, "horizontal bounds: " & ctl.Name
        Check ctl.Top + ctl.Height <= Forms(PICKER).Section(acDetail).Height, "vertical bounds: " & ctl.Name
    Next ctl
    WeekPicker_Cancel
    MakeHost

    OpenWeek WeekText("2026-40")
    Check Forms(PICKER).Controls("lblDisplay").Caption = WeekText("2026-40"), "default label restored"
    Check InStr(Forms(PICKER).Controls("lblMonth").Caption, "9") > 0, "default shows September"
    Check Forms(PICKER).Controls("lblD29").Caption = "27", "Sunday endpoint caption"
    Check Forms(PICKER).Controls("lblD35").Caption = "3", "next-month Saturday endpoint caption"
    Check Forms(PICKER).Controls("lblD29").BackColor = 16024898, "Sunday endpoint blue"
    Check Forms(PICKER).Controls("lblD35").BackColor = 16024898, "Saturday endpoint blue"
    For cell = 30 To 34
        Check Forms(PICKER).Controls("lblD" & Format(cell, "00")).BackColor = 16576233, "whole selected row highlighted"
    Next cell
    WeekPicker_DayMouseMove 9
    For cell = 8 To 14
        Check Forms(PICKER).Controls("lblD" & Format(cell, "00")).BackColor = 16576233, "hover whole row"
    Next cell
    WeekPicker_DayMouseMove 0
    Check Forms(PICKER).Controls("lblD09").BackColor = 16777215, "hover row resets"
    WeekPicker_Cancel
    For cell = 29 To 35
        OpenWeek WeekText("2026-40")
        Check WeekPicker_DayClick(cell), "click confirms"
        Check Not CurrentProject.AllForms(PICKER).IsLoaded, "click closes"
        Check Forms(HOST).Controls("txtWeek").Value = WeekText("2026-40"), "every weekday selects same week"
    Next cell
    OpenWeek WeekText("2026-40")
    WeekPicker_DayClick 1
    Check Forms(HOST).Controls("txtWeek").Value = WeekText("2026-36"), "previous-month cell selects its week"
    OpenWeek #10/16/2026 12:30:00#
    Check Forms(PICKER).Controls("lblDisplay").Caption = WeekText("2026-42"), "date default interpreted"
    WeekPicker_Cancel
    OpenWeek WeekText("2026-40")
    WeekPicker_Navigate 4
    Check Left(Forms(PICKER).Controls("lblMonth").Caption, 4) = "2027", "cross-year month navigation"
    Check Forms(PICKER).Controls("lblDisplay").Caption = WeekText("2026-40"), "navigation preserves selection"
    WeekPicker_Cancel
    Check Forms(HOST).Controls("txtWeek").Value = WeekText("2026-40"), "cancel does not write"
    OpenWeek WeekText("2026-40")
    WeekPicker_ThisWeek
    Check Forms(PICKER).Controls("lblDisplay").Caption = WeekPicker_WeekText(Date), "this week button"
    WeekPicker_Cancel

    OpenWeek WeekText("2026-40")
    WeekPicker_Clear
    Check IsNull(Forms(HOST).Controls("txtWeek").Value), "clear writes Null"
    OpenWeek Null
    Check Forms(PICKER).Controls("lblDisplay").Caption = WeekPicker_WeekText(Date), "empty defaults to current week"
    WeekPicker_Cancel
    Forms(HOST).Controls("txtWeek").Value = WeekText("2026-40")
    Check PickWeekFor(Forms(HOST).Controls("txtWeek"), True), "date output opens"
    WeekPicker_DayClick 32
    Check VarType(Forms(HOST).Controls("txtWeek").Value) = vbDate, "date output type"
    Check Forms(HOST).Controls("txtWeek").Value = #9/27/2026#, "date output is Sunday"
    AttachWeekPicker Forms(HOST), "txtWeek", True
    Check Forms(HOST).Controls("txtWeek").OnDblClick = "=PickWeekForCtl(""" & HOST & """,""txtWeek"",True)", "attach expression"
    Check PickWeekForCtl(HOST, "txtWeek"), "expression API opens"
    WeekPicker_Cancel

    OpenWeek WeekText("2026-40")
    Forms(HOST).SetFocus
    WeekPicker_Timer
    Check Not CurrentProject.AllForms(PICKER).IsLoaded, "outside closes dropdown"
    Check Forms(HOST).Controls("txtWeek").Value = WeekText("2026-40"), "outside preserves value"
    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=WeekTestModalTimer()"
    Forms(PICKER).TimerInterval = 50
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowWeekPicker(#9/27/2026#)
    Check VarType(chosen) = vbDate, "modal returns Date"
    Check chosen = #9/27/2026#, "modal returns Sunday"
    DoCmd.OpenForm PICKER, acDesign
    Forms(PICKER).OnTimer = "=WeekPicker_Cancel()"
    Forms(PICKER).TimerInterval = 50
    DoCmd.Close acForm, PICKER, acSaveYes
    chosen = ShowWeekPicker()
    Check IsNull(chosen), "modal cancel returns Null"
    Check CreateWeekPickerForm(True), "restore production events"
    DoCmd.Close acForm, HOST, acSaveNo
    RunWeekTests = "PASS: " & m_checks & " Access week assertions"
    Exit Function
Failed:
    Dim errorText As String
    errorText = "FAIL after " & m_checks & " checks: " & Err.Number & " " & Err.Description
    On Error Resume Next
    DoCmd.Close acForm, PICKER, acSaveNo
    DoCmd.Close acForm, HOST, acSaveNo
    RunWeekTests = errorText
End Function
'@

$temporaryDatabase = Join-Path $env:TEMP ('WeekTests-' + [guid]::NewGuid().ToString('N') + '.accdb')
Copy-Item -LiteralPath $DatabasePath -Destination $temporaryDatabase
$access = $null
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($temporaryDatabase)
    $project = $access.VBE.ActiveVBProject
    foreach ($entry in @(
        @{ Name = 'Module_WeekPicker'; Source = $sourceText },
        @{ Name = 'Module_WeekPickerTests'; Source = $testSource }
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
    $result = $access.Eval('RunWeekTests()')
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