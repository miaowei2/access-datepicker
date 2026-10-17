param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot '..\DatePicker.accdb'),
    [switch]$VerifyDemo
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot '..\Module_YearMonthPicker.bas'
$encoding = [Text.Encoding]::GetEncoding(936, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
$source = $encoding.GetString([IO.File]::ReadAllBytes($modulePath))
$source = $source -replace '(?m)^Attribute VB_Name[^\r\n]*\r?\n', ''
if ($source.Length -lt 10000 -or $source.Contains([char]0xFFFD)) { throw 'Missing or corrupted month picker source' }
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
    If Not passed Then Err.Raise vbObjectError + 1904, "MonthPickerTests", message
    m_checks = m_checks + 1
End Sub

Public Function RunMonthPickerTests(ByVal verifyLayout As Boolean) As String
    On Error GoTo Failed
    Dim monthNumber As Integer, chineseMonths As Variant
    Dim originalValue As Variant, savedError As String
    Dim fieldNames As Variant, fieldIndex As Integer, previousBottom As Long
    Dim yearInput As Control, candidate As Control, firstYear As Integer
    m_checks = 0
    CreateYearMonthPickerForm True
    DoCmd.OpenForm "Demo", acDesign
    For Each candidate In Forms("Demo").Controls
        If candidate.Name = "txtYear" Then Set yearInput = candidate
    Next candidate
    If yearInput Is Nothing Then
        Set yearInput = CreateControl("Demo", acTextBox, acDetail, , , 1980, 420, 9000, 420)
        yearInput.Name = "txtYear"
    End If
    Set yearInput = Nothing
    Set candidate = Nothing
    DoCmd.Close acForm, "Demo", acSaveYes
    DoCmd.OpenForm "Demo"
    originalValue = Forms("Demo").Controls("Text4").Value
    chineseMonths = Array(ChrW(19968), ChrW(20108), ChrW(19977), ChrW(22235), ChrW(20116), ChrW(20845), _
                         ChrW(19971), ChrW(20843), ChrW(20061), ChrW(21313), ChrW(21313) & ChrW(19968), ChrW(21313) & ChrW(20108))
    For monthNumber = 1 To 12
        Forms("Demo").Controls("Text4").Value = "2026-10"
        Check PickYearMonthForCtl("Demo", "Text4", "yyyy-mm"), "month dropdown opens"
        Check CurrentProject.AllForms("frmYearMonthPicker").IsLoaded, "month form loaded"
        Check Forms("frmYearMonthPicker").Controls("lblM" & Format$(monthNumber, "00")).Caption = chineseMonths(monthNumber - 1) & ChrW(26376), "Chinese month caption"
        Check Forms("frmYearMonthPicker").Controls("lblM" & Format$(monthNumber, "00")).Width >= 1020, "three-character month fits its cell"
        Check YMPicker_MonthClick(monthNumber), "month selection succeeds"
        Check Not CurrentProject.AllForms("frmYearMonthPicker").IsLoaded, "single click confirms"
        Check Forms("Demo").Controls("Text4").Value = "2026-" & Format$(monthNumber, "00"), "yyyy-mm value retained"
    Next monthNumber
    PickYearMonthForCtl "Demo", "Text4"
    YMPicker_NextYear
    YMPicker_MonthClick 1
    Check Forms("Demo").Controls("Text4").Value = "2027-01", "next year"
    PickYearMonthForCtl "Demo", "Text4"
    YMPicker_PrevYear
    YMPicker_MonthClick 12
    Check Forms("Demo").Controls("Text4").Value = "2026-12", "previous year"
    PickYearMonthForCtl "Demo", "Text4"
    YMPicker_PrevDecade
    YMPicker_MonthClick 2
    Check Forms("Demo").Controls("Text4").Value = "2016-02", "previous decade"
    PickYearMonthForCtl "Demo", "Text4"
    YMPicker_NextDecade
    YMPicker_MonthClick 2
    Check Forms("Demo").Controls("Text4").Value = "2026-02", "next decade"
    PickYearMonthForCtl "Demo", "Text4"
    Check Not YMPicker_MonthClick(0), "invalid month rejected"
    Check CurrentProject.AllForms("frmYearMonthPicker").IsLoaded, "invalid month does not close"
    YMPicker_Cancel
    Check Forms("Demo").Controls("Text4").Value = "2026-02", "cancel unchanged"
    PickYearMonthForCtl "Demo", "Text4"
    YMPicker_Clear
    Check IsNull(Forms("Demo").Controls("Text4").Value), "clear writes Null"
    PickYearMonthForCtl "Demo", "Text4"
    YMPicker_GoThisMonth
    YMPicker_OK
    Check Forms("Demo").Controls("Text4").Value = Format$(Date, "yyyy-mm"), "this month"

    For monthNumber = 1 To 12
        Forms("Demo").Controls("txtYear").Value = "2026"
        Check PickYearForCtl("Demo", "txtYear"), "year dropdown opens"
        firstYear = CInt(Forms("frmYearMonthPicker").Controls("lblM01").Caption)
        Check firstYear = 2020, "initial year page"
        Check YMPicker_MonthClick(monthNumber), "year click succeeds"
        Check Not CurrentProject.AllForms("frmYearMonthPicker").IsLoaded, "year click closes"
        Check Forms("Demo").Controls("txtYear").Value = CStr(firstYear + monthNumber - 1), "four-digit year writeback"
    Next monthNumber
    Forms("Demo").Controls("txtYear").Value = "2026"
    PickYearForCtl "Demo", "txtYear"
    YMPicker_NextYear
    Check Forms("frmYearMonthPicker").Controls("lblM01").Caption = "2032", "next twelve years"
    YMPicker_PrevYear
    Check Forms("frmYearMonthPicker").Controls("lblM01").Caption = "2020", "previous twelve years"
    YMPicker_NextDecade
    Check Forms("frmYearMonthPicker").Controls("lblM01").Caption = "2140", "next 120 years"
    YMPicker_PrevDecade
    Check Forms("frmYearMonthPicker").Controls("lblM01").Caption = "2020", "previous 120 years"
    YMPicker_MonthMouseMove 1
    Check Forms("frmYearMonthPicker").Controls("lblM01").BackColor = 16576233, "year hover"
    YMPicker_Cancel
    Check Forms("Demo").Controls("txtYear").Value = "2026", "year cancel unchanged"
    PickYearForCtl "Demo", "txtYear"
    Check Not YMPicker_MonthClick(13), "invalid year cell rejected"
    YMPicker_Clear
    Check IsNull(Forms("Demo").Controls("txtYear").Value), "year clear writes Null"
    PickYearForCtl "Demo", "txtYear"
    YMPicker_GoThisMonth
    Check Forms("Demo").Controls("txtYear").Value = CStr(Year(Date)), "current year confirms"
    Check Not CurrentProject.AllForms("frmYearMonthPicker").IsLoaded, "current year closes"
    Forms("Demo").Controls("txtYear").Value = "bad"
    PickYearForCtl "Demo", "txtYear"
    YMPicker_Cancel
    Check Forms("Demo").Controls("txtYear").Value = "bad", "invalid initial year unchanged on cancel"
    Forms("Demo").Controls("txtYear").Value = "0100"
    PickYearForCtl "Demo", "txtYear"
    YMPicker_PrevDecade
    Check Forms("frmYearMonthPicker").Controls("lblM01").Caption = "0100", "minimum year clamps"
    YMPicker_MonthClick 1
    Check Forms("Demo").Controls("txtYear").Value = "0100", "minimum year writeback"
    Forms("Demo").Controls("txtYear").Value = "9999"
    PickYearForCtl "Demo", "txtYear"
    YMPicker_NextDecade
    Check Forms("frmYearMonthPicker").Controls("lblM12").Caption = "9999", "maximum year clamps"
    YMPicker_MonthClick 12
    Check Forms("Demo").Controls("txtYear").Value = "9999", "maximum year writeback"
    Forms("Demo").Controls("txtYear").Value = "2026"
    PickYearForCtl "Demo", "txtYear", "yyyy" & ChrW(24180)
    YMPicker_MonthClick 7
    Check Forms("Demo").Controls("txtYear").Value = "2026" & ChrW(24180), "custom year format"
    Forms("Demo").Controls("Text4").Value = "2026-10"
    PickYearMonthForCtl "Demo", "Text4"
    Check Forms("frmYearMonthPicker").Controls("lblM01").Caption = chineseMonths(0) & ChrW(26376), "month mode restored after year mode"
    YMPicker_MonthClick 12
    Check Forms("Demo").Controls("Text4").Value = "2026-12", "month writeback remains yyyy-mm"
    If verifyLayout Then
        fieldNames = Array("Text2", "Text0", "txtYear", "Text4", "txtWeek", "txtTime", "txtDateRange", "txtDateTimeRange")
        For fieldIndex = 0 To UBound(fieldNames)
            With Forms("Demo").Controls(fieldNames(fieldIndex))
                Check .Top >= previousBottom, "rows do not overlap"
                Check .TabIndex = fieldIndex, "tab order matches row order"
                Check .Left = 1980 And .Width = 9000 And .Height = 420, "aligned input dimensions"
                Check .Top = 420 + fieldIndex * 720, "consistent row spacing"
                Check .Left + .Width <= Forms("Demo").Width, "input fits form"
                previousBottom = .Top + .Height
            End With
        Next fieldIndex
        Check Forms("Demo").Controls("Label5").Caption = ChrW(26376) & ChrW(20221), "month row label"
        Check Forms("Demo").Controls("Text4").OnClick = "=PickYearMonthForCtl(""Demo"",""Text4"",""yyyy-mm"")", "month click binding"
        Check Forms("Demo").Controls("lblYear").Caption = ChrW(24180) & ChrW(20221), "year row label"
        Check Forms("Demo").Controls("txtYear").OnClick = "=PickYearForCtl(""Demo"",""txtYear"")", "year click binding"
        Check Forms("Demo").Section(acDetail).Height >= previousBottom, "all rows fit"
    End If
    Forms("Demo").Controls("Text4").Value = originalValue
    DoCmd.Close acForm, "Demo", acSaveNo
    RunMonthPickerTests = "PASS: " & m_checks & " year/month/demo assertions"
    Exit Function
Failed:
    savedError = "FAIL: " & Err.Number & " " & Err.Description
    On Error Resume Next
    YMPicker_Cancel
    RunMonthPickerTests = savedError
End Function
'@

$temporaryDatabase = Join-Path $env:TEMP ('MonthPickerTests-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
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
    foreach ($entry in @(@('Module_YearMonthPicker', $source), @('Module_MonthPickerTests', $testSource))) {
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
    $layoutFlag = 'False'
    if ($VerifyDemo) { $layoutFlag = 'True' }
    $result = $access.Eval("RunMonthPickerTests($layoutFlag)")
    Write-Output $result
    if ($result -notlike 'PASS:*') { throw $result }
    Write-Output "VBA compile: PASS; longest continuation run: $longest"
} finally {
    if ($null -ne $project) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($project) }
    if ($null -ne $access) {
        $access.CloseCurrentDatabase()
        if ($null -ne $ownedProcess -and $ownedProcess.Id -notin $existingAccessIds) { Stop-Process -InputObject $ownedProcess -ErrorAction SilentlyContinue }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($access)
    }
}