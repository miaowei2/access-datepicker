Option Compare Database
Option Explicit

'====================================================
' 模块名称: Module_DateTimeRangePicker
' 功能描述: Access 纯窗体时间范围选择器，选择起止日期及时间
'
' 特点:
'   - 双月日历配合两组时、分、秒下拉框，精确到秒
'   - 选完两个日期仍可调整时间，只有确定才回写
'   - 支持单个合并文本框或两个 Date 文本框，自动排序完整时间戳
'
' 使用方法:
'   1. 将本模块以 GBK 编码导入 Access 数据库
'   2. 在立即窗口运行: CreateDateTimeRangePickerForm
'   3. 在代码中调用:
'      result = ShowDateTimeRangePicker()          ' 返回 Array(Date, Date) 或 Null
'      PickDateTimeRangeFor Me.txtRange            ' 合并文本，默认保留秒
'      PickDateTimeRangeFor Me.txtStart, Me.txtEnd  ' 分别回写两个 Date
'====================================================

' ============ Win32 API 声明 (定位下拉窗体，兼容 32/64 位 Access) ============
#If VBA7 Then
    Private Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr
    Private Declare PtrSafe Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function apiMonitorFromRect Lib "user32" Alias "MonitorFromRect" (lpRect As RECT, ByVal flags As Long) As LongPtr
    Private Declare PtrSafe Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal monitor As LongPtr, lpInfo As MONITORINFO) As Long
    Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As LongPtr, ByVal left As Long, ByVal top As Long, ByVal width As Long, ByVal height As Long, ByVal repaint As Long) As Long
#Else
    Private Declare Function apiGetFocus Lib "user32" Alias "GetFocus" () As Long
    Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function apiMonitorFromRect Lib "user32" Alias "MonitorFromRect" (lpRect As RECT, ByVal flags As Long) As Long
    Private Declare Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal monitor As Long, lpInfo As MONITORINFO) As Long
    Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal left As Long, ByVal top As Long, ByVal width As Long, ByVal height As Long, ByVal repaint As Long) As Long
#End If

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    flags As Long
End Type

' ============ 窗体名、分隔符与颜色常量 ============
Private Const RANGE_FORM As String = "frmDateTimeRangePicker"
Private Const SEPARATOR As String = " ~ "
Private Const OUTPUT_FORMAT As String = "yyyy-mm-dd hh:nn:ss"
Private Const CLR_BLUE As Long = 16024898
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_BLACK As Long = 0
Private Const CLR_GRAY As Long = 10526880
Private Const CLR_BORDER As Long = 13158600
Private Const CLR_RANGE As Long = 16576233

' ============ 模块级变量 ============
Private m_viewMonth As Date              ' 左侧月份第 1 天
Private m_startDate As Variant           ' 起始日期，不含时间；未选时为 Null
Private m_endDate As Variant             ' 结束日期，不含时间；未选时为 Null
Private m_startTime As Date              ' 起始时、分、秒
Private m_endTime As Date                ' 结束时、分、秒
Private m_waitingEnd As Boolean          ' 已选起点，正在等待终点
Private m_initialized As Boolean         ' 是否由入口初始化
Private m_nonModal As Boolean            ' 是否为文本框下拉模式
Private m_hoverCell As Integer           ' 当前悬停单元格，0 表示无
Private m_cellDates(1 To 84) As Variant   ' 双月日历的日期映射
Private m_result As Variant              ' 模态结果: Array(Date, Date) 或 Null
Private m_startInput As Control          ' 合并框或起始日期时间框
Private m_endInput As Control            ' 结束框；合并模式为 Nothing
Private m_outputFormat As String         ' 自定义回写格式；双框为空时写 Date

'====================================================
' 公共 API: 模态选择时间范围
' 参数: dtStart、dtEnd 为可选默认日期时间，保留到秒
' 返回: 按先后排序的 Array(起始 Date, 结束 Date)；取消或清除为 Null
' 示例: result = ShowDateTimeRangePicker(Now, DateAdd("h", 1, Now))
'====================================================
Public Function ShowDateTimeRangePicker(Optional ByVal dtStart As Variant, _
                                        Optional ByVal dtEnd As Variant) As Variant
    On Error GoTo Failed
    ShowDateTimeRangePicker = Null
    If Not DateTimeRange_Ready() Then
        Exit Function
    End If
    DateTimeRange_CloseExisting
    DateTimeRange_Initialize dtStart, dtEnd
    DoCmd.OpenForm RANGE_FORM, acNormal, , , , acDialog
    ShowDateTimeRangePicker = m_result
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "时间范围选择器"
End Function

'====================================================
' 通用 API: 弹出时间范围下拉窗体并回写一个或两个文本框
' 参数: ctlStart 为合并框或起始框；ctlEnd 为可选结束框
'       sFormat 为空时单框写 yyyy-mm-dd hh:nn:ss 范围，双框写 Date
'       指定 sFormat 时写格式化字符串，分隔符固定为 " ~ "
' 返回: 成功打开为 True；确定前不回写，取消不改原值，清除写 Null
' 示例: PickDateTimeRangeFor Me.txtStart, Me.txtEnd
'====================================================
Public Function PickDateTimeRangeFor(ctlStart As Control, _
                                     Optional ctlEnd As Control, _
                                     Optional ByVal sFormat As String = "") As Boolean
    On Error GoTo Failed
    If Not DateTimeRange_Ready() Then
        Exit Function
    End If
    If Not TypeOf ctlStart Is TextBox Then
        Err.Raise 5, , "开始控件必须是文本框"
    End If
    If Not ctlEnd Is Nothing Then
        If Not TypeOf ctlEnd Is TextBox Then
            Err.Raise 5, , "结束控件必须是文本框"
        End If
        If ctlStart Is ctlEnd Then
            Err.Raise 5, , "起止时间须使用不同的文本框"
        End If
    End If
    DateTimeRange_CloseExisting
    Dim startValue As Variant, endValue As Variant, parts As Variant
    startValue = Null
    endValue = Null
    If ctlEnd Is Nothing Then
        parts = Split(Nz(ctlStart.value, ""), SEPARATOR)
        If UBound(parts) = 1 Then
            startValue = parts(0)
            endValue = parts(1)
        End If
    Else
        startValue = ctlStart.value
        endValue = ctlEnd.value
    End If
    DateTimeRange_Initialize startValue, endValue
    Set m_startInput = ctlStart
    Set m_endInput = ctlEnd
    m_outputFormat = sFormat
    m_nonModal = True
    Dim inputRect As RECT
    ctlStart.SetFocus
    apiGetWindowRect apiGetFocus(), inputRect
    DoCmd.OpenForm RANGE_FORM, acNormal
    DateTimeRange_Position inputRect
    PickDateTimeRangeFor = True
    Exit Function
Failed:
    Set m_startInput = Nothing
    Set m_endInput = Nothing
    m_nonModal = False
    MsgBox Err.Description, vbExclamation, "时间范围选择器"
End Function

' 表达式入口: =PickDateTimeRangeForCtl("Demo","txtDateTimeRange")。
' 两框模式可再传结束控件名和自定义格式。
Public Function PickDateTimeRangeForCtl(ByVal sFormName As String, _
                                       ByVal sStartCtlName As String, _
                                       Optional ByVal sEndCtlName As String = "", _
                                       Optional ByVal sFormat As String = "") As Boolean
    On Error GoTo Failed
    Dim endControl As Control
    If sEndCtlName <> "" Then
        Set endControl = Forms(sFormName).Controls(sEndCtlName)
    End If
    PickDateTimeRangeForCtl = PickDateTimeRangeFor(Forms(sFormName).Controls(sStartCtlName), endControl, sFormat)
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "时间范围选择器"
End Function

' 自动绑定一个或两个文本框的双击事件，参数含义与 PickDateTimeRangeFor 相同。
Public Sub AttachDateTimeRangePicker(frm As Form, ByVal sStartCtlName As String, _
                                     Optional ByVal sEndCtlName As String = "", _
                                     Optional ByVal sFormat As String = "")
    Dim expression As String
    expression = "=PickDateTimeRangeForCtl(""" & Replace(frm.name, """", """""") & """,""" & _
                 Replace(sStartCtlName, """", """""") & """,""" & _
                 Replace(sEndCtlName, """", """""") & """,""" & _
                 Replace(sFormat, """", """""") & """)"
    frm.Controls(sStartCtlName).OnDblClick = expression
    If sEndCtlName <> "" Then
        frm.Controls(sEndCtlName).OnDblClick = expression
    End If
End Sub

' ============ 日期时间初始化与排序 ============
Private Sub DateTimeRange_Initialize(ByVal dtStart As Variant, ByVal dtEnd As Variant)
    m_startDate = Null
    m_endDate = Null
    m_startTime = TimeSerial(0, 0, 0)
    m_endTime = TimeSerial(23, 59, 59)
    If IsDate(dtStart) Then
        m_startDate = dateValue(CDate(dtStart))
        m_startTime = TimeSerial(Hour(CDate(dtStart)), Minute(CDate(dtStart)), Second(CDate(dtStart)))
    End If
    If IsDate(dtEnd) Then
        m_endDate = dateValue(CDate(dtEnd))
        m_endTime = TimeSerial(Hour(CDate(dtEnd)), Minute(CDate(dtEnd)), Second(CDate(dtEnd)))
    End If
    If IsNull(m_startDate) And Not IsNull(m_endDate) Then
        m_startDate = m_endDate
        m_startTime = m_endTime
        m_endDate = Null
        m_endTime = TimeSerial(23, 59, 59)
    End If
    DateTimeRange_Sort
    Dim baseDate As Date
    baseDate = Date
    If Not IsNull(m_startDate) Then
        baseDate = m_startDate
    End If
    m_viewMonth = DateSerial(Year(baseDate), Month(baseDate), 1)
    If m_viewMonth > DateSerial(9999, 11, 1) Then
        m_viewMonth = DateSerial(9999, 11, 1)
    End If
    m_waitingEnd = False
    m_hoverCell = 0
    m_initialized = True
    m_nonModal = False
    m_result = Null
    m_outputFormat = ""
    Set m_startInput = Nothing
    Set m_endInput = Nothing
End Sub

Private Function DateTimeRange_Combine(ByVal dayValue As Date, ByVal timeValue As Date) As Date
    ' 用 DateAdd 组合，避免直接相加 OLE Date 的负数日期导致公元早期时间偏移。
    DateTimeRange_Combine = DateAdd("s", CLng(Hour(timeValue)) * 3600 + _
                                   CLng(Minute(timeValue)) * 60 + Second(timeValue), dayValue)
End Function

Private Sub DateTimeRange_Sort()
    If IsNull(m_startDate) Or IsNull(m_endDate) Then
        Exit Sub
    End If
    ' 先比较日期再比较纯时间，避免负数 OLE Date 将同日时刻按相反顺序排列。
    If m_startDate > m_endDate Or (m_startDate = m_endDate And m_startTime > m_endTime) Then
        Dim temporaryDay As Variant, temporaryTime As Date
        temporaryDay = m_startDate
        m_startDate = m_endDate
        m_endDate = temporaryDay
        temporaryTime = m_startTime
        m_startTime = m_endTime
        m_endTime = temporaryTime
    End If
End Sub

' ============ 窗体与控件事件 (由表达式调用) ============
Public Function DateTimeRangePicker_FormLoad() As Boolean
    If Not m_initialized Then
        DateTimeRange_Initialize Null, Null
    End If
    DateTimeRange_Refresh
    DateTimeRangePicker_FormLoad = True
End Function

Public Function DateTimeRangePicker_DayClick(ByVal cell As Integer) As Boolean
    If cell < 1 Or cell > 84 Then
        Exit Function
    End If
    If IsNull(m_cellDates(cell)) Then
        Exit Function
    End If
    If Not m_waitingEnd Then
        m_startDate = m_cellDates(cell)
        m_endDate = Null
        m_waitingEnd = True
    Else
        m_endDate = m_cellDates(cell)
        m_waitingEnd = False
        DateTimeRange_Sort
    End If
    ' 与仅日期范围不同，选完终点仍留在窗体里，允许继续调整秒值。
    m_hoverCell = 0
    DateTimeRange_Refresh
    DateTimeRangePicker_DayClick = True
End Function

Public Function DateTimeRangePicker_DayMouseMove(ByVal cell As Integer) As Boolean
    If cell < 0 Or cell > 84 Then
        Exit Function
    End If
    If cell <> m_hoverCell Then
        m_hoverCell = cell
        DateTimeRange_Refresh
    End If
    DateTimeRangePicker_DayMouseMove = True
End Function

' 顶部日期输入与日历选择共用状态；无效输入恢复原值，清空则撤销该端点。
Public Function DateTimeRangePicker_DateChanged(ByVal endpoint As Integer) As Boolean
    On Error GoTo InvalidValue
    Dim value As Variant
    If endpoint = 0 Then
        value = Forms(RANGE_FORM).Controls("txtStartDate").value
    ElseIf endpoint = 1 Then
        value = Forms(RANGE_FORM).Controls("txtEndDate").value
    Else
        Exit Function
    End If
    If Not IsNull(value) Then
        If Not IsDate(value) Then
            GoTo InvalidValue
        End If
        value = dateValue(CDate(value))
    End If
    If endpoint = 0 Then
        m_startDate = value
    Else
        m_endDate = value
    End If
    m_waitingEnd = Not IsNull(m_startDate) And IsNull(m_endDate)
    m_hoverCell = 0
    DateTimeRange_Refresh
    DateTimeRangePicker_DateChanged = True
    Exit Function
InvalidValue:
    DateTimeRange_Refresh
End Function

Public Function DateTimeRangePicker_TimeChanged(ByVal endpoint As Integer) As Boolean
    On Error GoTo InvalidValue
    Dim prefix As String, frm As Form
    If endpoint = 0 Then
        prefix = "cboStart"
    ElseIf endpoint = 1 Then
        prefix = "cboEnd"
    Else
        Exit Function
    End If
    Set frm = Forms(RANGE_FORM)
    Dim hourValue As Variant, minuteValue As Variant, secondValue As Variant
    hourValue = frm.Controls(prefix & "Hour").value
    minuteValue = frm.Controls(prefix & "Min").value
    secondValue = frm.Controls(prefix & "Sec").value
    If IsNull(hourValue) Or IsNull(minuteValue) Or IsNull(secondValue) Then
        GoTo InvalidValue
    End If
    If Not IsNumeric(hourValue) Or Not IsNumeric(minuteValue) Or Not IsNumeric(secondValue) Then
        GoTo InvalidValue
    End If
    If hourValue < 0 Or hourValue > 23 Or minuteValue < 0 Or minuteValue > 59 Or secondValue < 0 Or secondValue > 59 Then
        GoTo InvalidValue
    End If
    If endpoint = 0 Then
        m_startTime = TimeSerial(CInt(hourValue), CInt(minuteValue), CInt(secondValue))
    Else
        m_endTime = TimeSerial(CInt(hourValue), CInt(minuteValue), CInt(secondValue))
    End If
    DateTimeRange_Refresh
    DateTimeRangePicker_TimeChanged = True
    Exit Function
InvalidValue:
    DateTimeRange_Refresh
End Function

Public Function DateTimeRangePicker_Navigate(ByVal months As Integer) As Boolean
    Dim monthNumber As Long
    monthNumber = CLng(Year(m_viewMonth)) * 12 + Month(m_viewMonth) - 1 + CLng(months)
    If monthNumber < 1200 Or monthNumber > 119998 Then
        Exit Function
    End If
    m_viewMonth = DateSerial(monthNumber \ 12, monthNumber Mod 12 + 1, 1)
    m_hoverCell = 0
    DateTimeRange_Refresh
    DateTimeRangePicker_Navigate = True
End Function

' “今天”仅导航回当前月，不修改已选的日期和时间。
Public Function DateTimeRangePicker_Today() As Boolean
    m_viewMonth = DateSerial(Year(Date), Month(Date), 1)
    m_hoverCell = 0
    DateTimeRange_Refresh
    DateTimeRangePicker_Today = True
End Function

Public Function DateTimeRangePicker_OK() As Boolean
    If IsNull(m_startDate) Or IsNull(m_endDate) Then
        Exit Function
    End If
    DateTimeRange_Sort
    If Not DateTimeRange_WriteBack(False) Then
        Exit Function
    End If
    m_result = Array(DateTimeRange_Combine(m_startDate, m_startTime), DateTimeRange_Combine(m_endDate, m_endTime))
    DoCmd.Close acForm, RANGE_FORM, acSaveNo
    DateTimeRangePicker_OK = True
End Function

Public Function DateTimeRangePicker_Clear() As Boolean
    If Not DateTimeRange_WriteBack(True) Then
        Exit Function
    End If
    m_result = Null
    DoCmd.Close acForm, RANGE_FORM, acSaveNo
    DateTimeRangePicker_Clear = True
End Function

Public Function DateTimeRangePicker_Cancel() As Boolean
    m_result = Null
    DoCmd.Close acForm, RANGE_FORM, acSaveNo
    DateTimeRangePicker_Cancel = True
End Function

Public Function DateTimeRangePicker_FormClose() As Boolean
    Set m_startInput = Nothing
    Set m_endInput = Nothing
    m_initialized = False
    m_nonModal = False
    DateTimeRangePicker_FormClose = True
End Function

Public Function DateTimeRangePicker_Deactivate() As Boolean
    If m_nonModal Then
        Forms(RANGE_FORM).TimerInterval = 100
    End If
    DateTimeRangePicker_Deactivate = True
End Function

Public Function DateTimeRangePicker_Timer() As Boolean
    On Error GoTo NoActiveForm
    Forms(RANGE_FORM).TimerInterval = 0
    If m_nonModal Then
        If Screen.ActiveForm.name <> RANGE_FORM Then
            DateTimeRangePicker_Cancel
        End If
    End If
    DateTimeRangePicker_Timer = True
    Exit Function
NoActiveForm:
    If m_nonModal Then
        DateTimeRangePicker_Cancel
    End If
End Function

' ============ 预览绘制与原子回写 ============
Private Sub DateTimeRange_Refresh()
    Dim frm As Form, endpoint As Integer, prefix As String, timeValue As Date
    Set frm = Forms(RANGE_FORM)
    frm.Controls("txtStartDate").value = m_startDate
    frm.Controls("txtEndDate").value = m_endDate
    For endpoint = 0 To 1
        prefix = "cboStart"
        timeValue = m_startTime
        If endpoint = 1 Then
            prefix = "cboEnd"
            timeValue = m_endTime
        End If
        frm.Controls(prefix & "Hour").value = Format$(Hour(timeValue), "00")
        frm.Controls(prefix & "Min").value = Format$(Minute(timeValue), "00")
        frm.Controls(prefix & "Sec").value = Format$(Second(timeValue), "00")
    Next endpoint
    Dim startText As String, endText As String
    startText = "--"
    endText = "--"
    If Not IsNull(m_startDate) Then
        startText = Format$(DateTimeRange_Combine(m_startDate, m_startTime), OUTPUT_FORMAT)
    End If
    If Not IsNull(m_endDate) Then
        endText = Format$(DateTimeRange_Combine(m_endDate, m_endTime), OUTPUT_FORMAT)
    End If
    frm.Controls("lblDisplay").caption = startText & SEPARATOR & endText
    frm.Controls("cmdOK").Enabled = Not IsNull(m_startDate) And Not IsNull(m_endDate)
    Dim rangeStart As Variant, rangeEnd As Variant, temporary As Variant
    rangeStart = m_startDate
    rangeEnd = m_endDate
    If m_waitingEnd And m_hoverCell > 0 Then
        rangeEnd = m_cellDates(m_hoverCell)
    End If
    If Not IsNull(rangeStart) And Not IsNull(rangeEnd) Then
        If rangeStart > rangeEnd Then
            temporary = rangeStart
            rangeStart = rangeEnd
            rangeEnd = temporary
        End If
    End If
    Dim panel As Integer, cell As Integer, index As Integer, ctl As Control
    Dim firstDate As Date, gridStart As Double, serial As Double, cellDate As Date
    For panel = 0 To 1
        firstDate = DateAdd("m", panel, m_viewMonth)
        frm.Controls("lblMonth" & panel).caption = Format$(firstDate, "yyyy 年 m 月")
        gridStart = CDbl(firstDate) - Weekday(firstDate, vbSunday) + 1
        For cell = 1 To 42
            index = panel * 42 + cell
            Set ctl = frm.Controls("lblD" & Format$(index, "00"))
            serial = gridStart + cell - 1
            m_cellDates(index) = Null
            ctl.caption = ""
            ctl.BackColor = CLR_WHITE
            ctl.ForeColor = CLR_BLACK
            ctl.FontWeight = 400
            If serial >= CDbl(DateSerial(100, 1, 1)) And serial <= CDbl(DateSerial(9999, 12, 31)) Then
                cellDate = CDate(serial)
                m_cellDates(index) = cellDate
                ctl.caption = CStr(Day(cellDate))
                If Month(cellDate) <> Month(firstDate) Then
                    ctl.ForeColor = CLR_GRAY
                End If
                If cellDate = Date Then
                    ctl.ForeColor = CLR_BLUE
                    ctl.FontWeight = 700
                End If
                If Not IsNull(rangeStart) And Not IsNull(rangeEnd) Then
                    If cellDate >= rangeStart And cellDate <= rangeEnd Then
                        ctl.BackColor = CLR_RANGE
                    End If
                End If
                If cellDate = m_startDate Or cellDate = m_endDate Then
                    ctl.BackColor = CLR_BLUE
                    ctl.ForeColor = CLR_WHITE
                    ctl.FontWeight = 700
                ElseIf index = m_hoverCell Then
                    ctl.BackColor = CLR_RANGE
                End If
            End If
        Next cell
    Next panel
End Sub

Private Function DateTimeRange_WriteBack(ByVal clearValues As Boolean) As Boolean
    On Error GoTo Failed
    DateTimeRange_WriteBack = True
    If Not m_nonModal Or m_startInput Is Nothing Then
        Exit Function
    End If
    Dim oldStart As Variant, oldEnd As Variant, startValue As Variant, endValue As Variant
    oldStart = m_startInput.value
    If Not m_endInput Is Nothing Then
        oldEnd = m_endInput.value
    End If
    startValue = Null
    endValue = Null
    If Not clearValues Then
        startValue = DateTimeRange_Combine(m_startDate, m_startTime)
        endValue = DateTimeRange_Combine(m_endDate, m_endTime)
        If m_endInput Is Nothing Then
            Dim formatText As String
            formatText = m_outputFormat
            If formatText = "" Then
                formatText = OUTPUT_FORMAT
            End If
            startValue = Format$(startValue, formatText) & SEPARATOR & Format$(endValue, formatText)
        ElseIf m_outputFormat <> "" Then
            startValue = Format$(startValue, m_outputFormat)
            endValue = Format$(endValue, m_outputFormat)
        End If
    End If
    m_startInput.value = startValue
    If Not m_endInput Is Nothing Then
        m_endInput.value = endValue
    End If
    Exit Function
Failed:
    Dim errorText As String
    errorText = Err.Description
    ' 两框回写失败时恢复原值，不留下只修改一个端点的范围。
    On Error Resume Next
    m_startInput.value = oldStart
    If Not m_endInput Is Nothing Then
        m_endInput.value = oldEnd
    End If
    DateTimeRange_WriteBack = False
    MsgBox "无法回写时间范围: " & errorText, vbExclamation, "时间范围选择器"
End Function

Private Function DateTimeRange_Ready() As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.name = RANGE_FORM Then
            DateTimeRange_Ready = True
            Exit Function
        End If
    Next obj
    MsgBox "请先运行 CreateDateTimeRangePickerForm 创建时间范围窗体。", vbExclamation, "时间范围选择器"
End Function

Private Sub DateTimeRange_CloseExisting()
    If CurrentProject.AllForms(RANGE_FORM).IsLoaded Then
        DoCmd.Close acForm, RANGE_FORM, acSaveNo
    End If
End Sub

Private Sub DateTimeRange_Position(ByRef inputRect As RECT)
    Dim windowRect As RECT, monitor As MONITORINFO, frm As Form
    Set frm = Forms(RANGE_FORM)
    apiGetWindowRect frm.hwnd, windowRect
    monitor.cbSize = LenB(monitor)
    If apiGetMonitorInfo(apiMonitorFromRect(inputRect, 2), monitor) = 0 Then
        Exit Sub
    End If
    Dim width As Long, height As Long, left As Long, top As Long
    width = windowRect.right - windowRect.left
    height = windowRect.bottom - windowRect.top
    left = inputRect.left
    top = inputRect.bottom
    If left + width > monitor.rcWork.right Then
        left = monitor.rcWork.right - width
    End If
    If left < monitor.rcWork.left Then
        left = monitor.rcWork.left
    End If
    If top + height > monitor.rcWork.bottom Then
        top = inputRect.top - height
    End If
    If top < monitor.rcWork.top Then
        top = monitor.rcWork.top
    End If
    apiMoveWindow frm.hwnd, left, top, width, height, True
End Sub

' ============ 窗体生成辅助过程 ============
Private Function DateTimeRange_Label(ByVal formName As String, ByVal name As String, _
                                     ByVal caption As String, ByVal left As Long, _
                                     ByVal top As Long, ByVal width As Long, ByVal height As Long) As Control
    Dim ctl As Control
    Set ctl = CreateControl(formName, acLabel, acDetail, , , left, top, width, height)
    ctl.name = name
    ctl.caption = caption
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.TextAlign = 2
    ctl.BackStyle = 1
    ctl.BackColor = CLR_WHITE
    ctl.ForeColor = CLR_BLACK
    ctl.BorderStyle = 0
    Set DateTimeRange_Label = ctl
End Function

Private Sub DateTimeRange_Button(ByVal formName As String, ByVal name As String, _
                                 ByVal caption As String, ByVal left As Long, ByVal expression As String)
    Dim ctl As Control
    Set ctl = CreateControl(formName, acCommandButton, acDetail, , , left, 5040, 900, 360)
    ctl.name = name
    ctl.caption = caption
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.OnClick = expression
End Sub

Private Sub DateTimeRange_Editors(ByVal formName As String, ByVal endpoint As Integer)
    Dim prefix As String, panelLeft As Long, ctl As Control, column As Integer
    prefix = "Start"
    If endpoint = 1 Then
        prefix = "End"
    End If
    panelLeft = 180 + endpoint * 5040
    Set ctl = DateTimeRange_Label(formName, "lbl" & prefix, IIf(endpoint = 0, "开始", "结束"), panelLeft, 120, 600, 360)
    Set ctl = CreateControl(formName, acTextBox, acDetail, , , panelLeft + 600, 120, 1740, 360)
    ctl.name = "txt" & prefix & "Date"
    ctl.Format = "yyyy-mm-dd"
    ctl.ShowDatePicker = 0
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.BorderColor = CLR_BORDER
    ctl.AfterUpdate = "=DateTimeRangePicker_DateChanged(" & endpoint & ")"
    Dim names As Variant, item As Integer, values As String, lastItem As Integer
    names = Array("Hour", "Min", "Sec")
    For column = 0 To 2
        Set ctl = CreateControl(formName, acComboBox, acDetail, , , panelLeft + 2460 + column * 780, 120, 600, 360)
        ctl.name = "cbo" & prefix & names(column)
        lastItem = 59
        If column = 0 Then
            lastItem = 23
        End If
        values = ""
        For item = 0 To lastItem
            If item > 0 Then
                values = values & ";"
            End If
            values = values & Format$(item, "00")
        Next item
        ctl.RowSourceType = "Value List"
        ctl.RowSource = values
        ctl.ColumnCount = 1
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        ctl.AllowValueListEdits = False
        ctl.ListRows = 10
        ctl.FontName = "Microsoft YaHei"
        ctl.FontSize = 10
        ctl.BorderColor = CLR_BORDER
        ctl.AfterUpdate = "=DateTimeRangePicker_TimeChanged(" & endpoint & ")"
        ctl.ControlTipText = Choose(column + 1, "时 (00-23)", "分 (00-59)", "秒 (00-59)")
        If column < 2 Then
            Set ctl = DateTimeRange_Label(formName, "lblColon" & endpoint & column, ":", _
                                          panelLeft + 3060 + column * 780, 120, 180, 360)
        End If
    Next column
End Sub

'====================================================
' 部署 API: 创建或重建 frmDateTimeRangePicker
' 参数: silent=True 时不显示提示，错误交给调用方处理
' 返回: 创建成功为 True；重建会替换已有的时间范围窗体
'====================================================
Public Function CreateDateTimeRangePickerForm(Optional ByVal silent As Boolean = False) As Boolean
    On Error GoTo Failed
    Dim obj As AccessObject, exists As Boolean
    For Each obj In CurrentProject.AllForms
        If obj.name = RANGE_FORM Then
            exists = True
        End If
    Next obj
    If exists Then
        DateTimeRange_CloseExisting
        DoEvents
        DoCmd.DeleteObject acForm, RANGE_FORM
    End If
    Dim frm As Form, temporaryName As String, ctl As Control
    Set frm = CreateForm
    temporaryName = frm.name
    With frm
        .caption = "时间范围"
        .DefaultView = 0
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = True
        .PopUp = True
        .Modal = False
        .BorderStyle = 0
        .MinMaxButtons = 0
        .CloseButton = False
        .width = 10020
        .Section(acDetail).height = 5550
        .Section(acDetail).BackColor = CLR_WHITE
        .OnLoad = "=DateTimeRangePicker_FormLoad()"
        .OnClose = "=DateTimeRangePicker_FormClose()"
        .OnDeactivate = "=DateTimeRangePicker_Deactivate()"
        .OnTimer = "=DateTimeRangePicker_Timer()"
        .Section(acDetail).OnMouseMove = "=DateTimeRangePicker_DayMouseMove(0)"
    End With
    Set ctl = CreateControl(temporaryName, acRectangle, acDetail, , , 15, 15, 9990, 5520)
    ctl.BorderColor = CLR_BORDER
    ctl.BackStyle = 0
    DateTimeRange_Editors temporaryName, 0
    DateTimeRange_Editors temporaryName, 1
    Set ctl = DateTimeRange_Label(temporaryName, "lblDisplay", "", 180, 720, 9660, 420)
    ctl.ForeColor = CLR_BLUE
    Dim panel As Integer, column As Integer, cell As Integer, index As Integer, panelLeft As Long
    Dim weekdays As Variant
    weekdays = Array("日", "一", "二", "三", "四", "五", "六")
    For panel = 0 To 1
        panelLeft = 180 + panel * 5040
        Set ctl = DateTimeRange_Label(temporaryName, "lblMonth" & panel, "", panelLeft + 1200, 1260, 2280, 420)
        For column = 0 To 6
            Set ctl = DateTimeRange_Label(temporaryName, "lblW" & panel & column, CStr(weekdays(column)), _
                                          panelLeft + column * 660, 1800, 660, 360)
        Next column
        For cell = 1 To 42
            index = panel * 42 + cell
            Set ctl = DateTimeRange_Label(temporaryName, "lblD" & Format$(index, "00"), "", _
                                          panelLeft + ((cell - 1) Mod 7) * 660, _
                                          2220 + ((cell - 1) \ 7) * 440, 660, 420)
            ctl.OnClick = "=DateTimeRangePicker_DayClick(" & index & ")"
            ctl.OnMouseMove = "=DateTimeRangePicker_DayMouseMove(" & index & ")"
        Next cell
    Next panel
    Dim names As Variant, captions As Variant, offsets As Variant, months As Variant
    names = Array("cmdPrevYear", "cmdPrevMonth", "cmdNextMonth", "cmdNextYear")
    captions = Array("<<", "<", ">", ">>")
    offsets = Array(180, 720, 8760, 9300)
    months = Array(-12, -1, 1, 12)
    For column = 0 To 3
        Set ctl = CreateControl(temporaryName, acCommandButton, acDetail, , , offsets(column), 1260, 540, 360)
        ctl.name = names(column)
        ctl.caption = captions(column)
        ctl.OnClick = "=DateTimeRangePicker_Navigate(" & months(column) & ")"
        ctl.ControlTipText = Choose(column + 1, "上一年", "上个月", "下个月", "下一年")
    Next column
    DateTimeRange_Button temporaryName, "cmdClear", "清除", 180, "=DateTimeRangePicker_Clear()"
    DateTimeRange_Button temporaryName, "cmdToday", "今天", 1260, "=DateTimeRangePicker_Today()"
    DateTimeRange_Button temporaryName, "cmdOK", "确定", 7680, "=DateTimeRangePicker_OK()"
    DateTimeRange_Button temporaryName, "cmdCancel", "取消", 8940, "=DateTimeRangePicker_Cancel()"
    frm.Controls("cmdOK").Default = True
    frm.Controls("cmdCancel").Cancel = True
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename RANGE_FORM, acForm, temporaryName
    CreateDateTimeRangePickerForm = True
    If Not silent Then
        MsgBox "时间范围窗体已创建。调用 ShowDateTimeRangePicker() 或 PickDateTimeRangeFor Me.txtRange。", vbInformation, "创建成功"
    End If
    Exit Function
Failed:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If temporaryName <> "" Then
        DoCmd.Close acForm, temporaryName, acSaveNo
    End If
    On Error GoTo 0
    If silent Then
        Err.Raise errorNumber, "CreateDateTimeRangePickerForm", errorText
    Else
        MsgBox "创建时间范围窗体失败: " & errorText, vbCritical, "创建失败"
    End If
End Function
