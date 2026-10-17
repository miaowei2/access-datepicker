Option Compare Database
Option Explicit

'====================================================
' 模块名称: Module_DateRangePicker
' 功能描述: Access 纯窗体日期范围选择器，双月日历选择起止日期
'
' 特点:
'   - 第一次点击选择起始日期，第二次点击选择结束日期并确认
'   - 支持反向选择、跨月选择与悬停预览
'   - 支持合并文本框或起止两个文本框，无需窗体代码模块
'
' 使用方法:
'   1. 将本模块以 GBK 编码导入 Access 数据库
'   2. 在立即窗口运行: CreateDateRangePickerForm
'   3. 在代码中调用:
'      result = ShowDateRangePicker()            ' 返回 Array(Date, Date) 或 Null
'      PickDateRangeFor Me.txtRange              ' 回写一个合并文本框
'      PickDateRangeFor Me.txtStart, Me.txtEnd    ' 分别回写起止日期
'====================================================

' ============ Win32 API 声明 (定位下拉窗体，兼容 32/64 位 Access) ============
#If VBA7 Then
    Private Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr
    Private Declare PtrSafe Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function apiMonitorFromWindow Lib "user32" Alias "MonitorFromWindow" (ByVal hwnd As LongPtr, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As LongPtr, lpMonitorInfo As MONITORINFO) As Long
    Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#Else
    Private Declare Function apiGetFocus Lib "user32" Alias "GetFocus" () As Long
    Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function apiMonitorFromWindow Lib "user32" Alias "MonitorFromWindow" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
    Private Declare Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, lpMonitorInfo As MONITORINFO) As Long
    Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#End If

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

' ============ 窗体名、范围分隔符与颜色常量 ============
Private Const RANGE_FORM As String = "frmDateRangePicker"
Private Const RANGE_SEPARATOR As String = " ~ "
Private Const CLR_BLUE As Long = 16024898
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_BLACK As Long = 0
Private Const CLR_GRAY As Long = 10526880
Private Const CLR_BORDER As Long = 13158600
Private Const CLR_RANGE As Long = 16576233

' ============ 模块级变量 ============
Private m_viewMonth As Date              ' 左侧显示月份的第 1 天
Private m_startDate As Variant           ' 起始日期，或 Null
Private m_endDate As Variant             ' 结束日期，或 Null
Private m_result As Variant              ' 模态结果: 日期数组或 Null
Private m_waitingEnd As Boolean          ' 已选起点，正在等待第二次点击
Private m_nonModal As Boolean            ' 是否为文本框下拉模式
Private m_hoverCell As Integer           ' 当前悬停单元格，0 表示无
Private m_cellDates(1 To 84) As Variant   ' 两个月各 42 格对应的日期
Private m_startInput As Control          ' 合并文本框，或起始日期文本框
Private m_endInput As Control            ' 结束日期文本框；合并模式为 Nothing
Private m_outputFormat As String         ' 两框模式为空时写 Date，否则写字符串

'====================================================
' 公共 API: 以对话框方式选择日期范围
' 参数: dtStart、dtEnd 为可选默认起止日期，时间部分会被去掉
' 返回: 确定时为 Array(起始 Date, 结束 Date)，按先后排序
'       取消或清除时为 Null
' 示例: result = ShowDateRangePicker(#2026-10-1#, #2026-10-7#)
'====================================================
Public Function ShowDateRangePicker(Optional ByVal dtStart As Variant, _
                                    Optional ByVal dtEnd As Variant) As Variant
    On Error GoTo Failed
    ShowDateRangePicker = Null
    If Not RangePicker_Ready() Then
        Exit Function
    End If
    RangePicker_CloseExisting
    RangePicker_Initialize dtStart, dtEnd
    m_nonModal = False
    DoCmd.OpenForm RANGE_FORM, acNormal, , , , acDialog
    ShowDateRangePicker = m_result
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "日期范围选择器"
End Function

'====================================================
' 通用 API: 为一个或两个文本框弹出范围选择器并回写
' 参数: ctlStart 为合并文本框或起始框；ctlEnd 为可选结束框
'       sFormat 为可选格式；两框模式默认写 Date，单框模式默认
'       写 "yyyy-mm-dd ~ yyyy-mm-dd" 文本
' 行为: 确认才回写，取消不改原值，清除将目标框设为 Null
' 示例: PickDateRangeFor Me.txtStart, Me.txtEnd, "yyyy-mm-dd"
'====================================================
Public Sub PickDateRangeFor(ctlStart As Control, _
                            Optional ctlEnd As Control, _
                            Optional ByVal sFormat As String = "")
    On Error GoTo Failed
    If Not RangePicker_Ready() Then
        Exit Sub
    End If
    If Not TypeOf ctlStart Is TextBox Then
        Err.Raise 5, , "开始控件必须是文本框"
    End If
    If Not ctlEnd Is Nothing Then
        If Not TypeOf ctlEnd Is TextBox Then
            Err.Raise 5, , "结束控件必须是文本框"
        End If
        If ctlStart Is ctlEnd Then
            Err.Raise 5, , "起止日期须使用不同的文本框"
        End If
    End If

    RangePicker_CloseExisting
    Dim startValue As Variant, endValue As Variant
    startValue = Null
    endValue = Null
    If ctlEnd Is Nothing Then
        ' 合并模式用相同分隔符读取已有范围，无法解析时保留默认状态。
        Dim parts As Variant
        parts = Split(Nz(ctlStart.Value, ""), RANGE_SEPARATOR)
        If UBound(parts) = 1 Then
            startValue = parts(0)
            endValue = parts(1)
        End If
    Else
        startValue = ctlStart.Value
        endValue = ctlEnd.Value
    End If
    RangePicker_Initialize startValue, endValue
    Set m_startInput = ctlStart
    Set m_endInput = ctlEnd
    m_outputFormat = sFormat
    m_nonModal = True

    Dim inputRect As RECT
    ctlStart.SetFocus
    apiGetWindowRect apiGetFocus(), inputRect
    DoCmd.OpenForm RANGE_FORM, acNormal
    RangePicker_Position inputRect
    Exit Sub
Failed:
    Set m_startInput = Nothing
    Set m_endInput = Nothing
    MsgBox Err.Description, vbExclamation, "日期范围选择器"
End Sub

' 表达式入口: =PickDateRangeForCtl("Demo","txtDateRange")。
' 两框模式可再传结束控件名和格式，适用于单击或双击事件表达式。
Public Function PickDateRangeForCtl(ByVal sFormName As String, _
                                    ByVal sStartCtlName As String, _
                                    Optional ByVal sEndCtlName As String = "", _
                                    Optional ByVal sFormat As String = "") As Boolean
    On Error GoTo Failed
    Dim endControl As Control
    If sEndCtlName <> "" Then
        Set endControl = Forms(sFormName).Controls(sEndCtlName)
    End If
    PickDateRangeFor Forms(sFormName).Controls(sStartCtlName), endControl, sFormat
    PickDateRangeForCtl = True
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "日期范围选择器"
End Function

' 自动绑定合并文本框或起止两个文本框的双击事件，使用同一选择器。
Public Sub AttachDateRangePicker(frm As Form, ByVal sStartCtlName As String, _
                                 Optional ByVal sEndCtlName As String = "", _
                                 Optional ByVal sFormat As String = "")
    Dim expression As String
    expression = "=PickDateRangeForCtl(""" & Replace(frm.Name, """", """""") & """,""" & _
                 Replace(sStartCtlName, """", """""") & """,""" & _
                 Replace(sEndCtlName, """", """""") & """,""" & _
                 Replace(sFormat, """", """""") & """)"
    frm.Controls(sStartCtlName).OnDblClick = expression
    If sEndCtlName <> "" Then
        frm.Controls(sEndCtlName).OnDblClick = expression
    End If
End Sub

Private Function RangePicker_DateOnly(ByVal value As Variant) As Variant
    RangePicker_DateOnly = Null
    If IsDate(value) Then
        RangePicker_DateOnly = DateValue(CDate(value))
    End If
End Function

Private Sub RangePicker_Initialize(ByVal dtStart As Variant, ByVal dtEnd As Variant)
    m_startDate = RangePicker_DateOnly(dtStart)
    m_endDate = RangePicker_DateOnly(dtEnd)
    If IsNull(m_startDate) Then
        m_startDate = m_endDate
        m_endDate = Null
    End If
    If Not IsNull(m_endDate) Then
        RangePicker_Sort m_startDate, m_endDate
    End If
    Dim baseDate As Date
    baseDate = Date
    If Not IsNull(m_startDate) Then
        baseDate = m_startDate
    End If
    m_viewMonth = DateSerial(Year(baseDate), Month(baseDate), 1)
    ' 右侧始终显示下个月，因此左侧最大只能到 9999 年 11 月。
    If m_viewMonth > DateSerial(9999, 11, 1) Then
        m_viewMonth = DateSerial(9999, 11, 1)
    End If
    m_waitingEnd = False
    m_hoverCell = 0
    m_result = Null
    m_outputFormat = ""
    Set m_startInput = Nothing
    Set m_endInput = Nothing
End Sub

Private Sub RangePicker_Sort(ByRef startValue As Variant, ByRef endValue As Variant)
    If startValue > endValue Then
        Dim temporary As Variant
        temporary = startValue
        startValue = endValue
        endValue = temporary
    End If
End Sub

' ============ 窗体与控件事件 (由表达式调用) ============
Public Function DateRangePicker_FormLoad() As Boolean
    If m_viewMonth = 0 Then
        RangePicker_Initialize Null, Null
    End If
    RangePicker_Refresh
    DateRangePicker_FormLoad = True
End Function

Public Function DateRangePicker_DayClick(ByVal cell As Integer) As Boolean
    If cell < 1 Or cell > 84 Then
        Exit Function
    End If
    If IsNull(m_cellDates(cell)) Then
        Exit Function
    End If
    If Not m_waitingEnd Then
        ' 第一次点击只设置起点，并进入等待终点状态。
        m_startDate = m_cellDates(cell)
        m_endDate = Null
        m_waitingEnd = True
        m_hoverCell = 0
        RangePicker_Refresh
    Else
        ' 第二次点击确定终点，自动交换反向范围后提交。
        m_endDate = m_cellDates(cell)
        RangePicker_Sort m_startDate, m_endDate
        m_waitingEnd = False
        RangePicker_Refresh
        DateRangePicker_OK
    End If
    DateRangePicker_DayClick = True
End Function

Public Function DateRangePicker_DayMouseMove(ByVal cell As Integer) As Boolean
    If cell < 0 Or cell > 84 Then
        Exit Function
    End If
    If cell <> m_hoverCell Then
        m_hoverCell = cell
        RangePicker_Refresh
    End If
    DateRangePicker_DayMouseMove = True
End Function

Public Function DateRangePicker_Navigate(ByVal months As Integer) As Boolean
    Dim monthNumber As Long
    monthNumber = CLng(Year(m_viewMonth)) * 12 + Month(m_viewMonth) - 1 + CLng(months)
    If monthNumber < 1200 Or monthNumber > 119998 Then
        Exit Function
    End If
    m_viewMonth = DateSerial(monthNumber \ 12, monthNumber Mod 12 + 1, 1)
    m_hoverCell = 0
    RangePicker_Refresh
    DateRangePicker_Navigate = True
End Function

' “今天”仅将日历导航回当前月，不改变已选择的起止日期。
Public Function DateRangePicker_Today() As Boolean
    m_viewMonth = DateSerial(Year(Date), Month(Date), 1)
    m_hoverCell = 0
    RangePicker_Refresh
    DateRangePicker_Today = True
End Function

Public Function DateRangePicker_OK() As Boolean
    If IsNull(m_startDate) Or IsNull(m_endDate) Or m_waitingEnd Then
        Exit Function
    End If
    If Not RangePicker_WriteBack(False) Then
        Exit Function
    End If
    m_result = Array(CDate(m_startDate), CDate(m_endDate))
    DoCmd.Close acForm, RANGE_FORM, acSaveNo
    DateRangePicker_OK = True
End Function

Public Function DateRangePicker_Clear() As Boolean
    If Not RangePicker_WriteBack(True) Then
        Exit Function
    End If
    m_result = Null
    DoCmd.Close acForm, RANGE_FORM, acSaveNo
    DateRangePicker_Clear = True
End Function

Public Function DateRangePicker_Cancel() As Boolean
    m_result = Null
    DoCmd.Close acForm, RANGE_FORM, acSaveNo
    DateRangePicker_Cancel = True
End Function

Public Function DateRangePicker_FormClose() As Boolean
    Set m_startInput = Nothing
    Set m_endInput = Nothing
    m_viewMonth = 0
    m_nonModal = False
    DateRangePicker_FormClose = True
End Function

Public Function DateRangePicker_Deactivate() As Boolean
    If m_nonModal Then
        ' 延迟到焦点切换完成后再取消，避免把正常窗体切换误判为外部点击。
        Forms(RANGE_FORM).TimerInterval = 100
    End If
    DateRangePicker_Deactivate = True
End Function

Public Function DateRangePicker_Timer() As Boolean
    Forms(RANGE_FORM).TimerInterval = 0
    If m_nonModal Then
        If Screen.ActiveForm.Name <> RANGE_FORM Then
            DateRangePicker_Cancel
        End If
    End If
    DateRangePicker_Timer = True
End Function

' ============ 双月日历绘制与结果回写 ============
Private Sub RangePicker_Refresh()
    Dim frm As Form
    Set frm = Forms(RANGE_FORM)
    Dim rangeStart As Variant, rangeEnd As Variant
    rangeStart = m_startDate
    rangeEnd = m_endDate
    If m_waitingEnd And m_hoverCell > 0 Then
        ' 悬停只计算临时终点，不修改已保存的选择状态。
        rangeEnd = m_cellDates(m_hoverCell)
        If Not IsNull(rangeEnd) Then
            RangePicker_Sort rangeStart, rangeEnd
        End If
    End If

    Dim panel As Integer, cell As Integer, index As Integer
    Dim firstDate As Date, gridStart As Double, serial As Double
    Dim cellDate As Date, ctl As Control
    For panel = 0 To 1
        firstDate = DateAdd("m", panel, m_viewMonth)
        frm.Controls("lblMonth" & panel).Caption = Format(firstDate, "yyyy 年 m 月")
        gridStart = CDbl(firstDate) - Weekday(firstDate, vbSunday) + 1
        For cell = 1 To 42
            index = panel * 42 + cell
            Set ctl = frm.Controls("lblD" & Format(index, "00"))
            serial = gridStart + cell - 1
            m_cellDates(index) = Null
            ctl.Caption = ""
            ctl.BackColor = CLR_WHITE
            ctl.ForeColor = CLR_BLACK
            ctl.FontWeight = 400
            If serial >= CDbl(DateSerial(100, 1, 1)) And serial <= CDbl(DateSerial(9999, 12, 31)) Then
                cellDate = CDate(serial)
                m_cellDates(index) = cellDate
                ctl.Caption = CStr(Day(cellDate))
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
    Dim display As String
    display = "选开始日期" & RANGE_SEPARATOR & "选结束日期"
    If Not IsNull(m_startDate) Then
        display = Format(m_startDate, "yyyy-mm-dd") & RANGE_SEPARATOR & "选结束日期"
        If Not IsNull(m_endDate) Then
            display = Format(m_startDate, "yyyy-mm-dd") & RANGE_SEPARATOR & Format(m_endDate, "yyyy-mm-dd")
        End If
    End If
    frm.Controls("lblDisplay").Caption = display
    frm.Controls("cmdOK").Enabled = (Not IsNull(m_startDate) And Not IsNull(m_endDate) And Not m_waitingEnd)
End Sub

Private Function RangePicker_WriteBack(ByVal clearValues As Boolean) As Boolean
    On Error GoTo Failed
    RangePicker_WriteBack = True
    If Not m_nonModal Or m_startInput Is Nothing Then
        Exit Function
    End If
    Dim oldStart As Variant, oldEnd As Variant
    ' 先备份两框原值；任一写入失败时恢复，避免只更新一半范围。
    oldStart = m_startInput.Value
    If Not m_endInput Is Nothing Then
        oldEnd = m_endInput.Value
    End If
    Dim startValue As Variant, endValue As Variant
    startValue = Null
    endValue = Null
    If Not clearValues Then
        If m_endInput Is Nothing Then
            Dim formatText As String
            formatText = m_outputFormat
            If formatText = "" Then
                formatText = "yyyy-mm-dd"
            End If
            startValue = Format(m_startDate, formatText) & RANGE_SEPARATOR & Format(m_endDate, formatText)
        ElseIf m_outputFormat <> "" Then
            startValue = Format(m_startDate, m_outputFormat)
            endValue = Format(m_endDate, m_outputFormat)
        Else
            startValue = CDate(m_startDate)
            endValue = CDate(m_endDate)
        End If
    End If
    m_startInput.Value = startValue
    If Not m_endInput Is Nothing Then
        m_endInput.Value = endValue
    End If
    Exit Function
Failed:
    Dim errorText As String
    errorText = Err.Description
    On Error Resume Next
    m_startInput.Value = oldStart
    If Not m_endInput Is Nothing Then
        m_endInput.Value = oldEnd
    End If
    RangePicker_WriteBack = False
    MsgBox "无法回写日期范围: " & errorText, vbExclamation, "日期范围选择器"
End Function

Private Function RangePicker_Ready() As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = RANGE_FORM Then
            RangePicker_Ready = True
            Exit Function
        End If
    Next obj
    MsgBox "请先运行 CreateDateRangePickerForm 创建日期范围窗体。", vbExclamation, "日期范围选择器"
End Function

Private Sub RangePicker_CloseExisting()
    If CurrentProject.AllForms(RANGE_FORM).IsLoaded Then
        DoCmd.Close acForm, RANGE_FORM, acSaveNo
    End If
End Sub

' 按显示器工作区限制弹窗位置，底部放不下时改放输入框上方。
Private Sub RangePicker_Position(ByRef inputRect As RECT)
    Dim windowRect As RECT, monitor As MONITORINFO
    Dim frm As Form
    Set frm = Forms(RANGE_FORM)
    apiGetWindowRect frm.hwnd, windowRect
    monitor.cbSize = LenB(monitor)
    If apiGetMonitorInfo(apiMonitorFromWindow(frm.hwnd, 2), monitor) = 0 Then
        Exit Sub
    End If
    Dim width As Long, height As Long, left As Long, top As Long
    width = windowRect.Right - windowRect.Left
    height = windowRect.Bottom - windowRect.Top
    left = inputRect.Left
    top = inputRect.Bottom
    If left + width > monitor.rcWork.Right Then
        left = monitor.rcWork.Right - width
    End If
    If left < monitor.rcWork.Left Then
        left = monitor.rcWork.Left
    End If
    If top + height > monitor.rcWork.Bottom Then
        top = inputRect.Top - height
    End If
    If top < monitor.rcWork.Top Then
        top = monitor.rcWork.Top
    End If
    apiMoveWindow frm.hwnd, left, top, width, height, True
End Sub

' ============ 窗体生成辅助过程 ============
Private Function RangePicker_Label(ByVal formName As String, ByVal name As String, _
                                    ByVal caption As String, ByVal left As Long, _
                                    ByVal top As Long, ByVal width As Long, _
                                    ByVal height As Long) As Control
    Dim ctl As Control
    Set ctl = CreateControl(formName, acLabel, acDetail, , , left, top, width, height)
    ctl.Name = name
    ctl.Caption = caption
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.TextAlign = 2
    ctl.BackStyle = 1
    ctl.BackColor = CLR_WHITE
    ctl.ForeColor = CLR_BLACK
    ctl.BorderStyle = 0
    Set RangePicker_Label = ctl
End Function

Private Sub RangePicker_Button(ByVal formName As String, ByVal name As String, _
                                ByVal caption As String, ByVal left As Long, _
                                ByVal top As Long, ByVal width As Long, _
                                ByVal expression As String)
    Dim ctl As Control
    Set ctl = CreateControl(formName, acCommandButton, acDetail, , , left, top, width, 360)
    ctl.Name = name
    ctl.Caption = caption
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.OnClick = expression
End Sub

'====================================================
' 部署 API: 创建或重建 frmDateRangePicker 双月窗体
' 参数: silent=True 时不显示提示，错误交给调用方处理
' 返回: 创建成功为 True；重建会替换已有的选择器窗体
' 示例: CreateDateRangePickerForm
'====================================================
Public Function CreateDateRangePickerForm(Optional ByVal silent As Boolean = False) As Boolean
    On Error GoTo Failed
    Dim obj As AccessObject, exists As Boolean
    For Each obj In CurrentProject.AllForms
        If obj.Name = RANGE_FORM Then
            exists = True
        End If
    Next obj
    If exists Then
        RangePicker_CloseExisting
        ' 关闭后让 Access 释放窗体，再删除旧对象。
        DoEvents
        DoCmd.DeleteObject acForm, RANGE_FORM
    End If
    Dim frm As Form, temporaryName As String
    Set frm = CreateForm
    temporaryName = frm.Name
    With frm
        .Caption = "日期范围"
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
        .Width = 10020
        .Section(acDetail).Height = 4950
        .Section(acDetail).BackColor = CLR_WHITE
        .OnLoad = "=DateRangePicker_FormLoad()"
        .OnClose = "=DateRangePicker_FormClose()"
        .OnDeactivate = "=DateRangePicker_Deactivate()"
        .OnTimer = "=DateRangePicker_Timer()"
        .Section(acDetail).OnMouseMove = "=DateRangePicker_DayMouseMove(0)"
    End With

    Dim ctl As Control
    Set ctl = CreateControl(temporaryName, acRectangle, acDetail, , , 15, 15, 9990, 4920)
    ctl.BorderColor = CLR_BORDER
    ctl.BackStyle = 0
    Set ctl = RangePicker_Label(temporaryName, "lblDisplay", "日期范围", 180, 120, 9660, 420)
    ctl.ForeColor = CLR_BLUE
    Dim panel As Integer, column As Integer, cell As Integer, index As Integer, panelLeft As Long
    Dim weekdays As Variant
    weekdays = Array("日", "一", "二", "三", "四", "五", "六")
    For panel = 0 To 1
        panelLeft = 180 + panel * 5040
        Set ctl = RangePicker_Label(temporaryName, "lblMonth" & panel, "", panelLeft + 1200, 660, 2280, 420)
        For column = 0 To 6
            Set ctl = RangePicker_Label(temporaryName, "lblW" & panel & column, CStr(weekdays(column)), _
                                        panelLeft + column * 660, 1200, 660, 360)
        Next column
        For cell = 1 To 42
            index = panel * 42 + cell
            Set ctl = RangePicker_Label(temporaryName, "lblD" & Format(index, "00"), "", _
                                        panelLeft + ((cell - 1) Mod 7) * 660, _
                                        1620 + ((cell - 1) \ 7) * 440, 660, 420)
            ctl.OnClick = "=DateRangePicker_DayClick(" & index & ")"
            ctl.OnMouseMove = "=DateRangePicker_DayMouseMove(" & index & ")"
        Next cell
    Next panel
    RangePicker_Button temporaryName, "cmdPrevYear", "<<", 180, 660, 540, "=DateRangePicker_Navigate(-12)"
    RangePicker_Button temporaryName, "cmdPrevMonth", "<", 720, 660, 540, "=DateRangePicker_Navigate(-1)"
    RangePicker_Button temporaryName, "cmdNextMonth", ">", 8760, 660, 540, "=DateRangePicker_Navigate(1)"
    RangePicker_Button temporaryName, "cmdNextYear", ">>", 9300, 660, 540, "=DateRangePicker_Navigate(12)"
    RangePicker_Button temporaryName, "cmdClear", "清除", 180, 4440, 900, "=DateRangePicker_Clear()"
    RangePicker_Button temporaryName, "cmdToday", "今天", 1260, 4440, 900, "=DateRangePicker_Today()"
    RangePicker_Button temporaryName, "cmdOK", "确定", 7680, 4440, 900, "=DateRangePicker_OK()"
    RangePicker_Button temporaryName, "cmdCancel", "取消", 8940, 4440, 900, "=DateRangePicker_Cancel()"
    frm.Controls("cmdOK").Default = True
    frm.Controls("cmdCancel").Cancel = True
    frm.Controls("cmdPrevYear").ControlTipText = "上一年"
    frm.Controls("cmdPrevMonth").ControlTipText = "上个月"
    frm.Controls("cmdNextMonth").ControlTipText = "下个月"
    frm.Controls("cmdNextYear").ControlTipText = "下一年"
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename RANGE_FORM, acForm, temporaryName
    CreateDateRangePickerForm = True
    If Not silent Then
        MsgBox "日期范围窗体已创建。调用 ShowDateRangePicker() 或 PickDateRangeFor Me.txtRange。", vbInformation, "创建成功"
    End If
    Exit Function
Failed:
    Dim errorNumber As Long, errorText As String
    ' 清理操作可能覆盖 Err，先保留原始错误。
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If temporaryName <> "" Then
        DoCmd.Close acForm, temporaryName, acSaveNo
    End If
    On Error GoTo 0
    If silent Then
        Err.Raise errorNumber, "CreateDateRangePickerForm", errorText
    Else
        MsgBox "创建日期范围窗体失败: " & errorText, vbCritical, "创建失败"
    End If
End Function