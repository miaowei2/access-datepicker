Option Compare Database
Option Explicit

'====================================================
' 模块名称: Module_WeekPicker
' 功能描述: Access 纯窗体周选择器，按整周选择日期
'
' 特点:
'   - 每周从周日到周六；含 1 月 1 日的周为新年第 1 周，非 ISO 周
'   - 支持周标签或周日起始日期回写，无需窗体代码模块
'   - 点击日期选中整周并确认，支持清除与取消
'
' 使用方法:
'   1. 将本模块以 GBK 编码导入 Access 数据库
'   2. 在立即窗口运行: CreateWeekPickerForm
'   3. 在代码中调用:
'      dt = ShowWeekPicker()           ' 返回所选周的周日 Date 或 Null
'      PickWeekFor Me.txtWeek          ' 默认写入 yyyy-ww周 文本
'      PickWeekFor Me.txtWeek, True    ' 写入该周周日的 Date
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

' ============ 窗体名与颜色常量 ============
Private Const WEEK_FORM As String = "frmWeekPicker"
Private Const CLR_BLUE As Long = 16024898
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_BLACK As Long = 0
Private Const CLR_GRAY As Long = 10526880
Private Const CLR_BORDER As Long = 13158600
Private Const CLR_WEEK As Long = 16576233

' ============ 模块级变量 ============
Private m_viewMonth As Date             ' 当前显示月份的第 1 天
Private m_selected As Variant           ' 所选周的周日 Date，或 Null
Private m_result As Variant             ' 模态结果: 周日 Date 或 Null
Private m_hoverRow As Integer           ' 当前悬停的周行，-1 表示无
Private m_nonModal As Boolean           ' 是否为文本框下拉模式
Private m_storeDate As Boolean          ' True 写周日 Date，False 写周标签
Private m_input As Control              ' 需要回写的文本框
Private m_cellDates(1 To 42) As Variant  ' 6 行 7 列日历对应的日期

'====================================================
' 日期 API: 获取某周的周日 (周起始日)
' 参数: value 可为日期或 yyyy-ww周 标签，如 "2026-40周"
' 返回: 纯日期 Date；标签无效或整周越出日期范围时返回 Null
' 示例: dt = WeekPicker_WeekStart(#2026-10-1#)
'====================================================
Public Function WeekPicker_WeekStart(ByVal value As Variant) As Variant
    On Error GoTo InvalidValue
    WeekPicker_WeekStart = Null
    Dim serial As Double
    If IsDate(value) Then
        Dim dateValue As Date
        dateValue = VBA.DateValue(CDate(value))
        serial = CDbl(dateValue) - Weekday(dateValue, vbSunday) + 1
    Else
        Dim text As String, weekYear As Long, weekNumber As Integer
        text = Trim$(Nz(value, ""))
        If Not text Like "####-##周" Then
            Exit Function
        End If
        weekYear = CLng(Left$(text, 4))
        weekNumber = CInt(Mid$(text, 6, 2))
        If weekYear < 100 Or weekYear > 9999 Or weekNumber < 1 Or weekNumber > 53 Then
            Exit Function
        End If
        dateValue = DateSerial(weekYear, 1, 1)
        serial = CDbl(dateValue) - Weekday(dateValue, vbSunday) + 1 + (weekNumber - 1) * 7
    End If
    If serial < CDbl(DateSerial(100, 1, 1)) Or serial + 6 > CDbl(DateSerial(9999, 12, 31)) Then
        Exit Function
    End If
    If text <> "" Then
        ' 反向格式化校验，拒绝并不存在的第 53 周等标签。
        If WeekPicker_TextFromStart(CDate(serial)) <> text Then
            Exit Function
        End If
    End If
    WeekPicker_WeekStart = CDate(serial)
InvalidValue:
End Function

' 获取同一周的周六；参数与 WeekPicker_WeekStart 相同，无效值返回 Null。
Public Function WeekPicker_WeekEnd(ByVal value As Variant) As Variant
    Dim startValue As Variant
    startValue = WeekPicker_WeekStart(value)
    WeekPicker_WeekEnd = Null
    If Not IsNull(startValue) Then
        WeekPicker_WeekEnd = CDate(startValue + 6)
    End If
End Function

' 获取 yyyy-ww周 标签；参数可为日期或周标签，无效值返回空字符串。
Public Function WeekPicker_WeekText(ByVal value As Variant) As String
    Dim startValue As Variant
    startValue = WeekPicker_WeekStart(value)
    If Not IsNull(startValue) Then
        WeekPicker_WeekText = WeekPicker_TextFromStart(CDate(startValue))
    End If
End Function

Private Function WeekPicker_TextFromStart(ByVal startValue As Date) As String
    Dim weekYear As Long, firstDay As Date, firstSunday As Double, weekNumber As Integer
    ' 由周六确定周所属年份，因此跨年且包含元旦的周计入新年。
    weekYear = Year(startValue + 6)
    firstDay = DateSerial(weekYear, 1, 1)
    firstSunday = CDbl(firstDay) - Weekday(firstDay, vbSunday) + 1
    weekNumber = CInt((CDbl(startValue) - firstSunday) / 7) + 1
    WeekPicker_TextFromStart = Format$(weekYear, "0000") & "-" & Format$(weekNumber, "00") & "周"
End Function

'====================================================
' 公共 API: 以对话框方式选择周
' 参数: dtDefault 为可选日期或周标签，无效值默认当前周
' 返回: 确定时为该周周日的 Date；取消或清除时为 Null
' 示例: dt = ShowWeekPicker("2026-40周")
'====================================================
Public Function ShowWeekPicker(Optional ByVal dtDefault As Variant) As Variant
    On Error GoTo Failed
    ShowWeekPicker = Null
    If Not WeekPicker_Ready() Then
        Exit Function
    End If
    WeekPicker_CloseExisting
    WeekPicker_Initialize dtDefault
    m_nonModal = False
    DoCmd.OpenForm WEEK_FORM, acNormal, , , , acDialog
    ShowWeekPicker = m_result
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "周选择器"
End Function

'====================================================
' 通用 API: 为文本框弹出周选择器并回写结果
' 参数: ctl 为目标文本框；bStoreDate=True 写周日 Date，默认写周标签
' 返回: 成功打开为 True；取消不改原值，清除写入 Null
' 示例: PickWeekFor Me.txtWeek, True
'====================================================
Public Function PickWeekFor(ctl As Control, _
                             Optional ByVal bStoreDate As Boolean = False) As Boolean
    On Error GoTo Failed
    If Not WeekPicker_Ready() Then
        Exit Function
    End If
    If Not TypeOf ctl Is TextBox Then
        Err.Raise 5, , "请选择一个文本框"
    End If
    WeekPicker_CloseExisting
    WeekPicker_Initialize ctl.Value
    Set m_input = ctl
    m_storeDate = bStoreDate
    m_nonModal = True
    Dim inputRect As RECT
    ctl.SetFocus
    apiGetWindowRect apiGetFocus(), inputRect
    DoCmd.OpenForm WEEK_FORM, acNormal
    WeekPicker_Position inputRect
    PickWeekFor = True
    Exit Function
Failed:
    Set m_input = Nothing
    MsgBox Err.Description, vbExclamation, "周选择器"
End Function

' 表达式入口: =PickWeekForCtl("Demo","txtWeek")，可用于控件的单击事件。
Public Function PickWeekForCtl(ByVal sFormName As String, ByVal sCtlName As String, _
                                Optional ByVal bStoreDate As Boolean = False) As Boolean
    On Error GoTo Failed
    PickWeekForCtl = PickWeekFor(Forms(sFormName).Controls(sCtlName), bStoreDate)
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "周选择器"
End Function

' 自动绑定文本框的双击事件；bStoreDate 的含义与 PickWeekFor 相同。
Public Sub AttachWeekPicker(frm As Form, ByVal sCtlName As String, _
                             Optional ByVal bStoreDate As Boolean = False)
    frm.Controls(sCtlName).OnDblClick = "=PickWeekForCtl(""" & _
        Replace(frm.Name, """", """""") & """,""" & Replace(sCtlName, """", """""") & _
        """," & IIf(bStoreDate, "True", "False") & ")"
End Sub

Private Sub WeekPicker_Initialize(ByVal value As Variant)
    m_selected = WeekPicker_WeekStart(value)
    Dim baseDate As Date
    baseDate = Date
    If IsDate(value) Then
        baseDate = VBA.DateValue(CDate(value))
    ElseIf Not IsNull(m_selected) Then
        ' 标签初始化时以周中日期选显示月份，避免跨月周总跳到上个月。
        baseDate = m_selected + 3
    End If
    If IsNull(m_selected) Then
        m_selected = WeekPicker_WeekStart(baseDate)
    End If
    m_viewMonth = DateSerial(Year(baseDate), Month(baseDate), 1)
    m_result = Null
    m_hoverRow = -1
    Set m_input = Nothing
    m_storeDate = False
End Sub

' ============ 窗体与控件事件 (由表达式调用) ============
Public Function WeekPicker_FormLoad() As Boolean
    If m_viewMonth = 0 Then
        WeekPicker_Initialize Null
    End If
    WeekPicker_Refresh
    WeekPicker_FormLoad = True
End Function

Public Function WeekPicker_DayClick(ByVal cell As Integer) As Boolean
    If cell < 1 Or cell > 42 Then
        Exit Function
    End If
    If IsNull(m_cellDates(cell)) Then
        Exit Function
    End If
    m_selected = WeekPicker_WeekStart(m_cellDates(cell))
    If IsNull(m_selected) Then
        Exit Function
    End If
    WeekPicker_Refresh
    ' 点击任意一天即选择该天所在整周，并立即确认。
    WeekPicker_DayClick = WeekPicker_OK()
End Function

Public Function WeekPicker_DayMouseMove(ByVal cell As Integer) As Boolean
    If cell < 0 Or cell > 42 Then
        Exit Function
    End If
    Dim row As Integer
    row = -1
    If cell > 0 Then
        row = (cell - 1) \ 7
    End If
    If row <> m_hoverRow Then
        m_hoverRow = row
        WeekPicker_Refresh
    End If
    WeekPicker_DayMouseMove = True
End Function

Public Function WeekPicker_Navigate(ByVal months As Integer) As Boolean
    Dim monthNumber As Long
    monthNumber = CLng(Year(m_viewMonth)) * 12 + Month(m_viewMonth) - 1 + CLng(months)
    If monthNumber < 1200 Or monthNumber > 119999 Then
        Exit Function
    End If
    m_viewMonth = DateSerial(monthNumber \ 12, monthNumber Mod 12 + 1, 1)
    m_hoverRow = -1
    WeekPicker_Refresh
    WeekPicker_Navigate = True
End Function

' “本周”只更新预览，仍由确定按钮提交。
Public Function WeekPicker_ThisWeek() As Boolean
    m_selected = WeekPicker_WeekStart(Date)
    m_viewMonth = DateSerial(Year(Date), Month(Date), 1)
    m_hoverRow = -1
    WeekPicker_Refresh
    WeekPicker_ThisWeek = True
End Function

Public Function WeekPicker_OK() As Boolean
    If IsNull(m_selected) Then
        Exit Function
    End If
    If Not WeekPicker_WriteBack(False) Then
        Exit Function
    End If
    m_result = CDate(m_selected)
    DoCmd.Close acForm, WEEK_FORM, acSaveNo
    WeekPicker_OK = True
End Function

Public Function WeekPicker_Clear() As Boolean
    If Not WeekPicker_WriteBack(True) Then
        Exit Function
    End If
    m_result = Null
    DoCmd.Close acForm, WEEK_FORM, acSaveNo
    WeekPicker_Clear = True
End Function

Public Function WeekPicker_Cancel() As Boolean
    m_result = Null
    DoCmd.Close acForm, WEEK_FORM, acSaveNo
    WeekPicker_Cancel = True
End Function

Public Function WeekPicker_FormClose() As Boolean
    Set m_input = Nothing
    m_viewMonth = 0
    m_nonModal = False
    WeekPicker_FormClose = True
End Function

Public Function WeekPicker_Deactivate() As Boolean
    If m_nonModal Then
        ' 延迟检查焦点，避免窗体切换尚未完成时误判。
        Forms(WEEK_FORM).TimerInterval = 100
    End If
    WeekPicker_Deactivate = True
End Function

Public Function WeekPicker_Timer() As Boolean
    On Error GoTo Inactive
    Forms(WEEK_FORM).TimerInterval = 0
    If m_nonModal Then
        If Screen.ActiveForm.Name <> WEEK_FORM Then
            WeekPicker_Cancel
        End If
    End If
    WeekPicker_Timer = True
    Exit Function
Inactive:
    If m_nonModal Then
        WeekPicker_Cancel
    End If
End Function

' ============ 日历绘制与结果回写 ============
Private Sub WeekPicker_Refresh()
    Dim frm As Form, ctl As Control
    Set frm = Forms(WEEK_FORM)
    frm.Controls("lblMonth").Caption = Format$(m_viewMonth, "yyyy 年 m 月")
    frm.Controls("lblDisplay").Caption = WeekPicker_WeekText(m_selected)
    frm.Controls("cmdOK").Enabled = Not IsNull(m_selected)
    Dim gridStart As Double, serial As Double, cell As Integer, cellDate As Date
    ' 从月初所在周的周日起绘制 42 格，边界外日期保留为不可选空格。
    gridStart = CDbl(m_viewMonth) - Weekday(m_viewMonth, vbSunday) + 1
    For cell = 1 To 42
        Set ctl = frm.Controls("lblD" & Format$(cell, "00"))
        serial = gridStart + cell - 1
        m_cellDates(cell) = Null
        ctl.Caption = ""
        ctl.BackColor = CLR_WHITE
        ctl.ForeColor = CLR_BLACK
        ctl.FontWeight = 400
        If serial >= CDbl(DateSerial(100, 1, 1)) And serial <= CDbl(DateSerial(9999, 12, 31)) Then
            cellDate = CDate(serial)
            m_cellDates(cell) = cellDate
            ctl.Caption = CStr(Day(cellDate))
            If Month(cellDate) <> Month(m_viewMonth) Then
                ctl.ForeColor = CLR_GRAY
            End If
            If cellDate = Date Then
                ctl.ForeColor = CLR_BLUE
                ctl.FontWeight = 700
            End If
            If (cell - 1) \ 7 = m_hoverRow Then
                ctl.BackColor = CLR_WEEK
            End If
            If Not IsNull(m_selected) Then
                If cellDate >= m_selected And cellDate <= m_selected + 6 Then
                    ctl.BackColor = CLR_WEEK
                    If cellDate = m_selected Or cellDate = m_selected + 6 Then
                        ctl.BackColor = CLR_BLUE
                        ctl.ForeColor = CLR_WHITE
                        ctl.FontWeight = 700
                    End If
                End If
            End If
        End If
    Next cell
End Sub

Private Function WeekPicker_WriteBack(ByVal clearValue As Boolean) As Boolean
    On Error GoTo Failed
    WeekPicker_WriteBack = True
    If Not m_nonModal Or m_input Is Nothing Then
        Exit Function
    End If
    If clearValue Then
        m_input.Value = Null
    ElseIf m_storeDate Then
        m_input.Value = CDate(m_selected)
    Else
        m_input.Value = WeekPicker_WeekText(m_selected)
    End If
    Exit Function
Failed:
    WeekPicker_WriteBack = False
    MsgBox "无法回写周选择: " & Err.Description, vbExclamation, "周选择器"
End Function

Private Function WeekPicker_Ready() As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = WEEK_FORM Then
            WeekPicker_Ready = True
            Exit Function
        End If
    Next obj
    MsgBox "请先运行 CreateWeekPickerForm 创建周选择器窗体。", vbExclamation, "周选择器"
End Function

Private Sub WeekPicker_CloseExisting()
    If CurrentProject.AllForms(WEEK_FORM).IsLoaded Then
        DoCmd.Close acForm, WEEK_FORM, acSaveNo
    End If
End Sub

' 限定在输入框所在显示器的工作区内，底部空间不足时向上展开。
Private Sub WeekPicker_Position(ByRef inputRect As RECT)
    Dim windowRect As RECT, monitor As MONITORINFO, frm As Form
    Set frm = Forms(WEEK_FORM)
    apiGetWindowRect frm.hwnd, windowRect
    monitor.cbSize = LenB(monitor)
    If apiGetMonitorInfo(apiMonitorFromRect(inputRect, 2), monitor) = 0 Then
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
Private Function WeekPicker_Label(ByVal formName As String, ByVal name As String, _
                                   ByVal caption As String, ByVal left As Long, _
                                   ByVal top As Long, ByVal width As Long, ByVal height As Long) As Control
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
    Set WeekPicker_Label = ctl
End Function

Private Sub WeekPicker_Button(ByVal formName As String, ByVal name As String, _
                               ByVal caption As String, ByVal left As Long, ByVal top As Long, _
                               ByVal width As Long, ByVal expression As String)
    Dim ctl As Control
    Set ctl = CreateControl(formName, acCommandButton, acDetail, , , left, top, width, 360)
    ctl.Name = name
    ctl.Caption = caption
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.OnClick = expression
End Sub

'====================================================
' 部署 API: 创建或重建 frmWeekPicker
' 参数: silent=True 时不显示提示，错误交给调用方处理
' 返回: 创建成功为 True；重建会替换已有的选择器窗体
' 示例: CreateWeekPickerForm
'====================================================
Public Function CreateWeekPickerForm(Optional ByVal silent As Boolean = False) As Boolean
    On Error GoTo Failed
    Dim obj As AccessObject, exists As Boolean
    For Each obj In CurrentProject.AllForms
        If obj.Name = WEEK_FORM Then
            exists = True
        End If
    Next obj
    If exists Then
        WeekPicker_CloseExisting
        ' 关闭后让 Access 释放窗体，再删除旧对象。
        DoEvents
        DoCmd.DeleteObject acForm, WEEK_FORM
    End If
    Dim frm As Form, temporaryName As String, ctl As Control
    Set frm = CreateForm
    temporaryName = frm.Name
    With frm
        .Caption = "周选择器"
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
        .Width = 5100
        .Section(acDetail).Height = 4950
        .Section(acDetail).BackColor = CLR_WHITE
        .OnLoad = "=WeekPicker_FormLoad()"
        .OnClose = "=WeekPicker_FormClose()"
        .OnDeactivate = "=WeekPicker_Deactivate()"
        .OnTimer = "=WeekPicker_Timer()"
        .Section(acDetail).OnMouseMove = "=WeekPicker_DayMouseMove(0)"
    End With
    Set ctl = CreateControl(temporaryName, acRectangle, acDetail, , , 15, 15, 5070, 4920)
    ctl.BorderColor = CLR_BORDER
    ctl.BackStyle = 0
    Set ctl = WeekPicker_Label(temporaryName, "lblDisplay", "", 180, 120, 4740, 420)
    ctl.ForeColor = CLR_BLUE
    Set ctl = WeekPicker_Label(temporaryName, "lblMonth", "", 1260, 660, 2580, 420)
    Dim weekdays As Variant, column As Integer, cell As Integer
    weekdays = Array("日", "一", "二", "三", "四", "五", "六")
    For column = 0 To 6
        Set ctl = WeekPicker_Label(temporaryName, "lblW" & column, CStr(weekdays(column)), _
                                   240 + column * 660, 1200, 660, 360)
    Next column
    For cell = 1 To 42
        Set ctl = WeekPicker_Label(temporaryName, "lblD" & Format$(cell, "00"), "", _
                                   240 + ((cell - 1) Mod 7) * 660, _
                                   1620 + ((cell - 1) \ 7) * 440, 660, 420)
        ctl.OnClick = "=WeekPicker_DayClick(" & cell & ")"
        ctl.OnMouseMove = "=WeekPicker_DayMouseMove(" & cell & ")"
    Next cell
    WeekPicker_Button temporaryName, "cmdPrevYear", "<<", 180, 660, 540, "=WeekPicker_Navigate(-12)"
    WeekPicker_Button temporaryName, "cmdPrevMonth", "<", 720, 660, 540, "=WeekPicker_Navigate(-1)"
    WeekPicker_Button temporaryName, "cmdNextMonth", ">", 3840, 660, 540, "=WeekPicker_Navigate(1)"
    WeekPicker_Button temporaryName, "cmdNextYear", ">>", 4380, 660, 540, "=WeekPicker_Navigate(12)"
    WeekPicker_Button temporaryName, "cmdClear", "清除", 180, 4440, 1050, "=WeekPicker_Clear()"
    WeekPicker_Button temporaryName, "cmdThisWeek", "本周", 1410, 4440, 1050, "=WeekPicker_ThisWeek()"
    WeekPicker_Button temporaryName, "cmdOK", "确定", 2640, 4440, 1050, "=WeekPicker_OK()"
    WeekPicker_Button temporaryName, "cmdCancel", "取消", 3870, 4440, 1050, "=WeekPicker_Cancel()"
    frm.Controls("cmdOK").Default = True
    frm.Controls("cmdCancel").Cancel = True
    frm.Controls("cmdPrevYear").ControlTipText = "上一年"
    frm.Controls("cmdPrevMonth").ControlTipText = "上个月"
    frm.Controls("cmdNextMonth").ControlTipText = "下个月"
    frm.Controls("cmdNextYear").ControlTipText = "下一年"
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename WEEK_FORM, acForm, temporaryName
    CreateWeekPickerForm = True
    If Not silent Then
        MsgBox "周选择器已创建。调用 ShowWeekPicker() 或 PickWeekFor Me.txtWeek。", vbInformation, "创建成功"
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
        Err.Raise errorNumber, "CreateWeekPickerForm", errorText
    Else
        MsgBox "创建周选择器失败: " & errorText, vbCritical, "创建失败"
    End If
End Function