Option Compare Database
Option Explicit

'====================================================
' 模块名称: Module_TimePicker
' 功能描述: Access 纯窗体时间选择器，独立选择时、分、秒
'
' 特点:
'   - 三列列表选择时间，精确到秒，不保留日期部分
'   - 支持模态返回和文本框下拉回写，无需窗体代码模块
'   - 仅在确定时回写，取消或点击外部不改变原值
'
' 使用方法:
'   1. 将本模块以 GBK 编码导入 Access 数据库
'   2. 在立即窗口运行: CreateTimePickerForm
'   3. 在代码中调用:
'      dt = ShowTimePicker()              ' 返回纯时间 Date 或 Null
'      PickTimeFor Me.txtTime              ' 默认写回 Date
'      PickTimeFor Me.txtTime, "hh:nn:ss"   ' 按指定格式写回字符串
'====================================================

' ============ Win32 API 声明 (定位下拉窗体，兼容 32/64 位 Access) ============
#If VBA7 Then
    Private Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr
    Private Declare PtrSafe Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function apiMonitorFromRect Lib "user32" Alias "MonitorFromRect" (lpRect As RECT, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As LongPtr, lpMonitorInfo As MONITORINFO) As Long
    Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#Else
    Private Declare Function apiGetFocus Lib "user32" Alias "GetFocus" () As Long
    Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function apiMonitorFromRect Lib "user32" Alias "MonitorFromRect" (lpRect As RECT, ByVal dwFlags As Long) As Long
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

' ============ 窗体名与颜色常量 ============
Private Const TIME_FORM As String = "frmTimePicker"
Private Const CLR_BLUE As Long = 16024898
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_BORDER As Long = 13158600

' ============ 模块级变量 ============
Private m_selected As Date              ' 当前预览时间，不含日期部分
Private m_result As Variant             ' 模态结果: Date 或 Null
Private m_initialized As Boolean        ' 是否已由调用入口初始化
Private m_nonModal As Boolean           ' 是否为文本框下拉模式
Private m_input As Control              ' 需要回写的文本框
Private m_outputFormat As String        ' 空串写 Date，否则写格式化字符串

'====================================================
' 公共 API: 以对话框方式选择时间
' 参数: dtDefault 为可选默认日期/时间；无效值使用当前时间
' 返回: 确定时为纯时间 Date；取消或清除时为 Null
' 示例: dt = ShowTimePicker(#12:25:30 PM#)
'====================================================
Public Function ShowTimePicker(Optional ByVal dtDefault As Variant) As Variant
    On Error GoTo Failed
    ShowTimePicker = Null
    If Not TimePicker_Ready() Then
        Exit Function
    End If
    TimePicker_CloseExisting
    TimePicker_Initialize dtDefault
    DoCmd.OpenForm TIME_FORM, acNormal, , , , acDialog
    ShowTimePicker = m_result
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "时间选择器"
End Function

'====================================================
' 通用 API: 为文本框弹出下拉选择器并回写结果
' 参数: ctl 为目标文本框；sFormat 为空时写 Date，否则写字符串
' 返回: 成功打开为 True；取消不改原值，清除写入 Null
' 示例: PickTimeFor Me.txtTime, "hh:nn:ss"
'====================================================
Public Function PickTimeFor(ctl As Control, Optional ByVal sFormat As String = "") As Boolean
    On Error GoTo Failed
    If Not TimePicker_Ready() Then
        Exit Function
    End If
    If Not TypeOf ctl Is TextBox Then
        Err.Raise 5, , "时间控件必须是文本框"
    End If
    TimePicker_CloseExisting
    TimePicker_Initialize ctl.Value
    Set m_input = ctl
    m_outputFormat = sFormat
    m_nonModal = True
    ' 先聚焦输入框，再记录其屏幕坐标，避免取得弹窗自身的位置。
    Dim inputRect As RECT
    ctl.SetFocus
    apiGetWindowRect apiGetFocus(), inputRect
    DoCmd.OpenForm TIME_FORM, acNormal
    TimePicker_Position inputRect
    PickTimeFor = True
    Exit Function
Failed:
    Set m_input = Nothing
    m_nonModal = False
    MsgBox Err.Description, vbExclamation, "时间选择器"
End Function

' 表达式入口: =PickTimeForCtl("Demo","txtTime")，可用于控件的单击事件。
Public Function PickTimeForCtl(ByVal sFormName As String, ByVal sCtlName As String, _
                               Optional ByVal sFormat As String = "") As Boolean
    On Error GoTo Failed
    PickTimeForCtl = PickTimeFor(Forms(sFormName).Controls(sCtlName), sFormat)
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "时间选择器"
End Function

' 自动绑定文本框的双击事件；sFormat 的含义与 PickTimeFor 相同。
Public Sub AttachTimePicker(frm As Form, ByVal sCtlName As String, _
                            Optional ByVal sFormat As String = "")
    frm.Controls(sCtlName).OnDblClick = "=PickTimeForCtl(""" & _
        Replace(frm.Name, """", """""") & """,""" & _
        Replace(sCtlName, """", """""") & """,""" & _
        Replace(sFormat, """", """""") & """)"
End Sub

Private Function TimePicker_DefaultTime(ByVal value As Variant) As Date
    On Error GoTo UseCurrent
    If IsDate(value) Then
        ' 用 TimeSerial 去掉日期，保留时、分、秒；午夜仍是合法的 Date 值。
        TimePicker_DefaultTime = TimeSerial(Hour(CDate(value)), Minute(CDate(value)), Second(CDate(value)))
        Exit Function
    End If
UseCurrent:
    TimePicker_DefaultTime = Time
End Function

Private Sub TimePicker_Initialize(ByVal value As Variant)
    m_selected = TimePicker_DefaultTime(value)
    m_result = Null
    m_initialized = True
    m_nonModal = False
    m_outputFormat = ""
    Set m_input = Nothing
End Sub

' ============ 窗体与控件事件 (由表达式调用) ============
Public Function TimePicker_FormLoad() As Boolean
    If Not m_initialized Then
        TimePicker_Initialize Null
    End If
    TimePicker_Refresh
    TimePicker_FormLoad = True
End Function

Public Function TimePicker_TimeChanged() As Boolean
    Dim frm As Form
    Set frm = Forms(TIME_FORM)
    If IsNull(frm.Controls("lstHour").Value) Or IsNull(frm.Controls("lstMin").Value) Or _
       IsNull(frm.Controls("lstSec").Value) Then
        Exit Function
    End If
    ' 列表变化只刷新预览，目标文本框留到确定时再写入。
    m_selected = TimeSerial(CInt(frm.Controls("lstHour").Value), _
                           CInt(frm.Controls("lstMin").Value), CInt(frm.Controls("lstSec").Value))
    frm.Controls("lblDisplay").Caption = Format$(m_selected, "hh:nn:ss")
    TimePicker_TimeChanged = True
End Function

Public Function TimePicker_Now() As Boolean
    m_selected = Time
    TimePicker_Refresh
    TimePicker_Now = True
End Function

Public Function TimePicker_OK() As Boolean
    If Not TimePicker_WriteBack(False) Then
        Exit Function
    End If
    m_result = m_selected
    DoCmd.Close acForm, TIME_FORM, acSaveNo
    TimePicker_OK = True
End Function

Public Function TimePicker_Clear() As Boolean
    If Not TimePicker_WriteBack(True) Then
        Exit Function
    End If
    m_result = Null
    DoCmd.Close acForm, TIME_FORM, acSaveNo
    TimePicker_Clear = True
End Function

Public Function TimePicker_Cancel() As Boolean
    m_result = Null
    DoCmd.Close acForm, TIME_FORM, acSaveNo
    TimePicker_Cancel = True
End Function

Public Function TimePicker_FormClose() As Boolean
    Set m_input = Nothing
    m_initialized = False
    m_nonModal = False
    TimePicker_FormClose = True
End Function

Public Function TimePicker_Deactivate() As Boolean
    If m_nonModal Then
        ' 延迟检查焦点，让 Access 先完成窗体激活切换。
        Forms(TIME_FORM).TimerInterval = 100
    End If
    TimePicker_Deactivate = True
End Function

Public Function TimePicker_Timer() As Boolean
    On Error GoTo NoActiveForm
    Forms(TIME_FORM).TimerInterval = 0
    If m_nonModal Then
        If Screen.ActiveForm.Name <> TIME_FORM Then
            TimePicker_Cancel
        End If
    End If
    TimePicker_Timer = True
    Exit Function
NoActiveForm:
    If m_nonModal Then
        TimePicker_Cancel
    End If
End Function

' ============ 界面刷新与结果回写 ============
Private Sub TimePicker_Refresh()
    Dim frm As Form
    Set frm = Forms(TIME_FORM)
    frm.Controls("lblDisplay").Caption = Format$(m_selected, "hh:nn:ss")
    frm.Controls("lstHour").Value = Format$(Hour(m_selected), "00")
    frm.Controls("lstMin").Value = Format$(Minute(m_selected), "00")
    frm.Controls("lstSec").Value = Format$(Second(m_selected), "00")
End Sub

Private Function TimePicker_WriteBack(ByVal clearValue As Boolean) As Boolean
    On Error GoTo Failed
    If m_nonModal And Not m_input Is Nothing Then
        If clearValue Then
            m_input.Value = Null
        ElseIf m_outputFormat <> "" Then
            m_input.Value = Format$(m_selected, m_outputFormat)
        Else
            m_input.Value = m_selected
        End If
    End If
    TimePicker_WriteBack = True
    Exit Function
Failed:
    MsgBox "无法回写时间: " & Err.Description, vbExclamation, "时间选择器"
End Function

Private Function TimePicker_Ready() As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = TIME_FORM Then
            TimePicker_Ready = True
            Exit Function
        End If
    Next obj
    MsgBox "请先运行 CreateTimePickerForm 创建时间选择器窗体。", vbExclamation, "时间选择器"
End Function

Private Sub TimePicker_CloseExisting()
    If CurrentProject.AllForms(TIME_FORM).IsLoaded Then
        DoCmd.Close acForm, TIME_FORM, acSaveNo
    End If
End Sub

' 在输入框所在显示器的工作区内定位，底部空间不足时改放到输入框上方。
Private Sub TimePicker_Position(ByRef inputRect As RECT)
    Dim windowRect As RECT, monitor As MONITORINFO, frm As Form
    Set frm = Forms(TIME_FORM)
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

'====================================================
' 部署 API: 创建或重建 frmTimePicker
' 参数: silent=True 时不显示提示，错误交给调用方处理
' 返回: 创建成功为 True；重建会替换已有的选择器窗体
' 示例: CreateTimePickerForm
'====================================================
Public Function CreateTimePickerForm(Optional ByVal silent As Boolean = False) As Boolean
    On Error GoTo Failed
    Dim obj As AccessObject, exists As Boolean
    For Each obj In CurrentProject.AllForms
        If obj.Name = TIME_FORM Then
            exists = True
        End If
    Next obj
    If exists Then
        TimePicker_CloseExisting
        ' 关闭后让 Access 释放窗体，再删除旧对象。
        DoEvents
        DoCmd.DeleteObject acForm, TIME_FORM
    End If
    Dim frm As Form, temporaryName As String, ctl As Control
    Set frm = CreateForm
    temporaryName = frm.Name
    With frm
        .Caption = "时间选择"
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
        .Width = 3600
        .Section(acDetail).Height = 4200
        .Section(acDetail).BackColor = CLR_WHITE
        .OnLoad = "=TimePicker_FormLoad()"
        .OnClose = "=TimePicker_FormClose()"
        .OnDeactivate = "=TimePicker_Deactivate()"
        .OnTimer = "=TimePicker_Timer()"
    End With
    Set ctl = CreateControl(temporaryName, acRectangle, acDetail, , , 15, 15, 3570, 4170)
    ctl.BorderColor = CLR_BORDER
    ctl.BackStyle = 0
    Set ctl = CreateControl(temporaryName, acLabel, acDetail, , , 180, 120, 3240, 420)
    ctl.Name = "lblDisplay"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 12
    ctl.ForeColor = CLR_BLUE
    ctl.TextAlign = 2
    Dim column As Integer, item As Integer, names As Variant, headings As Variant
    Dim listValues As String, lastItem As Integer
    names = Array("lstHour", "lstMin", "lstSec")
    headings = Array("时", "分", "秒")
    For column = 0 To 2
        Set ctl = CreateControl(temporaryName, acLabel, acDetail, , , 180 + column * 1170, 660, 900, 360)
        ctl.Name = "lblUnit" & column
        ctl.Caption = headings(column)
        ctl.FontName = "Microsoft YaHei"
        ctl.FontSize = 10
        ctl.TextAlign = 2
        lastItem = 59
        If column = 0 Then
            lastItem = 23
        End If
        listValues = ""
        For item = 0 To lastItem
            If item > 0 Then
                listValues = listValues & ";"
            End If
            listValues = listValues & Format$(item, "00")
        Next item
        Set ctl = CreateControl(temporaryName, acListBox, acDetail, , , 180 + column * 1170, 1080, 900, 2340)
        ctl.Name = names(column)
        ctl.RowSourceType = "Value List"
        ctl.RowSource = listValues
        ctl.ColumnCount = 1
        ctl.BoundColumn = 1
        ctl.MultiSelect = 0
        ctl.AllowValueListEdits = False
        ctl.FontName = "Microsoft YaHei"
        ctl.FontSize = 11
        ctl.BorderStyle = 1
        ctl.BorderColor = CLR_BORDER
        ctl.AfterUpdate = "=TimePicker_TimeChanged()"
    Next column
    names = Array("cmdClear", "cmdNow", "cmdOK", "cmdCancel")
    headings = Array("清除", "现在", "确定", "取消")
    Dim handlers As Variant
    handlers = Array("TimePicker_Clear", "TimePicker_Now", "TimePicker_OK", "TimePicker_Cancel")
    For column = 0 To 3
        Set ctl = CreateControl(temporaryName, acCommandButton, acDetail, , , 180 + column * 840, 3660, 720, 360)
        ctl.Name = names(column)
        ctl.Caption = headings(column)
        ctl.FontName = "Microsoft YaHei"
        ctl.FontSize = 9
        ctl.OnClick = "=" & handlers(column) & "()"
    Next column
    frm.Controls("cmdOK").Default = True
    frm.Controls("cmdCancel").Cancel = True
    DoCmd.Close acForm, temporaryName, acSaveYes
    DoCmd.Rename TIME_FORM, acForm, temporaryName
    CreateTimePickerForm = True
    If Not silent Then
        MsgBox "时间选择器已创建。调用 ShowTimePicker() 或 PickTimeFor Me.txtTime。", vbInformation, "创建成功"
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
        Err.Raise errorNumber, "CreateTimePickerForm", errorText
    Else
        MsgBox "创建时间选择器失败: " & errorText, vbCritical, "创建失败"
    End If
End Function