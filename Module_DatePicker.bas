Attribute VB_Name = "Module_DatePicker"

Option Compare Database
Option Explicit

' ============ Win32 API 声明 (定位无边框窗体到文本框下方) ============
#If VBA7 Then
    Private Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As Long
    Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare PtrSafe Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare PtrSafe Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
#Else
    Private Declare Function apiGetFocus Lib "user32" Alias "GetFocus" () As Long
    Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
#End If

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'====================================================
' 模块名称: Module_DatePicker
' 功能描述: Access纯窗体日期时间选择器 (无需Web控件)
' 作者: edonsoft缪炜
' 版本: 1.0
'
' 特点:
'   - 纯VBA + Access窗体实现，零外部依赖
'   - 日历网格 + 时间列表 完整日期时间选择
'   - 自动创建窗体，一键部署
'   - 所有事件通过表达式绑定，无需窗体代码模块
'
' 使用方法:
'   1. 导入本模块到Access数据库
'   2. 在立即窗口运行: CreateDatePickerForm
'   3. 在代码中调用:
'      Dim dt As Variant
'      dt = ShowDatePicker()        ' 日期+时间
'      dt = ShowDatePicker(Now, False) ' 仅日期
'====================================================

' ============ 模块级变量 ============
Private m_iYear As Integer          ' 当前显示年份
Private m_iMonth As Integer         ' 当前显示月份
Private m_iDay As Integer           ' 已选择的日
Private m_iHour As Integer          ' 已选择的时
Private m_iMinute As Integer        ' 已选择的分
Private m_bShowTime As Boolean      ' 是否显示时间选择器

Private m_CellDates(1 To 42) As Date   ' 42个格子对应的日期

Private m_dtResult As Variant       ' 最终返回值
Private m_bConfirmed As Boolean     ' 用户是否确认

Public g_txtDateInput As TextBox    ' 调用时保存的文本框引用
Private m_rcInputBox As RECT        ' 调用时文本框的屏幕位置
Private m_sCallbackFmt As String    ' 回调: 日期格式
Private m_bIsNonModal As Boolean    ' 是否非模态 (下拉模式)

' ============ 颜色常量 ============
Private Const CLR_BLUE As Long = 16024898       ' RGB(66,133,244) 主题蓝
Private Const CLR_WHITE As Long = 16777215      ' RGB(255,255,255)
Private Const CLR_BLACK As Long = 0             ' RGB(0,0,0)
Private Const CLR_GRAY As Long = 11842740       ' RGB(180,180,180) 非当前月
Private Const CLR_DKGRAY As Long = 6579300      ' RGB(100,100,100) 星期头
Private Const CLR_BORDER As Long = 13158600     ' RGB(200,200,200) 输入框边框
Private Const CLR_HOVER As Long = 16576233      ' RGB(233,240,252) 鼠标悬停

Private m_iHoverCell As Integer                  ' 当前悬停的日期格子索引 (0=无)

' ============ 窗体名 ============
Private Const FORM_NAME As String = "frmDatePicker"


'====================================================
' ★ 公共API: 显示日期选择器
'
' 参数:
'   dtDefault  - 默认日期时间 (可选, 默认=Now)
'   bShowTime  - 是否显示时间选择 (可选, 默认=True)
'
' 返回:
'   Date   - 用户选择的日期(时间)
'   Null   - 用户取消或清除
'
' 示例:
'   Dim dt As Variant
'   dt = ShowDatePicker()
'   dt = ShowDatePicker(#2026-2-23#, False)
'   dt = ShowDatePicker(Me.txtDate.Value, True)
'====================================================
Public Function ShowDatePicker(Optional ByVal dtDefault As Variant, _
                               Optional ByVal bShowTime As Boolean = True) As Variant
    ' 检查窗体是否存在
    If Not FormExists(FORM_NAME) Then
        MsgBox "日期选择器窗体 [" & FORM_NAME & "] 不存在！" & vbCrLf & _
               "请先在立即窗口(Ctrl+G)中运行:" & vbCrLf & _
               "  CreateDatePickerForm", vbExclamation, "日期选择器"
        ShowDatePicker = Null
        Exit Function
    End If
    
    ' 初始化返回状态
    m_bConfirmed = False
    m_dtResult = Null
    
    ' 构建 OpenArgs 参数
    Dim sArgs As String
    If IsMissing(dtDefault) Or IsEmpty(dtDefault) Or IsNull(dtDefault) Then
        sArgs = Format(Now, "yyyy/mm/dd hh:nn:ss")
    Else
        On Error Resume Next
        sArgs = Format(CDate(dtDefault), "yyyy/mm/dd hh:nn:ss")
        If Err.Number <> 0 Then
            sArgs = Format(Now, "yyyy/mm/dd hh:nn:ss")
        End If
        On Error GoTo 0
    End If
    sArgs = sArgs & "|" & IIf(bShowTime, "1", "0")
    
    ' 保存文本框位置 (如果是由 PickDateFor 调用则已经设置)
    m_bIsNonModal = False
    
    ' 以对话框模式打开 (代码在此阻塞，直到窗体关闭)
    DoCmd.OpenForm FORM_NAME, acNormal, , , , acDialog, sArgs
    
    ' 窗体关闭后返回结果
    If m_bConfirmed Then
        ShowDatePicker = m_dtResult
    Else
        ShowDatePicker = Null
    End If
End Function


'====================================================
' ★ 通用API: 为任意控件弹出日期选择器并回写结果
'
' 用法1: 在窗体代码中 (推荐)
'   Private Sub txtDate_Click()
'       PickDateFor Me.txtDate
'   End Sub
'
' 用法2: 传入 bShowTime 参数
'   PickDateFor Me.txtDate, False    ' 仅日期
'   PickDateFor Me.txtDate, True     ' 日期+时间
'
' 用法3: 支持日期格式化输出到文本框
'   PickDateFor Me.txtDate, False, "yyyy-mm-dd"
'====================================================
Public Sub PickDateFor(ctl As Control, _
                       Optional ByVal bShowTime As Boolean = True, _
                       Optional ByVal sFormat As String = "")
On Error GoTo Err_PickDateFor
    
    ' 检查窗体是否存在
    If Not FormExists(FORM_NAME) Then
        MsgBox "日期选择器窗体 [" & FORM_NAME & "] 不存在！" & vbCrLf & _
               "请先在立即窗口(Ctrl+G)中运行:" & vbCrLf & _
               "  CreateDatePickerForm", vbExclamation, "日期选择器"
        Exit Sub
    End If
    
    ' 如果日期选择器已打开，先关闭
    If IsFormOpen(FORM_NAME) Then
        DoCmd.Close acForm, FORM_NAME
    End If
    
    ' 保存文本框引用 (参照 modCalendar 的 g_txtDateInput 方法)
    Set g_txtDateInput = ctl
    m_sCallbackFmt = sFormat
    m_bIsNonModal = True
    m_bConfirmed = False
    m_dtResult = Null
    
    ' ★ 关键: 先 SetFocus 给文本框，然后用 apiGetFocus 获取其窗口句柄
    ' 再用 GetWindowRect 获取文本框在屏幕上的精确位置
    ctl.SetFocus
    apiGetWindowRect apiGetFocus(), m_rcInputBox
    
    ' 构建 OpenArgs
    Dim dtDefault As Variant
    dtDefault = Null
    On Error Resume Next
    If Not IsNull(ctl.value) And ctl.value <> "" Then
        dtDefault = CDate(ctl.value)
        If Err.Number <> 0 Then dtDefault = Null
    End If
    On Error GoTo Err_PickDateFor
    
    Dim sArgs As String
    If IsNull(dtDefault) Then
        sArgs = Format(Now, "yyyy/mm/dd hh:nn:ss")
    Else
        sArgs = Format(dtDefault, "yyyy/mm/dd hh:nn:ss")
    End If
    sArgs = sArgs & "|" & IIf(bShowTime, "1", "0")
    
    ' 以非模态打开 (下拉模式)
    DoCmd.OpenForm FORM_NAME, acNormal, , , , , sArgs
    
    ' ★ 打开后立即定位 (与 modCalendar.CalendarFor 一致)
    Dim typRect_Calendar As RECT, typRect_Desktop As RECT
    Dim lngHwnd As Long
    Dim intLeft As Integer, intTop As Integer
    Dim intWidth As Integer, intHeight As Integer
    
    lngHwnd = Forms(FORM_NAME).hwnd
    apiGetWindowRect lngHwnd, typRect_Calendar
    apiGetWindowRect apiGetDesktopWindow(), typRect_Desktop
    
    intWidth = typRect_Calendar.Right - typRect_Calendar.Left
    intHeight = typRect_Calendar.Bottom - typRect_Calendar.Top
    
    ' 左对齐文本框左边缘 (-1 补偿 Access 无边框窗体客户区内边距)
    intLeft = m_rcInputBox.Left - 1
    If intLeft < 0 Then
        intLeft = 0
    ElseIf intLeft + intWidth > typRect_Desktop.Right Then
        intLeft = typRect_Desktop.Right - intWidth
    End If
    
    ' 紧贴文本框正下方
    intTop = m_rcInputBox.Bottom
    If intTop < 0 Then
        intTop = 0
    ElseIf intTop + intHeight > typRect_Desktop.Bottom Then
        intTop = m_rcInputBox.Top - intHeight
    End If
    
    apiMoveWindow lngHwnd, intLeft, intTop, intWidth, intHeight, True
    
Exit_PickDateFor:
    Exit Sub

Err_PickDateFor:
    MsgBox Err.Description, vbCritical, "日期选择器"
    Resume Exit_PickDateFor
End Sub


'====================================================
' ★ 通用API: 通过窗体名+控件名弹出日期选择器 (支持事件表达式)
'
' 在控件属性的"单击"事件中直接填写表达式:
'   =PickDateForCtl("frmOrder","txtOrderDate",True)
'   =PickDateForCtl("frmOrder","txtBirthday",False)
'
' 这样无需编写任何窗体代码！
'====================================================
Public Function PickDateForCtl(ByVal sFormName As String, _
                               ByVal sCtlName As String, _
                               Optional ByVal bShowTime As Boolean = True, _
                               Optional ByVal sFormat As String = "") As Variant
    On Error GoTo ErrHandler
    
    Dim frm As Form
    Set frm = Forms(sFormName)
    
    Dim ctl As Control
    Set ctl = frm.Controls(sCtlName)
    
    PickDateFor ctl, bShowTime, sFormat
    
    PickDateForCtl = True
    Exit Function

ErrHandler:
    MsgBox "PickDateForCtl 错误: " & Err.Description, vbExclamation, "日期选择器"
    PickDateForCtl = False
End Function


'====================================================
' ★ 通用API: 一键为文本框绑定日期选择器
'
' 在窗体的 Form_Load 中调用:
'   AttachDatePicker Me, "txtDate"            ' 日期+时间
'   AttachDatePicker Me, "txtBirthday", False ' 仅日期
'
' 效果: 单击或获得焦点时自动弹出日期选择器
'====================================================
Public Sub AttachDatePicker(frm As Form, _
                            ByVal sCtlName As String, _
                            Optional ByVal bShowTime As Boolean = True)
    On Error GoTo ErrHandler
    
    Dim ctl As Control
    Set ctl = frm.Controls(sCtlName)
    
    ' 用表达式绑定 OnDblClick 事件
    Dim sExpr As String
    sExpr = "=PickDateForCtl(""" & frm.Name & """,""" & sCtlName & """," & _
            IIf(bShowTime, "True", "False") & ")"
    
    ctl.OnDblClick = sExpr
    
    ' 把文本框设为只读，使交互更友好
    If TypeOf ctl Is TextBox Then
        ctl.Locked = False   ' 保持可编辑（用户也可以直接输入）
    End If
    
    Exit Sub

ErrHandler:
    Debug.Print "AttachDatePicker 错误 [" & sCtlName & "]: " & Err.Description
End Sub


'====================================================
' ★ 通用API: 自动扫描窗体上所有日期型文本框并绑定
'
' 在窗体的 Form_Load 中调用:
'   AttachDatePickerAll Me
'
' 自动检测字段类型为日期的文本框并绑定日期选择器
' 也可传入控件名前缀来匹配:
'   AttachDatePickerAll Me, "dt"     ' 匹配 dt 开头的文本框
'   AttachDatePickerAll Me, "date"   ' 匹配 date 开头的文本框
'====================================================
Public Sub AttachDatePickerAll(frm As Form, _
                               Optional ByVal sPrefix As String = "", _
                               Optional ByVal bShowTime As Boolean = True)
    On Error Resume Next
    
    Dim ctl As Control
    Dim bMatch As Boolean
    
    For Each ctl In frm.Controls
        ' 只处理文本框
        If TypeOf ctl Is TextBox Then
            bMatch = False
            
            ' 方式1: 按前缀匹配
            If sPrefix <> "" Then
                If LCase(Left(ctl.Name, Len(sPrefix))) = LCase(sPrefix) Then
                    bMatch = True
                End If
            Else
                ' 方式2: 自动检测 ControlSource 对应字段是否为日期型
                ' 检查控件名是否包含常见日期关键词
                Dim sName As String
                sName = LCase(ctl.Name)
                If InStr(sName, "date") > 0 Or _
                   InStr(sName, "日期") > 0 Or _
                   InStr(sName, "time") > 0 Or _
                   InStr(sName, "时间") > 0 Or _
                   InStr(sName, "dt") > 0 Then
                    bMatch = True
                End If
                
                ' 也检查控件 Format 属性
                If Not bMatch Then
                    Dim sFmt As String
                    sFmt = ""
                    Err.Clear
                    sFmt = ctl.Format
                    If Err.Number = 0 And sFmt <> "" Then
                        If InStr(LCase(sFmt), "date") > 0 Or _
                           InStr(sFmt, "yyyy") > 0 Or _
                           InStr(sFmt, "mm/dd") > 0 Or _
                           InStr(LCase(sFmt), "short date") > 0 Or _
                           InStr(LCase(sFmt), "long date") > 0 Then
                            bMatch = True
                        End If
                    End If
                End If
            End If
            
            If bMatch Then
                AttachDatePicker frm, ctl.Name, bShowTime
            End If
        End If
    Next ctl
    
    On Error GoTo 0
End Sub


'====================================================
' 内部: 设置返回结果 (由事件处理器调用)
'====================================================
Public Sub DatePicker_SetResult(ByVal vResult As Variant, ByVal bConfirmed As Boolean)
    m_dtResult = vResult
    m_bConfirmed = bConfirmed
End Sub


' ╔══════════════════════════════════════════════╗
' ║     事件处理函数 (由窗体控件表达式调用)        ║
' ╚══════════════════════════════════════════════╝

'---------- 窗体加载 ----------
Public Function DatePicker_FormLoad() As Variant
    On Error GoTo ErrHandler
    
    Dim frm As Form
    Set frm = Forms(FORM_NAME)
    
    ' 解析 OpenArgs
    Dim sArgs As String
    sArgs = Nz(frm.OpenArgs, "")
    
    Dim dt As Date
    dt = Now
    m_bShowTime = True
    
    If sArgs <> "" Then
        Dim parts() As String
        parts = Split(sArgs, "|")
        
        If UBound(parts) >= 0 And parts(0) <> "" Then
            On Error Resume Next
            dt = CDate(parts(0))
            If Err.Number <> 0 Then dt = Now
            On Error GoTo ErrHandler
        End If
        
        If UBound(parts) >= 1 Then
            m_bShowTime = (parts(1) = "1")
        End If
    End If
    
    m_iYear = Year(dt)
    m_iMonth = Month(dt)
    m_iDay = Day(dt)
    m_iHour = Hour(dt)
    m_iMinute = Minute(dt)
    
    ' 时间控件显示/隐藏
    frm.Controls("lstHour").Visible = m_bShowTime
    frm.Controls("lstMin").Visible = m_bShowTime
    frm.Controls("lblHour").Visible = m_bShowTime
    frm.Controls("lblMin").Visible = m_bShowTime
    
    ' ★ 无时间模式: 缩窄窗体宽度 (仅显示7列日历区域+两侧45缇边距)
    If Not m_bShowTime Then
        Dim lngNarrowW As Long
        lngNarrowW = 7 * 480 + 45 * 2   ' 7列 × 480缇 + 90缇边距 = 3450
        frm.Width = lngNarrowW
        frm.Controls("rctBorder").Width = lngNarrowW - 50
        frm.Controls("txtDisplay").Width = lngNarrowW - 180
    End If
    
    If m_bShowTime Then
        ' 填充小时列表 00-23
        Dim SH As String
        Dim i As Integer
        For i = 0 To 23
            If i > 0 Then SH = SH & ";"
            SH = SH & Format(i, "00")
        Next i
        
        ' 填充分钟列表 00-59
        Dim sM As String
        For i = 0 To 59
            If i > 0 Then sM = sM & ";"
            sM = sM & Format(i, "00")
        Next i
        
        With frm.Controls("lstHour")
            .RowSourceType = "Value List"
            .RowSource = SH
            .value = Format(m_iHour, "00")
        End With
        
        With frm.Controls("lstMin")
            .RowSourceType = "Value List"
            .RowSource = sM
            .value = Format(m_iMinute, "00")
        End With
        
        frm.Controls("lblHour").Caption = Format(m_iHour, "00")
        frm.Controls("lblMin").Caption = Format(m_iMinute, "00")
    End If
    
    ' 重置悬停状态
    m_iHoverCell = 0
    
    ' 刷新日历 & 显示
    DatePicker_RefreshCalendar
    DatePicker_UpdateDisplay
    
    DatePicker_FormLoad = True
    Exit Function

ErrHandler:
    MsgBox "日期选择器初始化失败: " & Err.Description, vbExclamation, "错误"
    DatePicker_FormLoad = False
End Function


'---------- 日期格子 单击 ----------
Public Function DatePicker_DayClick(ByVal iCell As Integer) As Variant
    On Error Resume Next
    
    If iCell < 1 Or iCell > 42 Then
        DatePicker_DayClick = False
        Exit Function
    End If
    
    Dim dt As Date
    dt = m_CellDates(iCell)
    m_iYear = Year(dt)
    m_iMonth = Month(dt)
    m_iDay = Day(dt)
    
    DatePicker_RefreshCalendar
    DatePicker_UpdateDisplay
    
    ' 非模态模式: 单击日期即确认 (确定/取消按钮已移除)
    If m_bIsNonModal Then
        DatePicker_ConfirmAndClose
    End If
    
    DatePicker_DayClick = True
End Function


'---------- 日期格子 鼠标悬停 ----------
Public Function DatePicker_DayMouseMove(ByVal iCell As Integer) As Variant
    On Error Resume Next
    
    If iCell < 1 Or iCell > 42 Then
        DatePicker_DayMouseMove = False
        Exit Function
    End If
    
    ' 与上次相同则忽略
    If iCell = m_iHoverCell Then
        DatePicker_DayMouseMove = True
        Exit Function
    End If
    
    Dim frm As Form
    Set frm = Forms(FORM_NAME)
    Dim ctl As Control
    
    ' --- 恢复上一个悬停格子的原始样式 ---
    If m_iHoverCell >= 1 And m_iHoverCell <= 42 Then
        Set ctl = frm.Controls("lblD" & Format(m_iHoverCell, "00"))
        Dim dtOld As Date
        dtOld = m_CellDates(m_iHoverCell)
        Dim dtSel As Date
        dtSel = DateSerial(m_iYear, m_iMonth, m_iDay)
        
        ' 重置为默认
        ctl.BackColor = CLR_WHITE
        ctl.ForeColor = CLR_BLACK
        ctl.FontWeight = 400
        
        If Month(dtOld) <> m_iMonth Then ctl.ForeColor = CLR_GRAY
        
        If dtOld = Date And dtOld <> dtSel Then
            ctl.ForeColor = CLR_BLUE
            ctl.FontWeight = 700
        End If
        
        If dtOld = dtSel Then
            ctl.BackColor = CLR_BLUE
            ctl.ForeColor = CLR_WHITE
            ctl.FontWeight = 700
        End If
    End If
    
    ' --- 设置新悬停格子的高亮样式 ---
    m_iHoverCell = iCell
    Set ctl = frm.Controls("lblD" & Format(iCell, "00"))
    Dim dtNew As Date
    dtNew = m_CellDates(iCell)
    Dim dtSelected As Date
    dtSelected = DateSerial(m_iYear, m_iMonth, m_iDay)
    
    ' 仅在非选中状态下显示悬停色
    If dtNew <> dtSelected Then
        ctl.BackColor = CLR_HOVER
        If Month(dtNew) <> m_iMonth Then
            ctl.ForeColor = CLR_GRAY
        Else
            ctl.ForeColor = CLR_BLACK
        End If
        If dtNew = Date Then
            ctl.ForeColor = CLR_BLUE
            ctl.FontWeight = 700
        End If
    End If
    
    DatePicker_DayMouseMove = True
End Function


'---------- 日期格子 双击 (选中并确认) ----------
Public Function DatePicker_DayDblClick(ByVal iCell As Integer) As Variant
    DatePicker_DayClick iCell
    DatePicker_ConfirmAndClose
    DatePicker_DayDblClick = True
End Function


'---------- 上个月 ----------
Public Function DatePicker_PrevMonth() As Variant
    m_iMonth = m_iMonth - 1
    If m_iMonth < 1 Then
        m_iMonth = 12
        m_iYear = m_iYear - 1
    End If
    DatePicker_RefreshCalendar
    DatePicker_UpdateDisplay
    DatePicker_PrevMonth = True
End Function


'---------- 下个月 ----------
Public Function DatePicker_NextMonth() As Variant
    m_iMonth = m_iMonth + 1
    If m_iMonth > 12 Then
        m_iMonth = 1
        m_iYear = m_iYear + 1
    End If
    DatePicker_RefreshCalendar
    DatePicker_UpdateDisplay
    DatePicker_NextMonth = True
End Function


'---------- 上一年 ----------
Public Function DatePicker_PrevYear() As Variant
    m_iYear = m_iYear - 1
    DatePicker_RefreshCalendar
    DatePicker_UpdateDisplay
    DatePicker_PrevYear = True
End Function


'---------- 下一年 ----------
Public Function DatePicker_NextYear() As Variant
    m_iYear = m_iYear + 1
    DatePicker_RefreshCalendar
    DatePicker_UpdateDisplay
    DatePicker_NextYear = True
End Function


'---------- 小时列表 点击 ----------
Public Function DatePicker_HourClick() As Variant
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_NAME)
    
    If Not IsNull(frm.Controls("lstHour").value) Then
        m_iHour = CInt(frm.Controls("lstHour").value)
        frm.Controls("lblHour").Caption = Format(m_iHour, "00")
        DatePicker_UpdateDisplay
    End If
    
    DatePicker_HourClick = True
End Function


'---------- 分钟列表 点击 ----------
Public Function DatePicker_MinClick() As Variant
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_NAME)
    
    If Not IsNull(frm.Controls("lstMin").value) Then
        m_iMinute = CInt(frm.Controls("lstMin").value)
        frm.Controls("lblMin").Caption = Format(m_iMinute, "00")
        DatePicker_UpdateDisplay
    End If
    
    DatePicker_MinClick = True
End Function


'---------- 清除 ----------
Public Function DatePicker_Clear() As Variant
    DatePicker_SetResult Null, True
    DatePicker_WriteBack Null
    DoCmd.Close acForm, FORM_NAME
    DatePicker_Clear = True
End Function


'---------- 今天 ----------
Public Function DatePicker_GoToday() As Variant
    On Error Resume Next
    
    m_iYear = Year(Date)
    m_iMonth = Month(Date)
    m_iDay = Day(Date)
    m_iHour = Hour(Now)
    m_iMinute = Minute(Now)
    
    DatePicker_RefreshCalendar
    DatePicker_UpdateDisplay
    
    If m_bShowTime Then
        Dim frm As Form
        Set frm = Forms(FORM_NAME)
        frm.Controls("lblHour").Caption = Format(m_iHour, "00")
        frm.Controls("lblMin").Caption = Format(m_iMinute, "00")
        frm.Controls("lstHour").value = Format(m_iHour, "00")
        frm.Controls("lstMin").value = Format(m_iMinute, "00")
    End If
    
    DatePicker_GoToday = True
End Function


'---------- 确定 ----------
Public Function DatePicker_OK() As Variant
    DatePicker_ConfirmAndClose
    DatePicker_OK = True
End Function


'---------- 取消 ----------
Public Function DatePicker_Cancel() As Variant
    Set g_txtDateInput = Nothing
    DatePicker_SetResult Null, False
    DoCmd.Close acForm, FORM_NAME
    DatePicker_Cancel = True
End Function


' ╔══════════════════════════════════════════════╗
' ║              内部辅助函数                      ║
' ╚══════════════════════════════════════════════╝

'----------------------------------------------------
' 确认选择并关闭窗体
'----------------------------------------------------
Private Sub DatePicker_ConfirmAndClose()
    Dim dt As Variant
    
    If m_bShowTime Then
        dt = DateSerial(m_iYear, m_iMonth, m_iDay) + TimeSerial(m_iHour, m_iMinute, 0)
    Else
        dt = DateSerial(m_iYear, m_iMonth, m_iDay)
    End If
    
    DatePicker_SetResult dt, True
    DatePicker_WriteBack dt
    DoCmd.Close acForm, FORM_NAME
End Sub


'----------------------------------------------------
' 刷新日历网格显示
'----------------------------------------------------
Private Sub DatePicker_RefreshCalendar()
    On Error Resume Next
    
    Dim frm As Form
    Set frm = Forms(FORM_NAME)
    
    ' 禁止重绘 (提升性能)
    frm.Painting = False
    
    ' 本月第一天
    Dim dtFirst As Date
    dtFirst = DateSerial(m_iYear, m_iMonth, 1)
    
    ' 第一天是星期几 (1=周一 ... 7=周日)
    Dim iFirstDow As Integer
    iFirstDow = Weekday(dtFirst, vbMonday)
    
    ' 网格起始日期 (可能是上月的某天)
    Dim dtStart As Date
    dtStart = dtFirst - iFirstDow + 1
    
    ' 确保选择日在有效范围内
    Dim iMaxDay As Integer
    iMaxDay = Day(DateSerial(m_iYear, m_iMonth + 1, 0))
    If m_iDay > iMaxDay Then m_iDay = iMaxDay
    
    Dim dtSelected As Date
    dtSelected = DateSerial(m_iYear, m_iMonth, m_iDay)
    
    ' 更新月份年份标签
    frm.Controls("lblMonthYear").Caption = m_iYear & "年" & Format(m_iMonth, "00") & "月"
    
    ' 填充42个日期格子
    Dim i As Integer
    Dim ctl As Control
    Dim dtCell As Date
    
    For i = 1 To 42
        dtCell = dtStart + i - 1
        m_CellDates(i) = dtCell
        
        Set ctl = frm.Controls("lblD" & Format(i, "00"))
        ctl.Caption = CStr(Day(dtCell))
        
        ' === 重置样式 ===
        ctl.BackColor = CLR_WHITE
        ctl.ForeColor = CLR_BLACK
        ctl.BorderStyle = 0             ' 无边框
        ctl.FontWeight = 400            ' 正常粗细
        
        ' --- 非当前月: 灰色文字 ---
        If Month(dtCell) <> m_iMonth Then
            ctl.ForeColor = CLR_GRAY
        End If
        
        ' --- 今天: 蓝色边框 ---
        If dtCell = Date Then
            ctl.BorderStyle = 1         ' 实线边框
            ctl.BorderColor = CLR_BLUE
            ctl.BorderWidth = 2
            If dtCell <> dtSelected Then
                ctl.ForeColor = CLR_BLUE
                ctl.FontWeight = 700    ' 加粗
            End If
        End If
        
        ' --- 已选中: 蓝色背景白字 ---
        If dtCell = dtSelected Then
            ctl.BackColor = CLR_BLUE
            ctl.ForeColor = CLR_WHITE
            ctl.FontWeight = 700
        End If
    Next i
    
    ' 恢复重绘
    frm.Painting = True
End Sub


'----------------------------------------------------
' 更新顶部日期时间显示
'----------------------------------------------------
Private Sub DatePicker_UpdateDisplay()
    On Error Resume Next
    
    Dim frm As Form
    Set frm = Forms(FORM_NAME)
    
    If m_bShowTime Then
        frm.Controls("txtDisplay").value = _
            Format(DateSerial(m_iYear, m_iMonth, m_iDay), "yyyy/mm/dd") & " " & _
            Format(m_iHour, "00") & ":" & Format(m_iMinute, "00")
    Else
        frm.Controls("txtDisplay").value = _
            Format(DateSerial(m_iYear, m_iMonth, m_iDay), "yyyy/mm/dd")
    End If
End Sub


'----------------------------------------------------
' 检查窗体是否存在
'----------------------------------------------------
Private Function FormExists(sName As String) As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = sName Then
            FormExists = True
            Exit Function
        End If
    Next obj
    FormExists = False
End Function


'----------------------------------------------------
' 检查窗体是否已打开
'----------------------------------------------------
Private Function IsFormOpen(sName As String) As Boolean
    On Error Resume Next
    IsFormOpen = (SysCmd(acSysCmdGetObjectState, acForm, sName) <> 0)
    If Err.Number <> 0 Then IsFormOpen = False
    On Error GoTo 0
End Function


'----------------------------------------------------
' 定位窗体到文本框正下方 (参照 modCalendar 的 CalendarFor 方法)
'----------------------------------------------------
Private Sub DatePicker_PositionBelowInput(frm As Form)
    On Error Resume Next
    
    Dim typRect_Calendar As RECT, typRect_Desktop As RECT
    Dim lngHwnd As Long
    Dim intLeft As Integer, intTop As Integer
    Dim intWidth As Integer, intHeight As Integer
    
    lngHwnd = frm.hwnd
    apiGetWindowRect lngHwnd, typRect_Calendar
    apiGetWindowRect apiGetDesktopWindow(), typRect_Desktop
    
    intWidth = typRect_Calendar.Right - typRect_Calendar.Left
    intHeight = typRect_Calendar.Bottom - typRect_Calendar.Top
    
    ' 左对齐文本框左边缘 (-3 补偿 Access 无边框窗体客户区内边距)
    intLeft = m_rcInputBox.Left - 3
    If intLeft < 0 Then
        intLeft = 0
    ElseIf intLeft + intWidth > typRect_Desktop.Right Then
        intLeft = typRect_Desktop.Right - intWidth
    End If
    
    ' 放在文本框正下方
    intTop = m_rcInputBox.Bottom
    If intTop < 0 Then
        intTop = 0
    ElseIf intTop + intHeight > typRect_Desktop.Bottom Then
        ' 下方放不下，放到文本框上方
        intTop = m_rcInputBox.Top - intHeight
    End If
    
    apiMoveWindow lngHwnd, intLeft, intTop, intWidth, intHeight, True
End Sub


'----------------------------------------------------
' 回写结果到调用控件 (非模态下拉模式)
'----------------------------------------------------
Private Sub DatePicker_WriteBack(ByVal dt As Variant)
    On Error Resume Next
    
    If Not m_bIsNonModal Then Exit Sub
    If g_txtDateInput Is Nothing Then Exit Sub
    
    ' 写入值到调用控件
    If Not IsNull(dt) Then
        If m_sCallbackFmt <> "" Then
            g_txtDateInput.value = Format(dt, m_sCallbackFmt)
        Else
            g_txtDateInput.value = dt
        End If
    ElseIf m_bConfirmed Then
        ' 用户点了"清除" → 清空
        g_txtDateInput.value = Null
    End If
    
    Set g_txtDateInput = Nothing
    m_sCallbackFmt = ""
End Sub


'----------------------------------------------------
' 窗体失去焦点 → 自动关闭 (非模态下拉模式)
'----------------------------------------------------
Public Function DatePicker_FormDeactivate() As Variant
    On Error Resume Next
    
    If m_bIsNonModal Then
        ' 延时关闭 (避免在控件切换焦点时误关)
        Forms(FORM_NAME).TimerInterval = 100
    End If
    
    DatePicker_FormDeactivate = True
End Function


'----------------------------------------------------
' 定时器: 延时关闭窗体
'----------------------------------------------------
Public Function DatePicker_Timer() As Variant
    On Error Resume Next
    
    Dim frm As Form
    Set frm = Forms(FORM_NAME)
    frm.TimerInterval = 0
    
    ' 检查是否仍然失焦
    If m_bIsNonModal Then
        Dim sActive As String
        sActive = ""
        sActive = Screen.ActiveForm.Name
        
        If sActive <> FORM_NAME Then
            ' 确认丢失焦点 → 取消并关闭
            Set g_txtDateInput = Nothing
            DatePicker_SetResult Null, False
            DoCmd.Close acForm, FORM_NAME
        End If
    End If
    
    DatePicker_Timer = True
End Function


' ╔══════════════════════════════════════════════╗
' ║         窗体自动构建器                         ║
' ║   在立即窗口(Ctrl+G)运行: CreateDatePickerForm ║
' ╚══════════════════════════════════════════════╝

Public Sub CreateDatePickerForm()
    On Error GoTo ErrHandler
    
    ' ========================================
    ' 布局参数 (单位: 缇 twips, 1英寸=1440缇)
    ' ========================================
    Dim LM As Long:     LM = 45         ' 左边距 (45缇内边距)
    Dim CW As Long:     CW = 480        ' 日期格子宽度
    Dim CH As Long:     CH = 340        ' 日期格子高度
    Dim FW As Long:     FW = 5370       ' 窗体宽度 (7*480+右侧时间区)
    Dim SH As Long:     SH = 3900       ' Detail区高度
    
    ' 垂直位置 (全部+45缇顶部内边距)
    Dim yDisp As Long:  yDisp = 105     ' 日期显示文本框
    Dim hDisp As Long:  hDisp = 380     ' 显示框高度
    Dim yNav As Long:   yNav = 545      ' 月份导航行
    Dim hNav As Long:   hNav = 370      ' 导航行高度
    Dim yDow As Long:   yDow = 965      ' 星期标头
    Dim hDow As Long:   hDow = 280      ' 星期标头高度
    Dim yGrid As Long:  yGrid = 1275    ' 日期网格起始
    Dim yBtn As Long:   yBtn = 3345     ' 底部按钮
    Dim hBtn As Long:   hBtn = 360      ' 按钮高度
    
    ' 时间区域
    Dim xTime As Long:  xTime = 3675    ' 时间区左边 (+45)
    Dim wTime As Long:  wTime = 700     ' 时间列宽度
    Dim gTime As Long:  gTime = 120     ' 时间两列间距
    
    ' ========================================
    ' 删除已有窗体
    ' ========================================
    If FormExists(FORM_NAME) Then
        On Error Resume Next
        DoCmd.Close acForm, FORM_NAME, acSaveNo
        On Error GoTo ErrHandler
        DoCmd.DeleteObject acForm, FORM_NAME
    End If
    
    ' ========================================
    ' 创建窗体
    ' ========================================
    Dim frm As Form
    Set frm = CreateForm
    Dim sTmp As String
    sTmp = frm.Name     ' 临时窗体名 (如 "Form1")
    
    ' --- 窗体属性 ---
    With frm
        .Caption = ""
        .DefaultView = 0                ' 单一窗体
        .ScrollBars = 0                 ' 无滚动条
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = False             ' 不自动居中 (我们手动定位)
        .PopUp = True
        .Modal = False                  ' 非模态 (支持点击外部关闭)
        .BorderStyle = 0                ' ★ 无边框 (无标题栏)
        .MinMaxButtons = 0              ' 无最小最大化
        .CloseButton = False
        .Width = FW
        .Section(acDetail).Height = SH
        .Section(acDetail).BackColor = CLR_WHITE
    End With
    
    ' --- 窗体事件 ---
    frm.OnLoad = "=DatePicker_FormLoad()"
    frm.OnDeactivate = "=DatePicker_FormDeactivate()"
    frm.OnTimer = "=DatePicker_Timer()"
    frm.KeyPreview = True
    
    Dim ctl As Control
    
    ' ========================================
    ' 0. 边框矩形 (视觉边框，因为窗体无边框)
    ' ========================================
    Set ctl = CreateControl(sTmp, acRectangle, acDetail, "", "", _
                            45, 45, FW - 90, SH - 90)
    ctl.Name = "rctBorder"
    ctl.BorderStyle = 1             ' 实线
    ctl.BorderColor = CLR_BORDER
    ctl.BorderWidth = 1
    ctl.BackStyle = 0               ' 透明
    ' ========================================
    ' 1. 日期显示文本框 (顶部)
    ' ========================================
    Set ctl = CreateControl(sTmp, acTextBox, acDetail, "", "", _
                            LM, yDisp, FW - LM * 2, hDisp)
    ctl.Name = "txtDisplay"
    ctl.FontSize = 11
    ctl.FontName = "Segoe UI"
    ctl.Locked = True
    ctl.TabStop = False
    ctl.BackColor = CLR_WHITE
    ctl.ForeColor = CLR_BLACK
    ctl.BorderStyle = 1
    ctl.BorderColor = CLR_BORDER
    ' 删除可能自动生成的关联标签
    On Error Resume Next
    DeleteControl sTmp, ctl.Controls(0).Name
    On Error GoTo ErrHandler
    
    ' ========================================
    ' 2. 月份/年份导航
    ' ========================================
    
    ' 月份年份标签
    Set ctl = CreateControl(sTmp, acLabel, acDetail, "", "", _
                            LM + 45, yNav, 2100, hNav)
    ctl.Name = "lblMonthYear"
    ctl.Caption = "2026年02月"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.FontWeight = 700
    ctl.ForeColor = CLR_BLACK
    ctl.BackStyle = 0               ' 透明
    ctl.TextAlign = 1               ' 左对齐
    
    ' << 上一年按钮
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            1645, yNav, 420, hNav)
    ctl.Name = "cmdPrevYear"
    ctl.Caption = "<<"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=DatePicker_PrevYear()"
    
    ' < 上个月按钮
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            2065, yNav, 420, hNav)
    ctl.Name = "cmdPrev"
    ctl.Caption = "<"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=DatePicker_PrevMonth()"
    
    ' > 下个月按钮
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            2485, yNav, 420, hNav)
    ctl.Name = "cmdNext"
    ctl.Caption = ">"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=DatePicker_NextMonth()"
    
    ' >> 下一年按钮
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            2905, yNav, 420, hNav)
    ctl.Name = "cmdNextYear"
    ctl.Caption = ">>"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=DatePicker_NextYear()"
    
    ' ========================================
    ' 3. 时间显示头部 (蓝色标签)
    ' ========================================
    
    ' 小时 显示
    Set ctl = CreateControl(sTmp, acLabel, acDetail, "", "", _
                            xTime, yNav, wTime, hNav)
    ctl.Name = "lblHour"
    ctl.Caption = "15"
    ctl.FontSize = 14
    ctl.FontName = "Segoe UI"
    ctl.FontWeight = 700
    ctl.TextAlign = 2               ' 居中
    ctl.BackStyle = 1               ' 不透明
    ctl.BackColor = CLR_BLUE
    ctl.ForeColor = CLR_WHITE
    
    ' 分钟 显示
    Set ctl = CreateControl(sTmp, acLabel, acDetail, "", "", _
                            xTime + wTime + gTime, yNav, wTime, hNav)
    ctl.Name = "lblMin"
    ctl.Caption = "01"
    ctl.FontSize = 14
    ctl.FontName = "Segoe UI"
    ctl.FontWeight = 700
    ctl.TextAlign = 2
    ctl.BackStyle = 1
    ctl.BackColor = CLR_BLUE
    ctl.ForeColor = CLR_WHITE
    
    ' ========================================
    ' 4. 星期标头 (一 二 三 四 五 六 日)
    ' ========================================
    Dim sDows As Variant
    sDows = Array("一", "二", "三", "四", "五", "六", "日")
    
    Dim iCol As Integer
    For iCol = 0 To 6
        Set ctl = CreateControl(sTmp, acLabel, acDetail, "", "", _
                                LM + iCol * CW, yDow, CW, hDow)
        ctl.Name = "lblW" & (iCol + 1)
        ctl.Caption = sDows(iCol)
        ctl.TextAlign = 2           ' 居中
        ctl.FontSize = 9
        ctl.FontName = "Segoe UI"
        ctl.ForeColor = CLR_DKGRAY
        ctl.BackStyle = 0           ' 透明
        
        ' 周末（周六、周日）用蓝灰色
        If iCol >= 5 Then
            ctl.ForeColor = CLR_BLUE
        End If
    Next iCol
    
    ' ========================================
    ' 5. 42个日期格子 (6行 × 7列 的标签)
    ' ========================================
    Dim iRow As Integer, iCell As Integer
    iCell = 0
    
    For iRow = 0 To 5
        For iCol = 0 To 6
            iCell = iCell + 1
            
            Set ctl = CreateControl(sTmp, acLabel, acDetail, "", "", _
                                    LM + iCol * CW, _
                                    yGrid + iRow * CH, _
                                    CW, CH)
            ctl.Name = "lblD" & Format(iCell, "00")
            ctl.Caption = ""
            ctl.TextAlign = 2       ' 居中
            ctl.BackStyle = 1       ' 不透明 (才能设背景色)
            ctl.BackColor = CLR_WHITE
            ctl.ForeColor = CLR_BLACK
            ctl.FontSize = 9
            ctl.FontName = "Segoe UI"
            ctl.BorderStyle = 0     ' 无边框
            
            ' 绑定事件
            ctl.OnClick = "=DatePicker_DayClick(" & iCell & ")"
            ctl.OnDblClick = "=DatePicker_DayDblClick(" & iCell & ")"
            ctl.OnMouseMove = "=DatePicker_DayMouseMove(" & iCell & ")"
        Next iCol
    Next iRow
    
    ' ========================================
    ' 6. 时间选择列表
    ' ========================================
    Dim yListTop As Long
    yListTop = yDow
    Dim hList As Long
    hList = yBtn - yDow             ' 列表高度 = 网格区域高度
    
    ' 小时列表
    Set ctl = CreateControl(sTmp, acListBox, acDetail, "", "", _
                            xTime, yListTop, wTime, hList)
    ctl.Name = "lstHour"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.ColumnCount = 1
    ctl.AllowValueListEdits = False    ' 禁止编辑值列表
    ctl.OnClick = "=DatePicker_HourClick()"
    ' 删除可能的关联标签
    On Error Resume Next
    DeleteControl sTmp, ctl.Controls(0).Name
    On Error GoTo ErrHandler
    
    ' 分钟列表
    Set ctl = CreateControl(sTmp, acListBox, acDetail, "", "", _
                            xTime + wTime + gTime, yListTop, wTime, hList)
    ctl.Name = "lstMin"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.ColumnCount = 1
    ctl.AllowValueListEdits = False    ' 禁止编辑值列表
    ctl.OnClick = "=DatePicker_MinClick()"
    ' 删除可能的关联标签
    On Error Resume Next
    DeleteControl sTmp, ctl.Controls(0).Name
    On Error GoTo ErrHandler
    
    ' ========================================
    ' 7. 底部按钮
    ' ========================================
    
    ' 清除
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            LM + 45, yBtn, 960, hBtn)
    ctl.Name = "cmdClear"
    ctl.Caption = "清除"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=DatePicker_Clear()"
    
    ' 今天
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            2205, yBtn, 960, hBtn)
    ctl.Name = "cmdToday"
    ctl.Caption = "今天"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=DatePicker_GoToday()"
    
    ' (确定/取消按钮已移除，选日期单击即确认，点击外部自动关闭)
    
    ' ========================================
    ' 保存并重命名窗体
    ' ========================================
    DoCmd.Close acForm, sTmp, acSaveYes
    DoCmd.Rename FORM_NAME, acForm, sTmp
    
    MsgBox "日期选择器窗体 [" & FORM_NAME & "] 创建成功！" & vbCrLf & vbCrLf & _
           "═══ 使用方法 ═══" & vbCrLf & vbCrLf & _
           "  Dim dt As Variant" & vbCrLf & _
           "  dt = ShowDatePicker()" & vbCrLf & vbCrLf & _
           "含日期+时间:" & vbCrLf & _
           "  dt = ShowDatePicker(Now, True)" & vbCrLf & vbCrLf & _
           "仅选日期:" & vbCrLf & _
           "  dt = ShowDatePicker(Now, False)", _
           vbInformation, "创建成功"
    
    Exit Sub

ErrHandler:
    MsgBox "创建窗体失败: " & Err.Description & vbCrLf & _
           "错误号: " & Err.Number, vbCritical, "创建失败"
    ' 尝试清理
    On Error Resume Next
    DoCmd.Close acForm, sTmp, acSaveNo
End Sub