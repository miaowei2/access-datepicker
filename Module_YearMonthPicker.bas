Attribute VB_Name = "Module_YearMonthPicker"

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
' 模块名称: Module_YearMonthPicker
' 功能描述: Access纯窗体年月选择器 (无需Web控件)
' 作者: edonsoft缪炜
' 版本: 1.0
'
' 特点:
'   - 纯VBA + Access窗体实现，零外部依赖
'   - 4列×3行月份网格 + 年份导航
'   - 自动创建窗体，一键部署
'   - 所有事件通过表达式绑定，无需窗体代码模块
'
' 使用方法:
'   1. 导入本模块到Access数据库
'   2. 在立即窗口运行: CreateYearMonthPickerForm
'   3. 在代码中调用:
'      Dim dt As Variant
'      dt = ShowYearMonthPicker()            ' 返回月份首日日期
'      PickYearMonthFor Me.txtMonth          ' 绑定文本框
'      PickYearMonthFor Me.txtMonth, "yyyy-mm"
'====================================================

' ============ 颜色常量 ============
Private Const CLR_BLUE As Long = 16024898       ' RGB(66,133,244) 主题蓝
Private Const CLR_WHITE As Long = 16777215      ' RGB(255,255,255)
Private Const CLR_BLACK As Long = 0             ' RGB(0,0,0)
Private Const CLR_BORDER As Long = 13158600     ' RGB(200,200,200) 输入框边框
Private Const CLR_HOVER As Long = 16576233      ' RGB(233,240,252) 鼠标悬停

Private m_ymHoverMonth As Integer                ' 当前悬停的月份索引 (0=无)

' ============ 窗体名 ============
Private Const YM_FORM_NAME As String = "frmYearMonthPicker"

' ============ 模块级变量 ============
Private m_ymYear As Integer             ' 当前显示/选择的年份
Private m_ymMonth As Integer            ' 当前选择的月份
Private m_ymResult As Variant           ' 返回结果
Private m_ymConfirmed As Boolean        ' 是否确认

Public g_txtYMInput As TextBox          ' 年月选择器调用时的文本框引用
Private m_rcYMInputBox As RECT          ' 调用时文本框的屏幕位置
Private m_sYMCallbackFmt As String      ' 回调: 日期格式
Private m_bYMIsNonModal As Boolean      ' 是否非模态 (下拉模式)


'====================================================
' ★ 公共API: 显示年月选择器
'
' 参数:
'   dtDefault  - 默认日期 (可选, 默认=Now, 取其年月)
'
' 返回:
'   Date   - 所选月份第1天的日期 (如 2026-02-01)
'   Null   - 用户取消或清除
'
' 示例:
'   Dim dt As Variant
'   dt = ShowYearMonthPicker()
'   dt = ShowYearMonthPicker(#2026-6-15#)
'====================================================
Public Function ShowYearMonthPicker(Optional ByVal dtDefault As Variant) As Variant
    ' 检查窗体是否存在
    If Not FormExists(YM_FORM_NAME) Then
        MsgBox "年月选择器窗体 [" & YM_FORM_NAME & "] 不存在！" & vbCrLf & _
               "请先在立即窗口(Ctrl+G)中运行:" & vbCrLf & _
               "  CreateYearMonthPickerForm", vbExclamation, "年月选择器"
        ShowYearMonthPicker = Null
        Exit Function
    End If
    
    ' 初始化返回状态
    m_ymConfirmed = False
    m_ymResult = Null
    
    ' 构建 OpenArgs
    Dim sArgs As String
    If IsMissing(dtDefault) Or IsEmpty(dtDefault) Or IsNull(dtDefault) Then
        sArgs = Format(Now, "yyyy/mm")
    Else
        On Error Resume Next
        sArgs = Format(CDate(dtDefault), "yyyy/mm")
        If Err.Number <> 0 Then sArgs = Format(Now, "yyyy/mm")
        On Error GoTo 0
    End If
    
    m_bYMIsNonModal = False
    
    ' 以对话框模式打开
    DoCmd.OpenForm YM_FORM_NAME, acNormal, , , , acDialog, sArgs
    
    If m_ymConfirmed Then
        ShowYearMonthPicker = m_ymResult
    Else
        ShowYearMonthPicker = Null
    End If
End Function


'====================================================
' ★ 通用API: 为任意控件弹出年月选择器并回写结果
'
' 用法:
'   Private Sub txtMonth_Click()
'       PickYearMonthFor Me.txtMonth
'   End Sub
'
'   PickYearMonthFor Me.txtMonth, "yyyy-mm"
'   PickYearMonthFor Me.txtMonth, "yyyy年mm月"
'====================================================
Public Sub PickYearMonthFor(ctl As Control, _
                            Optional ByVal sFormat As String = "")
On Error GoTo Err_Pick
    
    If Not FormExists(YM_FORM_NAME) Then
        MsgBox "年月选择器窗体 [" & YM_FORM_NAME & "] 不存在！" & vbCrLf & _
               "请先在立即窗口(Ctrl+G)中运行:" & vbCrLf & _
               "  CreateYearMonthPickerForm", vbExclamation, "年月选择器"
        Exit Sub
    End If
    
    If IsFormOpen(YM_FORM_NAME) Then
        DoCmd.Close acForm, YM_FORM_NAME
    End If
    
    Set g_txtYMInput = ctl
    m_sYMCallbackFmt = sFormat
    m_bYMIsNonModal = True
    m_ymConfirmed = False
    m_ymResult = Null
    
    ' 获取文本框屏幕位置
    ctl.SetFocus
    apiGetWindowRect apiGetFocus(), m_rcYMInputBox
    
    ' 解析默认值
    Dim sArgs As String
    Dim dtDefault As Variant
    dtDefault = Null
    On Error Resume Next
    If Not IsNull(ctl.value) And ctl.value <> "" Then
        ' 尝试直接解析
        dtDefault = CDate(ctl.value)
        If Err.Number <> 0 Then
            Err.Clear
            ' 尝试 yyyy-mm 格式
            dtDefault = CDate(ctl.value & "-01")
            If Err.Number <> 0 Then
                Err.Clear
                dtDefault = Null
            End If
        End If
    End If
    On Error GoTo Err_Pick
    
    If IsNull(dtDefault) Then
        sArgs = Format(Now, "yyyy/mm")
    Else
        sArgs = Format(dtDefault, "yyyy/mm")
    End If
    
    ' 非模态打开
    DoCmd.OpenForm YM_FORM_NAME, acNormal, , , , , sArgs
    
    ' 定位到文本框下方
    Dim typRect_YM As RECT, typRect_Desktop As RECT
    Dim lngHwnd As Long
    Dim intLeft As Integer, intTop As Integer
    Dim intWidth As Integer, intHeight As Integer
    
    lngHwnd = Forms(YM_FORM_NAME).hwnd
    apiGetWindowRect lngHwnd, typRect_YM
    apiGetWindowRect apiGetDesktopWindow(), typRect_Desktop
    
    intWidth = typRect_YM.Right - typRect_YM.Left
    intHeight = typRect_YM.Bottom - typRect_YM.Top
    
    intLeft = m_rcYMInputBox.Left - 1
    If intLeft < 0 Then intLeft = 0
    If intLeft + intWidth > typRect_Desktop.Right Then _
        intLeft = typRect_Desktop.Right - intWidth
    
    intTop = m_rcYMInputBox.Bottom
    If intTop < 0 Then intTop = 0
    If intTop + intHeight > typRect_Desktop.Bottom Then _
        intTop = m_rcYMInputBox.Top - intHeight
    
    apiMoveWindow lngHwnd, intLeft, intTop, intWidth, intHeight, True
    
Exit_Pick:
    Exit Sub

Err_Pick:
    MsgBox Err.Description, vbCritical, "年月选择器"
    Resume Exit_Pick
End Sub


'====================================================
' ★ 通用API: 通过窗体名+控件名弹出年月选择器 (支持事件表达式)
'
' 在控件属性的"单击"/"双击"事件中直接填写:
'   =PickYearMonthForCtl("frmOrder","txtMonth")
'   =PickYearMonthForCtl("frmOrder","txtMonth","yyyy-mm")
'====================================================
Public Function PickYearMonthForCtl(ByVal sFormName As String, _
                                    ByVal sCtlName As String, _
                                    Optional ByVal sFormat As String = "") As Variant
    On Error GoTo ErrHandler
    PickYearMonthFor Forms(sFormName).Controls(sCtlName), sFormat
    PickYearMonthForCtl = True
    Exit Function
ErrHandler:
    MsgBox "PickYearMonthForCtl 错误: " & Err.Description, vbExclamation, "年月选择器"
    PickYearMonthForCtl = False
End Function


'====================================================
' ★ 通用API: 一键为文本框绑定年月选择器
'
' 在窗体的 Form_Load 中调用:
'   AttachYearMonthPicker Me, "txtMonth"
'   AttachYearMonthPicker Me, "txtMonth", "yyyy-mm"
'====================================================
Public Sub AttachYearMonthPicker(frm As Form, _
                                 ByVal sCtlName As String, _
                                 Optional ByVal sFormat As String = "")
    On Error GoTo ErrHandler
    
    Dim ctl As Control
    Set ctl = frm.Controls(sCtlName)
    
    Dim sExpr As String
    If sFormat <> "" Then
        sExpr = "=PickYearMonthForCtl(""" & frm.Name & """,""" & sCtlName & """,""" & sFormat & """)"
    Else
        sExpr = "=PickYearMonthForCtl(""" & frm.Name & """,""" & sCtlName & """)"
    End If
    
    ctl.OnDblClick = sExpr
    Exit Sub
ErrHandler:
    Debug.Print "AttachYearMonthPicker 错误 [" & sCtlName & "]: " & Err.Description
End Sub


' ╔══════════════════════════════════════════════╗
' ║  年月选择器 事件处理函数 (由窗体控件表达式调用)  ║
' ╚══════════════════════════════════════════════╝

'---------- 年月选择器 设置返回结果 ----------
Public Sub YMPicker_SetResult(ByVal vResult As Variant, ByVal bConfirmed As Boolean)
    m_ymResult = vResult
    m_ymConfirmed = bConfirmed
End Sub


'---------- 年月选择器 窗体加载 ----------
Public Function YMPicker_FormLoad() As Variant
    On Error GoTo ErrHandler
    
    Dim frm As Form
    Set frm = Forms(YM_FORM_NAME)
    
    ' 解析 OpenArgs (格式: "yyyy/mm")
    Dim sArgs As String
    sArgs = Nz(frm.OpenArgs, "")
    
    If sArgs <> "" Then
        On Error Resume Next
        Dim parts() As String
        parts = Split(sArgs, "/")
        If UBound(parts) >= 1 Then
            m_ymYear = CInt(parts(0))
            m_ymMonth = CInt(parts(1))
        Else
            m_ymYear = Year(Now)
            m_ymMonth = Month(Now)
        End If
        If Err.Number <> 0 Then
            m_ymYear = Year(Now)
            m_ymMonth = Month(Now)
        End If
        On Error GoTo ErrHandler
    Else
        m_ymYear = Year(Now)
        m_ymMonth = Month(Now)
    End If
    
    ' 重置悬停状态
    m_ymHoverMonth = 0
    
    YMPicker_RefreshGrid
    YMPicker_UpdateDisplay
    
    YMPicker_FormLoad = True
    Exit Function
ErrHandler:
    MsgBox "年月选择器初始化失败: " & Err.Description, vbExclamation, "错误"
    YMPicker_FormLoad = False
End Function


'---------- 月份格子 单击 ----------
Public Function YMPicker_MonthClick(ByVal iMonth As Integer) As Variant
    On Error Resume Next
    
    If iMonth < 1 Or iMonth > 12 Then
        YMPicker_MonthClick = False
        Exit Function
    End If
    
    m_ymMonth = iMonth
    YMPicker_RefreshGrid
    YMPicker_UpdateDisplay
    
    ' 非模态: 选月即确认
    If m_bYMIsNonModal Then
        YMPicker_ConfirmAndClose
    End If
    
    YMPicker_MonthClick = True
End Function


'---------- 月份格子 鼠标悬停 ----------
Public Function YMPicker_MonthMouseMove(ByVal iMonth As Integer) As Variant
    On Error Resume Next
    
    If iMonth < 1 Or iMonth > 12 Then
        YMPicker_MonthMouseMove = False
        Exit Function
    End If
    
    ' 与上次相同则忽略
    If iMonth = m_ymHoverMonth Then
        YMPicker_MonthMouseMove = True
        Exit Function
    End If
    
    Dim frm As Form
    Set frm = Forms(YM_FORM_NAME)
    Dim ctl As Control
    
    ' --- 恢复上一个悬停格子的原始样式 ---
    If m_ymHoverMonth >= 1 And m_ymHoverMonth <= 12 Then
        Set ctl = frm.Controls("lblM" & Format(m_ymHoverMonth, "00"))
        
        ' 重置为默认
        ctl.BackColor = CLR_WHITE
        ctl.ForeColor = CLR_BLACK
        ctl.FontWeight = 400
        ctl.BorderStyle = 0
        
        ' 当前月份 (今天): 蓝色边框
        If m_ymYear = Year(Date) And m_ymHoverMonth = Month(Date) Then
            ctl.BorderStyle = 1
            ctl.BorderColor = CLR_BLUE
            ctl.BorderWidth = 2
            If m_ymHoverMonth <> m_ymMonth Then
                ctl.ForeColor = CLR_BLUE
                ctl.FontWeight = 700
            End If
        End If
        
        ' 已选中月份: 蓝色背景白字
        If m_ymHoverMonth = m_ymMonth Then
            ctl.BackColor = CLR_BLUE
            ctl.ForeColor = CLR_WHITE
            ctl.FontWeight = 700
        End If
    End If
    
    ' --- 设置新悬停格子的高亮样式 ---
    m_ymHoverMonth = iMonth
    Set ctl = frm.Controls("lblM" & Format(iMonth, "00"))
    
    ' 仅在非选中状态下显示悬停色
    If iMonth <> m_ymMonth Then
        ctl.BackColor = CLR_HOVER
        ctl.ForeColor = CLR_BLACK
        If m_ymYear = Year(Date) And iMonth = Month(Date) Then
            ctl.ForeColor = CLR_BLUE
            ctl.FontWeight = 700
        End If
    End If
    
    YMPicker_MonthMouseMove = True
End Function


'---------- 月份格子 双击 (选中并确认) ----------
Public Function YMPicker_MonthDblClick(ByVal iMonth As Integer) As Variant
    YMPicker_MonthClick iMonth
    YMPicker_ConfirmAndClose
    YMPicker_MonthDblClick = True
End Function


'---------- 上一年 ----------
Public Function YMPicker_PrevYear() As Variant
    m_ymYear = m_ymYear - 1
    YMPicker_RefreshGrid
    YMPicker_UpdateDisplay
    YMPicker_PrevYear = True
End Function


'---------- 下一年 ----------
Public Function YMPicker_NextYear() As Variant
    m_ymYear = m_ymYear + 1
    YMPicker_RefreshGrid
    YMPicker_UpdateDisplay
    YMPicker_NextYear = True
End Function


'---------- 上十年 ----------
Public Function YMPicker_PrevDecade() As Variant
    m_ymYear = m_ymYear - 10
    YMPicker_RefreshGrid
    YMPicker_UpdateDisplay
    YMPicker_PrevDecade = True
End Function


'---------- 下十年 ----------
Public Function YMPicker_NextDecade() As Variant
    m_ymYear = m_ymYear + 10
    YMPicker_RefreshGrid
    YMPicker_UpdateDisplay
    YMPicker_NextDecade = True
End Function


'---------- 清除 ----------
Public Function YMPicker_Clear() As Variant
    YMPicker_SetResult Null, True
    YMPicker_WriteBack Null
    DoCmd.Close acForm, YM_FORM_NAME
    YMPicker_Clear = True
End Function


'---------- 本月 ----------
Public Function YMPicker_GoThisMonth() As Variant
    m_ymYear = Year(Date)
    m_ymMonth = Month(Date)
    YMPicker_RefreshGrid
    YMPicker_UpdateDisplay
    YMPicker_GoThisMonth = True
End Function


'---------- 确定 ----------
Public Function YMPicker_OK() As Variant
    YMPicker_ConfirmAndClose
    YMPicker_OK = True
End Function


'---------- 取消 ----------
Public Function YMPicker_Cancel() As Variant
    Set g_txtYMInput = Nothing
    YMPicker_SetResult Null, False
    DoCmd.Close acForm, YM_FORM_NAME
    YMPicker_Cancel = True
End Function


' ╔══════════════════════════════════════════════╗
' ║        年月选择器 内部辅助函数                  ║
' ╚══════════════════════════════════════════════╝

'----------------------------------------------------
' 确认选择并关闭
'----------------------------------------------------
Private Sub YMPicker_ConfirmAndClose()
    Dim dt As Date
    dt = DateSerial(m_ymYear, m_ymMonth, 1)
    
    YMPicker_SetResult dt, True
    YMPicker_WriteBack dt
    DoCmd.Close acForm, YM_FORM_NAME
End Sub


'----------------------------------------------------
' 刷新月份网格
'----------------------------------------------------
Private Sub YMPicker_RefreshGrid()
    On Error Resume Next
    
    Dim frm As Form
    Set frm = Forms(YM_FORM_NAME)
    frm.Painting = False
    
    ' 更新年份标签
    frm.Controls("lblYear").Caption = m_ymYear & "年"
    
    ' 刷新12个月份按钮
    Dim i As Integer
    Dim ctl As Control
    
    For i = 1 To 12
        Set ctl = frm.Controls("lblM" & Format(i, "00"))
        ctl.Caption = Format(i, "00") & "月"
        
        ' 重置样式
        ctl.BackColor = CLR_WHITE
        ctl.ForeColor = CLR_BLACK
        ctl.BorderStyle = 0
        ctl.FontWeight = 400
        
        ' 当前月份 (今天的年月): 蓝色边框
        If m_ymYear = Year(Date) And i = Month(Date) Then
            ctl.BorderStyle = 1
            ctl.BorderColor = CLR_BLUE
            ctl.BorderWidth = 2
            If i <> m_ymMonth Then
                ctl.ForeColor = CLR_BLUE
                ctl.FontWeight = 700
            End If
        End If
        
        ' 已选中月份: 蓝色背景白字
        If i = m_ymMonth Then
            ctl.BackColor = CLR_BLUE
            ctl.ForeColor = CLR_WHITE
            ctl.FontWeight = 700
        End If
    Next i
    
    frm.Painting = True
End Sub


'----------------------------------------------------
' 更新顶部显示
'----------------------------------------------------
Private Sub YMPicker_UpdateDisplay()
    On Error Resume Next
    
    Dim frm As Form
    Set frm = Forms(YM_FORM_NAME)
    
    frm.Controls("txtYMDisplay").value = m_ymYear & "年" & Format(m_ymMonth, "00") & "月"
End Sub


'----------------------------------------------------
' 回写结果到调用控件
'----------------------------------------------------
Private Sub YMPicker_WriteBack(ByVal dt As Variant)
    On Error Resume Next
    
    If Not m_bYMIsNonModal Then Exit Sub
    If g_txtYMInput Is Nothing Then Exit Sub
    
    If Not IsNull(dt) Then
        If m_sYMCallbackFmt <> "" Then
            g_txtYMInput.value = Format(dt, m_sYMCallbackFmt)
        Else
            ' 默认格式: yyyy-mm
            g_txtYMInput.value = Format(dt, "yyyy-mm")
        End If
    ElseIf m_ymConfirmed Then
        g_txtYMInput.value = Null
    End If
    
    Set g_txtYMInput = Nothing
    m_sYMCallbackFmt = ""
End Sub


'----------------------------------------------------
' 窗体失去焦点 → 自动关闭
'----------------------------------------------------
Public Function YMPicker_FormDeactivate() As Variant
    On Error Resume Next
    
    If m_bYMIsNonModal Then
        Forms(YM_FORM_NAME).TimerInterval = 100
    End If
    
    YMPicker_FormDeactivate = True
End Function


'----------------------------------------------------
' 定时器: 延时关闭
'----------------------------------------------------
Public Function YMPicker_Timer() As Variant
    On Error Resume Next
    
    Dim frm As Form
    Set frm = Forms(YM_FORM_NAME)
    frm.TimerInterval = 0
    
    If m_bYMIsNonModal Then
        Dim sActive As String
        sActive = ""
        sActive = Screen.ActiveForm.Name
        
        If sActive <> YM_FORM_NAME Then
            Set g_txtYMInput = Nothing
            YMPicker_SetResult Null, False
            DoCmd.Close acForm, YM_FORM_NAME
        End If
    End If
    
    YMPicker_Timer = True
End Function


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


' ╔══════════════════════════════════════════════════════════════╗
' ║         年月选择器 窗体自动构建器                              ║
' ║   在立即窗口(Ctrl+G)运行: CreateYearMonthPickerForm           ║
' ╚══════════════════════════════════════════════════════════════╝

Public Sub CreateYearMonthPickerForm()
    On Error GoTo ErrHandler
    
    ' ========================================
    ' 布局参数 (单位: 缇 twips)
    ' ========================================
    Dim LM As Long:     LM = 45         ' 左边距
    Dim MW As Long:     MW = 720        ' 月份格子宽度
    Dim MH As Long:     MH = 420        ' 月份格子高度
    Dim MG As Long:     MG = 60         ' 月份格子间距
    Dim FW As Long                      ' 窗体宽度 (4列 × 格子宽 + 3间距 + 左右边距)
    FW = LM * 2 + MW * 4 + MG * 3      ' = 90 + 2880 + 180 = 3150
    Dim SH As Long                      ' Detail区高度
    
    ' 垂直位置
    Dim yDisp As Long:  yDisp = 105     ' 显示文本框
    Dim hDisp As Long:  hDisp = 380     ' 显示框高度
    Dim yNav As Long:   yNav = 545      ' 年份导航行
    Dim hNav As Long:   hNav = 370      ' 导航行高度
    Dim yGrid As Long:  yGrid = 985     ' 月份网格起始
    Dim yBtn As Long                    ' 底部按钮
    yBtn = yGrid + MH * 3 + MG * 2 + 90   ' = 985 + 1260 + 120 + 90 = 2455
    Dim hBtn As Long:   hBtn = 360      ' 按钮高度
    SH = yBtn + hBtn + 105              ' = 2455 + 360 + 105 = 2920
    
    ' ========================================
    ' 删除已有窗体
    ' ========================================
    If FormExists(YM_FORM_NAME) Then
        On Error Resume Next
        DoCmd.Close acForm, YM_FORM_NAME, acSaveNo
        On Error GoTo ErrHandler
        DoCmd.DeleteObject acForm, YM_FORM_NAME
    End If
    
    ' ========================================
    ' 创建窗体
    ' ========================================
    Dim frm As Form
    Set frm = CreateForm
    Dim sTmp As String
    sTmp = frm.Name
    
    With frm
        .Caption = ""
        .DefaultView = 0
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = False
        .PopUp = True
        .Modal = False
        .BorderStyle = 0
        .MinMaxButtons = 0
        .CloseButton = False
        .Width = FW
        .Section(acDetail).Height = SH
        .Section(acDetail).BackColor = CLR_WHITE
    End With
    
    ' 窗体事件
    frm.OnLoad = "=YMPicker_FormLoad()"
    frm.OnDeactivate = "=YMPicker_FormDeactivate()"
    frm.OnTimer = "=YMPicker_Timer()"
    frm.KeyPreview = True
    
    Dim ctl As Control
    
    ' ========================================
    ' 0. 边框矩形
    ' ========================================
    Set ctl = CreateControl(sTmp, acRectangle, acDetail, "", "", _
                            45, 45, FW - 90, SH - 90)
    ctl.Name = "rctBorder"
    ctl.BorderStyle = 1
    ctl.BorderColor = CLR_BORDER
    ctl.BorderWidth = 1
    ctl.BackStyle = 0
    
    ' ========================================
    ' 1. 年月显示文本框 (顶部)
    ' ========================================
    Set ctl = CreateControl(sTmp, acTextBox, acDetail, "", "", _
                            LM, yDisp, FW - LM * 2, hDisp)
    ctl.Name = "txtYMDisplay"
    ctl.FontSize = 11
    ctl.FontName = "Segoe UI"
    ctl.Locked = True
    ctl.TabStop = False
    ctl.BackColor = CLR_WHITE
    ctl.ForeColor = CLR_BLACK
    ctl.BorderStyle = 1
    ctl.BorderColor = CLR_BORDER
    On Error Resume Next
    DeleteControl sTmp, ctl.Controls(0).Name
    On Error GoTo ErrHandler
    
    ' ========================================
    ' 2. 年份导航
    ' ========================================
    
    ' 年份标签
    Set ctl = CreateControl(sTmp, acLabel, acDetail, "", "", _
                            LM + 45, yNav, 1200, hNav)
    ctl.Name = "lblYear"
    ctl.Caption = "2026年"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.FontWeight = 700
    ctl.ForeColor = CLR_BLACK
    ctl.BackStyle = 0
    ctl.TextAlign = 1
    
    ' 按钮宽度与间距
    Dim navBtnW As Long: navBtnW = 420
    Dim navStartX As Long
    navStartX = FW - LM - navBtnW * 4 - 45     ' 右对齐
    
    ' << 上十年
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            navStartX, yNav, navBtnW, hNav)
    ctl.Name = "cmdPrevDecade"
    ctl.Caption = "<<"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=YMPicker_PrevDecade()"
    
    ' < 上一年
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            navStartX + navBtnW, yNav, navBtnW, hNav)
    ctl.Name = "cmdYMPrevYear"
    ctl.Caption = "<"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=YMPicker_PrevYear()"
    
    ' > 下一年
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            navStartX + navBtnW * 2, yNav, navBtnW, hNav)
    ctl.Name = "cmdYMNextYear"
    ctl.Caption = ">"
    ctl.FontSize = 10
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=YMPicker_NextYear()"
    
    ' >> 下十年
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            navStartX + navBtnW * 3, yNav, navBtnW, hNav)
    ctl.Name = "cmdNextDecade"
    ctl.Caption = ">>"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=YMPicker_NextDecade()"
    
    ' ========================================
    ' 3. 12个月份格子 (3行 × 4列 标签)
    ' ========================================
    Dim iRow As Integer, iCol As Integer, iMonth As Integer
    iMonth = 0
    
    For iRow = 0 To 2
        For iCol = 0 To 3
            iMonth = iMonth + 1
            
            Set ctl = CreateControl(sTmp, acLabel, acDetail, "", "", _
                                    LM + iCol * (MW + MG), _
                                    yGrid + iRow * (MH + MG), _
                                    MW, MH)
            ctl.Name = "lblM" & Format(iMonth, "00")
            ctl.Caption = Format(iMonth, "00") & "月"
            ctl.TextAlign = 2           ' 居中
            ctl.BackStyle = 1           ' 不透明
            ctl.BackColor = CLR_WHITE
            ctl.ForeColor = CLR_BLACK
            ctl.FontSize = 10
            ctl.FontName = "Segoe UI"
            ctl.BorderStyle = 0
            
            ctl.OnClick = "=YMPicker_MonthClick(" & iMonth & ")"
            ctl.OnDblClick = "=YMPicker_MonthDblClick(" & iMonth & ")"
            ctl.OnMouseMove = "=YMPicker_MonthMouseMove(" & iMonth & ")"
        Next iCol
    Next iRow
    
    ' ========================================
    ' 4. 底部按钮
    ' ========================================
    
    ' 清除
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            LM + 45, yBtn, 960, hBtn)
    ctl.Name = "cmdYMClear"
    ctl.Caption = "清除"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=YMPicker_Clear()"
    
    ' 本月
    Set ctl = CreateControl(sTmp, acCommandButton, acDetail, "", "", _
                            FW - LM - 960 - 45, yBtn, 960, hBtn)
    ctl.Name = "cmdThisMonth"
    ctl.Caption = "本月"
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.OnClick = "=YMPicker_GoThisMonth()"
    
    ' ========================================
    ' 保存并重命名窗体
    ' ========================================
    DoCmd.Close acForm, sTmp, acSaveYes
    DoCmd.Rename YM_FORM_NAME, acForm, sTmp
    
    MsgBox "年月选择器窗体 [" & YM_FORM_NAME & "] 创建成功！" & vbCrLf & vbCrLf & _
           "═══ 使用方法 ═══" & vbCrLf & vbCrLf & _
           "  Dim dt As Variant" & vbCrLf & _
           "  dt = ShowYearMonthPicker()" & vbCrLf & vbCrLf & _
           "绑定文本框:" & vbCrLf & _
           "  PickYearMonthFor Me.txtMonth" & vbCrLf & _
           "  PickYearMonthFor Me.txtMonth, ""yyyy-mm""", _
           vbInformation, "创建成功"
    
    Exit Sub

ErrHandler:
    MsgBox "创建窗体失败: " & Err.Description & vbCrLf & _
           "错误号: " & Err.Number, vbCritical, "创建失败"
    On Error Resume Next
    DoCmd.Close acForm, sTmp, acSaveNo
End Sub
