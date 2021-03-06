VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5741
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    Dim lngRow As Long  '表格行号记录
    
    If Not blnLoad Then Exit Sub
    
    '从公共变量或注册表中加载配置信息
    With Me.Grid1
        '窗口控制参数
        lngRow = 2
        .Cell(lngRow, 1).Text = gVar.ParaBlnWindowCloseMin   '关闭时最小化
        .Cell(lngRow, 5).Text = gVar.ParaBlnWindowMinHide    '最小化时隐藏
        .Cell(lngRow + 1, 1).Text = gVar.ParaBlnWindowStartMinC '启动时最小化
        
        '服务端参数
        lngRow = lngRow + 4
        .Cell(lngRow, 3).Text = gVar.TCPSetIP   '要连接服务端IP地址
        .Cell(lngRow, 7).Text = gVar.TCPSetPort  '要连接的服务器端口
        
        '数据库服务器参数
        lngRow = lngRow + 3
        .Cell(lngRow, 3).Text = gVar.ConSource   '服务器名称/IP
        .Cell(lngRow, 7).Text = gVar.ConDatabase '数据库名
        .Cell(lngRow + 2, 3).Text = gVar.ConUserID '登陆名
        .Cell(lngRow + 2, 7).Text = String(Len(gVar.ConPassword), "*") '登陆密码*号显示
        
        '客户端参数
        lngRow = lngRow + 5
        .Cell(lngRow, 1).Text = gVar.ParaBlnAutoStartupAtBoot   '开机自动启动
        .Cell(lngRow, 5).Text = gVar.ParaBlnRememberUserList '记住用户名
        .Cell(lngRow + 1, 1).Text = gVar.ParaBlnRememberUserPassword '记住密码
        .Cell(lngRow + 1, 5).Text = gVar.ParaBlnUserAutoLogin '自动登陆
        
    End With
    
End Sub

Private Sub msSaveParameter(Optional ByVal blnSave As Boolean = True)
    Dim tempVal
    Dim lngRow As Long  '表格行号记录
    
    If Not blnSave Then Exit Sub
    
    '参数值更新至公共变量
    With Grid1
        '窗口控制参数
        lngRow = 2
        gVar.ParaBlnWindowCloseMin = .Cell(lngRow, 1).Text   '关闭时最小化
        gVar.ParaBlnWindowMinHide = .Cell(lngRow, 5).Text    '最小化时隐藏
        gVar.ParaBlnWindowStartMinC = .Cell(lngRow + 1, 1).Text '启动时最小化
        
        '服务端参数
        lngRow = lngRow + 4
        gVar.TCPSetIP = gfCheckIP(.Cell(lngRow, 3).Text) '要连接的服务端IP地址
        tempVal = Val(.Cell(lngRow, 7).Text)                 '要连接的服务器端口
        gVar.TCPSetPort = IIf(tempVal < 10000, gVar.TCPDefaultPort, tempVal)
        
        '数据库服务器参数
        lngRow = lngRow + 3
        '数据库服务器参数只显示，不可修改
        
        
        '客户端参数
        lngRow = lngRow + 5
        gVar.ParaBlnAutoStartupAtBoot = .Cell(lngRow, 1).Text    '开机自动启动
        gVar.ParaBlnRememberUserList = .Cell(lngRow, 5).Text    '记住用户名
        gVar.ParaBlnRememberUserPassword = .Cell(lngRow + 1, 1).Text  '记住密码
        gVar.ParaBlnUserAutoLogin = .Cell(lngRow + 1, 5).Text '自动登陆
        If gVar.ParaBlnRememberUserPassword Then '同时勾选记住用户名
            gVar.ParaBlnRememberUserList = True
        End If
        If gVar.ParaBlnUserAutoLogin Then '同时勾选记住用户名与密码
            gVar.ParaBlnRememberUserList = True
            gVar.ParaBlnRememberUserPassword = True
        End If
        
    End With
    
    '参数值通过公用变量保存进注册表中
    With gVar
        '窗口控制参数
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, IIf(.ParaBlnWindowCloseMin, 1, 0))    '关闭时最小化
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, IIf(.ParaBlnWindowMinHide, 1, 0))  '最小化时隐藏
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowStartMinC, IIf(.ParaBlnWindowStartMinC, 1, 0))  '启动时最小化
        
        '服务端参数
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, .TCPSetPort)  '要连接的服务器端口
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPSetIP) '要连接的服务端IP地址
        
        '数据库服务器参数
        '数据库连接信息只显示不处理
        
        '客户端参数
        If .ParaBlnAutoStartupAtBoot Then   '注册表中添加 开机自动启动 启动项
            .ParaBlnAutoStartupAtBoot = gfStartUpSet(True, RegWrite)
        Else    '注册表中删除启动项
            Call gfStartUpSet(True, RegDelete)
        End If
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, IIf(.ParaBlnAutoStartupAtBoot, 1, 0)) '开机自动启动
        Call SaveSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserList, IIf(.ParaBlnRememberUserList, 1, 0)) '记住用户名
        Call SaveSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserPassword, IIf(.ParaBlnRememberUserPassword, 1, 0)) '记住密码
        Call SaveSetting(.RegAppName, .RegSectionUser, .RegKeyParaUserAutoLogin, IIf(.ParaBlnUserAutoLogin, 1, 0)) '自动登陆
    End With
    
    Call msLoadParameter(True)  '窗口重新加载一次保存后的值
    
    If MsgBox("参数保存完成！是否现在退出窗口？", vbInformation + vbYesNo, "提示") = vbYes Then Unload Me
    
End Sub


Private Sub Form_Load()
    Dim strFile As String
    Dim K As Long, lngSum As Long
    
    Me.Icon = LoadPicture("")
    strFile = gVar.FolderNameBin & "OptionWindowClient.cel"
    If Not gfFileExist(strFile) Then
        MsgBox "以下配置文件加载失败，请解决后再重新打开窗口。" & vbCrLf & strFile, vbCritical, "异常提示"
        Exit Sub
    End If
    With Grid1
        .AutoRedraw = False
        .OpenFile (strFile) '加载模板
        
        .Appearance = Flat
        .Column(0).Width = 0
        .rowHeight(0) = 0
        .ExtendLastCol = True   '扩展最后一列
        .GridColor = vbWhite    '网格线的颜色
        .BorderColor = Me.BackColor '边框的颜色
        .BackColorBkg = Me.BackColor    '空白区域的背景色
        .ReadOnlyFocusRect = Solid  '锁定（只读）单元格所显示的虚框样式
        .DisplayFocusRect = False   '活动单元格是否显示一个虚框
        .SelectionMode = cellSelectionNone  '表格的选择模式
        
        Call msLoadParameter(True) '加载参数值
        
        For K = 0 To .Rows - 1  '计算表格的实际高度
            lngSum = lngSum + .rowHeight(K) * 15    'FC此属性值单位为像素，转成VB的缇要*15.
        Next
        .Height = lngSum    '设置表格高度
        Me.Height = .Top + lngSum + 220 '设置窗口高度
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Grid1.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub Grid1_CellChange(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long
    
    If Not Me.Visible Then Exit Sub
    
    '响应记住密码选择的设置：同时勾选记住用户名
    lngRow = 15 '记住密码参数的行号
    If Row = lngRow And Col = 1 Then
        If Me.Grid1.Cell(Row, Col).Text Then
            Me.Grid1.Cell(lngRow - 1, 5).Text = 1
        End If
    End If
    
    '响应自动登陆选项的设置：同时勾选记住密码与用户名
    If Row = lngRow And Col = 5 Then
        If Me.Grid1.Cell(Row, Col).Text Then
            Me.Grid1.Cell(lngRow - 1, 5).Text = 1
            Me.Grid1.Cell(lngRow, 1).Text = 1
        End If
    End If
    
End Sub

Private Sub Grid1_HyperLinkClick(ByVal Row As Long, ByVal Col As Long, URL As String, Changed As Boolean)
    '保存设置值
    
    URL = ""
    Changed = True
    If Row <> (Grid1.Rows - 1) Then Exit Sub
    
    If Col = 1 Then '保存
        If MsgBox("确定保存所有参数值吗？", vbQuestion + vbOKCancel, "保存询问") = vbOK Then Call msSaveParameter(True)
    ElseIf Col = 5 Then '退出
        Unload Me
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 65 To 90, 97 To 122  '0-9,A-Z,a-z
'            Debug.Print KeyAscii & ":" & Chr(KeyAscii)
        Case Else
            KeyAscii = 0    '密码：限制字母数字以外的输入
    End Select
End Sub

Private Sub Text1_LostFocus()
    Grid1.Cell(11, 7).Text = String(Len(Text1.Text), "*")   '表格只显示等数量的*号
    Text1.Visible = False
End Sub
