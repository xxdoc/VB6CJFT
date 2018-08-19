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
    
    If Not blnLoad Then Exit Sub
    
    '从公共变量或注册表中加载配置信息
    With Me.Grid1
        .Cell(2, 1).Text = gVar.ParaBlnWindowCloseMin   '关闭时最小化
        .Cell(2, 5).Text = gVar.ParaBlnWindowMinHide    '最小化时隐藏
        
        .Cell(5, 3).Text = gVar.TCPSetPort  '侦听端口
        
        .Cell(8, 3).Text = gVar.ConSource   '服务器名称/IP
        .Cell(8, 7).Text = gVar.ConDatabase '数据库名
        .Cell(10, 3).Text = gVar.ConUserID  '登陆名
        Text1.Text = gVar.ConPassword       '登陆密码
        .Cell(10, 7).Text = String(Len(gVar.ConPassword), "*") '登陆密码*号显示
        
        
    End With
    
End Sub

Private Sub msSaveParameter(Optional ByVal blnSave As Boolean = True)
    Dim TempVal
    
    If Not blnSave Then Exit Sub
    
    '参数值更新至公共变量
    With Grid1
        gVar.ParaBlnWindowCloseMin = .Cell(2, 1).Text   '关闭时最小化
        gVar.ParaBlnWindowMinHide = .Cell(2, 5).Text    '最小化时隐藏
        
        TempVal = Val(.Cell(5, 3).Text)                 '侦听端口
        gVar.TCPSetPort = IIf(TempVal < 10000, gVar.TCPDefaultPort, TempVal)
        
        gVar.ConSource = gfCheckIP(Trim(.Cell(8, 3).Text))    '服务器名称/IP
        gVar.ConDatabase = Trim(.Cell(8, 7).Text)   '数据库名
        gVar.ConUserID = Trim(.Cell(10, 3).Text)    '登陆名
        gVar.ConPassword = Text1.Text               '登陆密码
        
        
    End With
    
    '参数值通过公用变量保存进注册表中
    With gVar
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, IIf(.ParaBlnWindowCloseMin, 1, 0))    '关闭时最小化
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, IIf(.ParaBlnWindowMinHide, 1, 0))  '最小化时隐藏
        
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, .TCPSetPort)  '侦听端口
        
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .ConSource)
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString(.ConDatabase, .EncryptKey)) '数据库名
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString(.ConUserID, .EncryptKey)) '登陆名
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString(.ConPassword, .EncryptKey)) '登陆密码
        
    End With
    
    Call msLoadParameter(True)  '窗口重新加载一次保存后的值
    
    If MsgBox("参数保存完成！是否现在退出窗口？", vbInformation + vbYesNo, "提示") = vbYes Then Unload Me
    
End Sub


Private Sub Form_Load()
    Dim strFile As String
    
    strFile = gVar.FolderNameBin & "OptionWindow.cel"
    If Not gfFileExist(strFile) Then
        MsgBox "以下配置文件加载失败，请解决后再重新打开窗口。" & vbCrLf & strFile, vbCritical, "异常提示"
        Exit Sub
    End If
    With Grid1
        .AutoRedraw = False
        .OpenFile (strFile) '加载模板
        
        .Appearance = Flat
        .Column(0).Width = 0
        .RowHeight(0) = 0
        .ExtendLastCol = True   '扩展最后一列
        .GridColor = vbWhite    '网格线的颜色
        .BorderColor = Me.BackColor '边框的颜色
        .BackColorBkg = Me.BackColor    '空白区域的背景色
        .ReadOnlyFocusRect = Solid  '锁定（只读）单元格所显示的虚框样式
        .DisplayFocusRect = False   '活动单元格是否显示一个虚框
        .SelectionMode = cellSelectionNone  '表格的选择模式
        
        Call msLoadParameter(True)
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub Form_Resize()
    Grid1.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub Grid1_Click()
    With Grid1.ActiveCell
        If .Row = 10 And .Col = 7 Then  '密码单元格借用TextBox控件处理成星号*
            Text1.Move .Left * 15 + 100, .Top * 15 + 100, .Width * 15, .Height * 15
            With Text1
                .Visible = True
                .ZOrder
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    End With
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
    Grid1.Cell(10, 7).Text = String(Len(Text1.Text), "*")   '表格只显示等数量的*号
    Text1.Visible = False
End Sub
