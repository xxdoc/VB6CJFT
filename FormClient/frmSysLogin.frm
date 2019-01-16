VERSION 5.00
Begin VB.Form frmSysLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   Icon            =   "frmSysLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5295
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008000&
      Cancel          =   -1  'True
      Caption         =   "取消自动登陆(5秒)"
      Height          =   300
      Left            =   1410
      TabIndex        =   8
      Top             =   1845
      Width           =   2730
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登陆"
      Default         =   -1  'True
      Height          =   375
      Left            =   1410
      TabIndex        =   2
      Top             =   2160
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1770
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎登陆系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   2
      Left            =   1200
      TabIndex        =   7
      Top             =   240
      Width           =   2715
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   915
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   1500
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   6330
      Left            =   120
      Picture         =   "frmSysLogin.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12000
   End
End
Attribute VB_Name = "frmSysLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mconCount As Integer = 5 '自动登陆时的倒计时时间，单位秒


Private Function mfInputCheck(ByVal strName As String, strPWD As String) As Boolean
    '输入验证
    
    Dim strCheck As String
    
    strName = Trim(strName)
    If Len(strName) = 0 Then '空值检测
        MsgBox "用户名不能为空!", vbCritical, "用户名警告"
        Combo1.SetFocus
        Exit Function
    End If
    If Len(strPWD) = 0 Then
        MsgBox "密码不能为空!", vbCritical, "密码警告"
        Text1.SetFocus
        Exit Function
    End If
    
    strCheck = gfStringCheck(strName)
    If Len(strCheck) > 0 Then
        MsgBox "用户名中不能含有特殊字符【" & strCheck & "】!", vbCritical, "用户名警告"
        Combo1.SetFocus
        Exit Function
    End If
    strCheck = gfStringCheck(strPWD) '检查是否包含特殊字符
    If Len(strCheck) > 0 Then
        MsgBox "密码中不能含有特殊字符【" & strCheck & "】!", vbCritical, "密码警告"
        Text1.Text = ""
        Text1.SetFocus
        Exit Function
    End If
    
    mfInputCheck = True '验证没问题后返回真值
    
End Function

Private Function mfLoginCheck(ByVal strName As String, strPWD As String) As Boolean
    '登陆验证
    
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    
    strPWD = EncryptString(strPWD, gVar.EncryptKey) '加密密码,再网络传输
    strSQL = "EXEC sp_FT_Sys_UserLogin '" & strName & "','" & strPWD & "'" '生成SQL语句,参数中的单引号[']不可少
    Rem Debug.Print gVar.ConString
    Rem Debug.Print strSQL
    
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then '该账号密码未检索到记录时
        gWind.Winsock1.Item(1).Close
        MsgBox "用户名 或 密码错误！", vbExclamation, "输入错误警告"
        Combo1.SetFocus
        GoTo LineEnd
    End If
    If rsUser.RecordCount > 1 Then '同一账号密码检索出几条记录时
        gWind.Winsock1.Item(1).Close
        MsgBox "该用户名在系统中无法识别，请联系管理员！", vbExclamation, "系统警告"
        GoTo LineEnd
    End If
    
    With gVar '用户信息保存在公用变量中
        .UserLoginName = strName
        .UserPassword = strPWD
        .UserFullName = "" & rsUser.Fields("UserFullName")
        .UserAutoID = "" & rsUser.Fields("UserAutoID")
        .UserDepartment = "" & rsUser.Fields("DeptID")
        Rem Debug.Print .UserLoginIP,.UserComputerName '在主窗体加载参数函数中已获取
    End With
    
    mfLoginCheck = True '验证没问题后函数返回真值
    
    Call msSaveUserInfo(True) '登陆成功则保存用户信息进注册表中
    gVar.ClientLoginCheckOver = True '检验完成标志
    
LineEnd:
    If Not rsUser Is Nothing Then If rsUser.State <> adStateClosed Then rsUser.Close
    Set rsUser = Nothing
    
End Function

Private Sub msLoadUserInfo(Optional ByVal blnLoad As Boolean = True)
    '加载登陆过的用户信息
    
    Dim strReg As String, arrUser() As String, strLast As String, strPWDde As String
    Dim K As Long, C As Long
    
    '加载用户名列表
    strReg = GetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, "") '获取列表
    If Len(strReg) = 0 Then Exit Sub    '没有保存的用户名则退出
    arrUser = Split(strReg, gVar.CmdLineSeparator) '分解列表
    C = UBound(arrUser)
    Combo1.Clear
    For K = 0 To C
        Combo1.AddItem Trim(arrUser(K)) '将每个用户名加载进下拉列表中
    Next
    
    '加载最近登陆的用户名
    strLast = GetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, "")
    Combo1.Text = Trim(strLast)
    
    '如果勾选了记住密码，则自动填充对应密码
    Call msLoadPassword(strLast) '加载密码
    
End Sub

Private Sub msLoadPassword(ByVal strName As String)
    '加载指定账号的密码
    Dim strPWDde As String
    
    If gVar.ParaBlnRememberUserPassword And Len(strName) > 0 Then
        strPWDde = GetSetting(gVar.RegAppName, gVar.RegSectionUser, strName, "")
        If Len(strPWDde) > 0 Then
            On Error Resume Next    '密文异常时可能报错
            Text1.Text = DecryptString(strPWDde, gVar.EncryptKey)
            If Err.Number <> 0 Then
                Call gsAlarmAndLog("密码被破坏警报")
                Text1.Text = "" '清空密码框
            End If
        Else
            Text1.Text = "" '清空密码框
        End If
    End If
End Sub

Private Sub msEnabledSet(ByVal blnSet As Boolean)
    '自动登模式下，某些控件的Enabled属性设置为False
    
    Command1.Enabled = blnSet   '登陆按钮
'    Combo1.Enabled = blnSet     '用户名
'    Text1.Text = blnSet         '密码
    Label3.Enabled = blnSet     '设置
End Sub

Private Sub msSaveUserInfo(Optional ByVal blnSave As Boolean = True)
    '保存登陆过的用户信息
    Dim strCurUser As String, strList As String, strCombo As String
    Dim K As Long, C As Long
    
    '用户列表处理
    If gVar.ParaBlnRememberUserList Then '保存用户列表
        strCurUser = Trim(Combo1.Text)
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, strCurUser) '记录最近登陆过的用户名
        strList = strCurUser '当前登陆用户总是排在列表第一位
        C = Combo1.ListCount
        If C > 0 Then '下拉中有其他用户名
            strCurUser = LCase(strCurUser)
            C = C - 1
            For K = 0 To C '生成新顺序的用户名列表
                strCombo = LCase(Trim(Combo1.List(K)))
                If strCombo <> strCurUser Then
                    strList = strList & gVar.CmdLineSeparator & strCombo '两个名之间用分隔符gVar.CmdLineSeparator
                End If
            Next
        End If
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, strList) '保存用户列表到注册表中
    Else '清除注册表中的用户
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, "最新登陆用户名删除异常") '删除最近登陆用户
        End If
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, "用户名记录列表删除异常") '删除用户列表
        End If
    End If
    
    '记住密码处理
    strCurUser = Trim(Combo1.Text)
    If gVar.ParaBlnRememberUserPassword Then '加密保存密码
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser, EncryptString(Text1.Text, gVar.EncryptKey))
    Else '删除密码
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser, "用户" & strCurUser & "记住密码删除异常")
        End If
    End If
    
End Sub

Private Sub Combo1_Click()
    '如果有记住密码，自动加载保存过的密码
    Dim strName As String
    
    If gVar.ParaBlnRememberUserPassword Then '勾选了记住密码
        strName = Trim(Combo1.Text) '获取用户名
        Call msLoadPassword(strName) '加载密码
    End If
End Sub

Private Sub Command1_Click()
    '登陆系统
    Dim strName As String, strPWD As String
    
    strName = Trim(Combo1.Text) '获取用户名
    strPWD = Text1.Text '获取密码
    If Not mfInputCheck(strName, strPWD) Then
        Call msEnabledSet(True) '激活控件
        Exit Sub '输入不规范则退出过程
    End If
    
    Call gsConnectToServer(gWind.Winsock1.Item(1), True)      '与务器建立连接。连接成功后则自动校验账号密码.
    Timer2.Enabled = True '激活连接等待与自动校验
    Me.MousePointer = 13
End Sub

Private Sub Command2_Click()
    '退出程序
    
    gVar.UnloadFromLogin = True '不能直接使用卸载语句，会报警，权且通过此公共变量并在主窗体中的计时器来实现
End Sub

Private Sub Command3_Click()
    '取消自动登陆
    
    gVar.ClientCancelAutoLogin = True   '取消标志
    Command3.Visible = False    '隐藏取消按钮
    Call msEnabledSet(True) '激活控件
End Sub

Private Sub Form_Load()
    '加载窗体
    
    Me.Icon = gWind.Icon
    Timer1.Enabled = False
    Timer1.Interval = 1000 '只能设1秒
    Timer2.Enabled = False
    Timer2.Interval = 1000 '只能设1秒
    Command3.Visible = False
    
    If gVar.ParaBlnRememberUserList Then
        Call msLoadUserInfo(True) '加载用户信息
    End If
    
    If gVar.ParaBlnUserAutoLogin Then '本判断语句后不要再有其它语句
        If Len(Trim(Combo1.Text)) > 0 And Len(Text1.Text) > 0 Then '勾选了自动登陆且账号密码不为空时
            Timer1.Enabled = True '自动登陆.本想在此触发，但窗体卸载会报警，权且通过计时器触发。
            Call msEnabledSet(False) '禁用控件
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    '背景图片自适应窗口大小
    
    On Error Resume Next
    Image1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '点击窗口关闭按钮时触发
    
    gVar.UnloadFromLogin = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label3.FontUnderline Then    '复原样式
        Label3.FontUnderline = False '去除下划线
        Label3.ForeColor = vbBlack  '字体黑色
    End If
End Sub

Private Sub Label3_Click()
    '弹出设置窗口
    Call gsOpenTheWindow("frmOption", vbModal, vbNormal)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '通API函数鼠标指针变手指
    
    Dim hHandCursor As Long
    
    Label3.FontUnderline = True '显示下划线
    Label3.ForeColor = vbRed '字体红色
    hHandCursor = LoadCursor(0, IDC_HAND) '调用API载入光标
    Call SetCursor(hHandCursor) '调用API使指针变手指状
End Sub

Private Sub Timer1_Timer()
    '自动登陆。放在Form_Load事件中会报错，暂转移至此。
    Static ReverseCount As Integer
    
    If gVar.ParaBlnUserAutoLogin And (Not gVar.ClientCancelAutoLogin) Then
        ReverseCount = ReverseCount + 1 '计时
        Command3.Visible = True '显示取消按钮
        Command3.Caption = "取消自动登陆(" & CStr(mconCount - ReverseCount) & "秒)"
        If ReverseCount = mconCount Then
            ReverseCount = 0    '清零静态变量
            Timer1.Enabled = False '关闭计时器
            Command1.Value = 1 '相当于自动点击登陆按钮，触发该按钮的click事件
            Command3.Visible = False '隐藏取消按钮
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    '触发校验账号密码
    
    If gVar.RestoreDBInfoOver Then '已接收到数据库的连接信息
        Timer2.Enabled = False '关闭计时器
        If mfLoginCheck(Trim(Combo1.Text), Text1.Text) Then '校验账号密码
            gVar.ClientLoginCheckOver = True    '校验通过标志
            Me.MousePointer = 0
        Else
            Call msEnabledSet(True) '激活控件
        End If
    End If
End Sub
