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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008000&
      Cancel          =   -1  'True
      Caption         =   "ȡ���Զ���½(5��)"
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
      Caption         =   "��½"
      Default         =   -1  'True
      Height          =   375
      Left            =   1410
      TabIndex        =   2
      Top             =   2160
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��ӭ��½ϵͳ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�û���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
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

Private Const mconCount As Integer = 5 '�Զ���½ʱ�ĵ���ʱʱ�䣬��λ��


Private Function mfInputCheck(ByVal strName As String, strPWD As String) As Boolean
    '������֤
    
    Dim strCheck As String
    
    strName = Trim(strName)
    If Len(strName) = 0 Then '��ֵ���
        MsgBox "�û�������Ϊ��!", vbCritical, "�û�������"
        Combo1.SetFocus
        Exit Function
    End If
    If Len(strPWD) = 0 Then
        MsgBox "���벻��Ϊ��!", vbCritical, "���뾯��"
        Text1.SetFocus
        Exit Function
    End If
    
    strCheck = gfStringCheck(strName)
    If Len(strCheck) > 0 Then
        MsgBox "�û����в��ܺ��������ַ���" & strCheck & "��!", vbCritical, "�û�������"
        Combo1.SetFocus
        Exit Function
    End If
    strCheck = gfStringCheck(strPWD) '����Ƿ���������ַ�
    If Len(strCheck) > 0 Then
        MsgBox "�����в��ܺ��������ַ���" & strCheck & "��!", vbCritical, "���뾯��"
        Text1.Text = ""
        Text1.SetFocus
        Exit Function
    End If
    
    mfInputCheck = True '��֤û����󷵻���ֵ
    
End Function

Private Function mfLoginCheck(ByVal strName As String, strPWD As String) As Boolean
    '��½��֤
    
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    
    strPWD = EncryptString(strPWD, gVar.EncryptKey) '��������,�����紫��
    strSQL = "EXEC sp_FT_Sys_UserLogin '" & strName & "','" & strPWD & "'" '����SQL���,�����еĵ�����[']������
    Rem Debug.Print gVar.ConString
    Rem Debug.Print strSQL
    
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then '���˺�����δ��������¼ʱ
        gWind.Winsock1.Item(1).Close
        MsgBox "�û��� �� �������", vbExclamation, "������󾯸�"
        Combo1.SetFocus
        GoTo LineEnd
    End If
    If rsUser.RecordCount > 1 Then 'ͬһ�˺����������������¼ʱ
        gWind.Winsock1.Item(1).Close
        MsgBox "���û�����ϵͳ���޷�ʶ������ϵ����Ա��", vbExclamation, "ϵͳ����"
        GoTo LineEnd
    End If
    
    With gVar '�û���Ϣ�����ڹ��ñ�����
        .UserLoginName = strName
        .UserPassword = strPWD
        .UserFullName = "" & rsUser.Fields("UserFullName")
        .UserAutoID = "" & rsUser.Fields("UserAutoID")
        .UserDepartment = "" & rsUser.Fields("DeptID")
        Rem Debug.Print .UserLoginIP,.UserComputerName '����������ز����������ѻ�ȡ
    End With
    
    mfLoginCheck = True '��֤û�������������ֵ
    
    Call msSaveUserInfo(True) '��½�ɹ��򱣴��û���Ϣ��ע�����
    gVar.ClientLoginCheckOver = True '������ɱ�־
    
LineEnd:
    If Not rsUser Is Nothing Then If rsUser.State <> adStateClosed Then rsUser.Close
    Set rsUser = Nothing
    
End Function

Private Sub msLoadUserInfo(Optional ByVal blnLoad As Boolean = True)
    '���ص�½�����û���Ϣ
    
    Dim strReg As String, arrUser() As String, strLast As String, strPWDde As String
    Dim K As Long, C As Long
    
    '�����û����б�
    strReg = GetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, "") '��ȡ�б�
    If Len(strReg) = 0 Then Exit Sub    'û�б�����û������˳�
    arrUser = Split(strReg, gVar.CmdLineSeparator) '�ֽ��б�
    C = UBound(arrUser)
    Combo1.Clear
    For K = 0 To C
        Combo1.AddItem Trim(arrUser(K)) '��ÿ���û������ؽ������б���
    Next
    
    '���������½���û���
    strLast = GetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, "")
    Combo1.Text = Trim(strLast)
    
    '�����ѡ�˼�ס���룬���Զ�����Ӧ����
    Call msLoadPassword(strLast) '��������
    
End Sub

Private Sub msLoadPassword(ByVal strName As String)
    '����ָ���˺ŵ�����
    Dim strPWDde As String
    
    If gVar.ParaBlnRememberUserPassword And Len(strName) > 0 Then
        strPWDde = GetSetting(gVar.RegAppName, gVar.RegSectionUser, strName, "")
        If Len(strPWDde) > 0 Then
            On Error Resume Next    '�����쳣ʱ���ܱ���
            Text1.Text = DecryptString(strPWDde, gVar.EncryptKey)
            If Err.Number <> 0 Then
                Call gsAlarmAndLog("���뱻�ƻ�����")
                Text1.Text = "" '��������
            End If
        Else
            Text1.Text = "" '��������
        End If
    End If
End Sub

Private Sub msEnabledSet(ByVal blnSet As Boolean)
    '�Զ���ģʽ�£�ĳЩ�ؼ���Enabled��������ΪFalse
    
    Command1.Enabled = blnSet   '��½��ť
'    Combo1.Enabled = blnSet     '�û���
'    Text1.Text = blnSet         '����
    Label3.Enabled = blnSet     '����
End Sub

Private Sub msSaveUserInfo(Optional ByVal blnSave As Boolean = True)
    '�����½�����û���Ϣ
    Dim strCurUser As String, strList As String, strCombo As String
    Dim K As Long, C As Long
    
    '�û��б���
    If gVar.ParaBlnRememberUserList Then '�����û��б�
        strCurUser = Trim(Combo1.Text)
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, strCurUser) '��¼�����½�����û���
        strList = strCurUser '��ǰ��½�û����������б��һλ
        C = Combo1.ListCount
        If C > 0 Then '�������������û���
            strCurUser = LCase(strCurUser)
            C = C - 1
            For K = 0 To C '������˳����û����б�
                strCombo = LCase(Trim(Combo1.List(K)))
                If strCombo <> strCurUser Then
                    strList = strList & gVar.CmdLineSeparator & strCombo '������֮���÷ָ���gVar.CmdLineSeparator
                End If
            Next
        End If
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, strList) '�����û��б�ע�����
    Else '���ע����е��û�
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, "���µ�½�û���ɾ���쳣") 'ɾ�������½�û�
        End If
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, "�û�����¼�б�ɾ���쳣") 'ɾ���û��б�
        End If
    End If
    
    '��ס���봦��
    strCurUser = Trim(Combo1.Text)
    If gVar.ParaBlnRememberUserPassword Then '���ܱ�������
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser, EncryptString(Text1.Text, gVar.EncryptKey))
    Else 'ɾ������
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser, "�û�" & strCurUser & "��ס����ɾ���쳣")
        End If
    End If
    
End Sub

Private Sub Combo1_Click()
    '����м�ס���룬�Զ����ر����������
    Dim strName As String
    
    If gVar.ParaBlnRememberUserPassword Then '��ѡ�˼�ס����
        strName = Trim(Combo1.Text) '��ȡ�û���
        Call msLoadPassword(strName) '��������
    End If
End Sub

Private Sub Command1_Click()
    '��½ϵͳ
    Dim strName As String, strPWD As String
    
    strName = Trim(Combo1.Text) '��ȡ�û���
    strPWD = Text1.Text '��ȡ����
    If Not mfInputCheck(strName, strPWD) Then
        Call msEnabledSet(True) '����ؼ�
        Exit Sub '���벻�淶���˳�����
    End If
    
    Call gsConnectToServer(gWind.Winsock1.Item(1), True)      '�������������ӡ����ӳɹ������Զ�У���˺�����.
    Timer2.Enabled = True '�������ӵȴ����Զ�У��
    Me.MousePointer = 13
End Sub

Private Sub Command2_Click()
    '�˳�����
    
    gVar.UnloadFromLogin = True '����ֱ��ʹ��ж����䣬�ᱨ����Ȩ��ͨ���˹������������������еļ�ʱ����ʵ��
End Sub

Private Sub Command3_Click()
    'ȡ���Զ���½
    
    gVar.ClientCancelAutoLogin = True   'ȡ����־
    Command3.Visible = False    '����ȡ����ť
    Call msEnabledSet(True) '����ؼ�
End Sub

Private Sub Form_Load()
    '���ش���
    
    Me.Icon = gWind.Icon
    Timer1.Enabled = False
    Timer1.Interval = 1000 'ֻ����1��
    Timer2.Enabled = False
    Timer2.Interval = 1000 'ֻ����1��
    Command3.Visible = False
    
    If gVar.ParaBlnRememberUserList Then
        Call msLoadUserInfo(True) '�����û���Ϣ
    End If
    
    If gVar.ParaBlnUserAutoLogin Then '���ж�����Ҫ�����������
        If Len(Trim(Combo1.Text)) > 0 And Len(Text1.Text) > 0 Then '��ѡ���Զ���½���˺����벻Ϊ��ʱ
            Timer1.Enabled = True '�Զ���½.�����ڴ˴�����������ж�ػᱨ����Ȩ��ͨ����ʱ��������
            Call msEnabledSet(False) '���ÿؼ�
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    '����ͼƬ����Ӧ���ڴ�С
    
    On Error Resume Next
    Image1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '������ڹرհ�ťʱ����
    
    gVar.UnloadFromLogin = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label3.FontUnderline Then    '��ԭ��ʽ
        Label3.FontUnderline = False 'ȥ���»���
        Label3.ForeColor = vbBlack  '�����ɫ
    End If
End Sub

Private Sub Label3_Click()
    '�������ô���
    Call gsOpenTheWindow("frmOption", vbModal, vbNormal)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ͨAPI�������ָ�����ָ
    
    Dim hHandCursor As Long
    
    Label3.FontUnderline = True '��ʾ�»���
    Label3.ForeColor = vbRed '�����ɫ
    hHandCursor = LoadCursor(0, IDC_HAND) '����API������
    Call SetCursor(hHandCursor) '����APIʹָ�����ָ״
End Sub

Private Sub Timer1_Timer()
    '�Զ���½������Form_Load�¼��лᱨ����ת�����ˡ�
    Static ReverseCount As Integer
    
    If gVar.ParaBlnUserAutoLogin And (Not gVar.ClientCancelAutoLogin) Then
        ReverseCount = ReverseCount + 1 '��ʱ
        Command3.Visible = True '��ʾȡ����ť
        Command3.Caption = "ȡ���Զ���½(" & CStr(mconCount - ReverseCount) & "��)"
        If ReverseCount = mconCount Then
            ReverseCount = 0    '���㾲̬����
            Timer1.Enabled = False '�رռ�ʱ��
            Command1.Value = 1 '�൱���Զ������½��ť�������ð�ť��click�¼�
            Command3.Visible = False '����ȡ����ť
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    '����У���˺�����
    
    If gVar.RestoreDBInfoOver Then '�ѽ��յ����ݿ��������Ϣ
        Timer2.Enabled = False '�رռ�ʱ��
        If mfLoginCheck(Trim(Combo1.Text), Text1.Text) Then 'У���˺�����
            gVar.ClientLoginCheckOver = True    'У��ͨ����־
            Me.MousePointer = 0
        Else
            Call msEnabledSet(True) '����ؼ�
        End If
    End If
End Sub
