VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ��"
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
   StartUpPosition =   1  '����������
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
    
    '�ӹ���������ע����м���������Ϣ
    With Me.Grid1
        .Cell(2, 1).Text = gVar.ParaBlnWindowCloseMin   '�ر�ʱ��С��
        .Cell(2, 5).Text = gVar.ParaBlnWindowMinHide    '��С��ʱ����
        
        .Cell(5, 3).Text = gVar.TCPSetIP   'Ҫ���ӷ����IP��ַ
        .Cell(5, 7).Text = gVar.TCPSetPort  'Ҫ���ӵķ������˿�
        
        .Cell(8, 3).Text = gVar.ConSource   '����������/IP
        .Cell(8, 7).Text = gVar.ConDatabase '���ݿ���
        .Cell(10, 3).Text = gVar.ConUserID  '��½��
        .Cell(10, 7).Text = String(Len(gVar.ConPassword), "*") '��½����*����ʾ
        
        .Cell(13, 1).Text = gVar.ParaBlnAutoStartupAtBoot   '�����Զ�����
        .Cell(13, 5).Text = gVar.ParaBlnRememberUserList '��ס�û���
        .Cell(14, 1).Text = gVar.ParaBlnRememberUserPassword '��ס����
        .Cell(14, 5).Text = gVar.ParaBlnUserAutoLogin '�Զ���½
        
    End With
    
End Sub

Private Sub msSaveParameter(Optional ByVal blnSave As Boolean = True)
    Dim tempVal
    
    If Not blnSave Then Exit Sub
    
    '����ֵ��������������
    With Grid1
        gVar.ParaBlnWindowCloseMin = .Cell(2, 1).Text   '�ر�ʱ��С��
        gVar.ParaBlnWindowMinHide = .Cell(2, 5).Text    '��С��ʱ����
        
        gVar.TCPSetIP = gfCheckIP(.Cell(5, 3).Text) 'Ҫ���ӵķ����IP��ַ
        tempVal = Val(.Cell(5, 7).Text)                 'Ҫ���ӵķ������˿�
        gVar.TCPSetPort = IIf(tempVal < 10000, gVar.TCPDefaultPort, tempVal)
        
        '���ݿ����������ֻ��ʾ�������޸�
                
        gVar.ParaBlnAutoStartupAtBoot = .Cell(13, 1).Text    '�����Զ�����
        gVar.ParaBlnRememberUserList = .Cell(13, 5).Text    '��ס�û���
        gVar.ParaBlnRememberUserPassword = .Cell(14, 1).Text    '��ס����
        gVar.ParaBlnUserAutoLogin = .Cell(14, 5).Text   '�Զ���½
        If gVar.ParaBlnRememberUserPassword Then 'ͬʱ��ѡ��ס�û���
            gVar.ParaBlnRememberUserList = True
        End If
        If gVar.ParaBlnUserAutoLogin Then 'ͬʱ��ѡ��ס�û���������
            gVar.ParaBlnRememberUserList = True
            gVar.ParaBlnRememberUserPassword = True
        End If
        
    End With
    
    '����ֵͨ�����ñ��������ע�����
    With gVar
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, IIf(.ParaBlnWindowCloseMin, 1, 0))    '�ر�ʱ��С��
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, IIf(.ParaBlnWindowMinHide, 1, 0))  '��С��ʱ����
        
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, .TCPSetPort)  'Ҫ���ӵķ������˿�
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPSetIP) 'Ҫ���ӵķ����IP��ַ
        
        '���ݿ�������Ϣֻ��ʾ������
        
        If .ParaBlnAutoStartupAtBoot Then   'ע�������� �����Զ����� ������
            .ParaBlnAutoStartupAtBoot = gfStartUpSet(True, RegWrite)
        Else    'ע�����ɾ��������
            Call gfStartUpSet(True, RegDelete)
        End If
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, IIf(.ParaBlnAutoStartupAtBoot, 1, 0)) '�����Զ�����
        Call SaveSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserList, IIf(.ParaBlnRememberUserList, 1, 0)) '��ס�û���
        Call SaveSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserPassword, IIf(.ParaBlnRememberUserPassword, 1, 0)) '��ס����
        Call SaveSetting(.RegAppName, .RegSectionUser, .RegKeyParaUserAutoLogin, IIf(.ParaBlnUserAutoLogin, 1, 0)) '�Զ���½
    End With
    
    Call msLoadParameter(True)  '�������¼���һ�α�����ֵ
    
    If MsgBox("����������ɣ��Ƿ������˳����ڣ�", vbInformation + vbYesNo, "��ʾ") = vbYes Then Unload Me
    
End Sub


Private Sub Form_Load()
    Dim strFile As String
    
    Me.Icon = LoadPicture("")
    strFile = gVar.FolderNameBin & "OptionWindowClient.cel"
    If Not gfFileExist(strFile) Then
        MsgBox "���������ļ�����ʧ�ܣ������������´򿪴��ڡ�" & vbCrLf & strFile, vbCritical, "�쳣��ʾ"
        Exit Sub
    End If
    With Grid1
        .AutoRedraw = False
        .OpenFile (strFile) '����ģ��
        
        .Appearance = Flat
        .Column(0).Width = 0
        .RowHeight(0) = 0
        .ExtendLastCol = True   '��չ���һ��
        .GridColor = vbWhite    '�����ߵ���ɫ
        .BorderColor = Me.BackColor '�߿����ɫ
        .BackColorBkg = Me.BackColor    '�հ�����ı���ɫ
        .ReadOnlyFocusRect = Solid  '������ֻ������Ԫ������ʾ�������ʽ
        .DisplayFocusRect = False   '���Ԫ���Ƿ���ʾһ�����
        .SelectionMode = cellSelectionNone  '����ѡ��ģʽ
        
        Call msLoadParameter(True)
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub Form_Resize()
    Grid1.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub Grid1_CellChange(ByVal Row As Long, ByVal Col As Long)
    If Not Me.Visible Then Exit Sub
    
    '��Ӧ��ס����ѡ������ã�ͬʱ��ѡ��ס�û���
    If Row = 14 And Col = 1 Then
        If Me.Grid1.Cell(Row, Col).Text Then
            Me.Grid1.Cell(13, 5).Text = 1
        End If
    End If
    
    '��Ӧ�Զ���½ѡ������ã�ͬʱ��ѡ��ס�������û���
    If Row = 14 And Col = 5 Then
        If Me.Grid1.Cell(Row, Col).Text Then
            Me.Grid1.Cell(13, 5).Text = 1
            Me.Grid1.Cell(14, 1).Text = 1
        End If
    End If
    
End Sub

Private Sub Grid1_HyperLinkClick(ByVal Row As Long, ByVal Col As Long, URL As String, Changed As Boolean)
    '��������ֵ
    
    URL = ""
    Changed = True
    If Row <> (Grid1.Rows - 1) Then Exit Sub
    
    If Col = 1 Then '����
        If MsgBox("ȷ���������в���ֵ��", vbQuestion + vbOKCancel, "����ѯ��") = vbOK Then Call msSaveParameter(True)
    ElseIf Col = 5 Then '�˳�
        Unload Me
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 65 To 90, 97 To 122  '0-9,A-Z,a-z
'            Debug.Print KeyAscii & ":" & Chr(KeyAscii)
        Case Else
            KeyAscii = 0    '���룺������ĸ�������������
    End Select
End Sub

Private Sub Text1_LostFocus()
    Grid1.Cell(10, 7).Text = String(Len(Text1.Text), "*")   '���ֻ��ʾ��������*��
    Text1.Visible = False
End Sub
