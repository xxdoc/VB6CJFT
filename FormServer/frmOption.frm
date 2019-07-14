VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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

Private Const mconstrTip As String = "ѡ��һ���ļ���" 'ѡ��һ���ļ��е���򿪼���


Private Function mfCheckFolder(ByVal strFolder As String) As String
    '��ѡ����ļ��������
    Dim strCheck As String, strEnd As String
    
    strCheck = Trim(strFolder)
    If Len(strCheck) = 0 Then
        strCheck = gVar.FolderNameBackup
    ElseIf LCase(strFolder) = LCase(mconstrTip) Then    'δѡ��·��������ȡ��
        strCheck = ""
    Else
        strEnd = Mid(strCheck, InStrRev(strCheck, "\") + 1)
        If LCase(strEnd) = LCase(mconstrTip) Then
            strCheck = Left(strCheck, InStrRev(strCheck, "\"))
        End If
        If Not gfFolderRepair(strCheck) Then
            strCheck = gVar.FolderNameBackup
        End If
    End If
    mfCheckFolder = strCheck
End Function

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    Dim lngRow As Long  '����кż�¼
    
    If Not blnLoad Then Exit Sub
    
    '�ӹ���������ע����м���������Ϣ
    With Me.Grid1
        '���ڿ��Ʋ���
        lngRow = 2
        .Cell(lngRow, 1).Text = gVar.ParaBlnWindowCloseMin   '�ر�ʱ��С��
        .Cell(lngRow, 5).Text = gVar.ParaBlnWindowMinHide    '��С��ʱ����
        .Cell(lngRow + 1, 1).Text = gVar.ParaBlnWindowStartMinS '����ʱ��С��
        
        '����˲���
        lngRow = lngRow + 4
        .Cell(lngRow, 3).Text = gVar.TCPSetPort  '�����˿�
        .Cell(lngRow, 5).Text = gVar.ParaBlnAutoStartupAtBoot   '�����Զ�����
        
        '���ݿ����������
        lngRow = lngRow + 4
        .Cell(lngRow, 3).Text = gVar.ConSource   '����������/IP
        .Cell(lngRow, 7).Text = gVar.ConDatabase '���ݿ���
        .Cell(lngRow + 2, 3).Text = gVar.ConUserID '��½��
        Text1.Text = gVar.ConPassword       '��½����
        .Cell(lngRow + 2, 7).Text = String(Len(gVar.ConPassword), "*") '��½����*����ʾ
        
        '�ͻ��˿��Ʋ���
        lngRow = lngRow + 5
        .Cell(lngRow, 1).Text = gVar.ParaBlnLimitClientConnect '���ƿͻ�������
        .Cell(lngRow, 7).Text = gVar.ParaLimitClientConnectTime '���ƿͻ�������ʱ��
        .Cell(lngRow + 1, 3).Text = gVar.TCPConnectMax '���ƿͻ���������
        
        '������ļ����ݲ���
        lngRow = lngRow + 4
        .Cell(lngRow, 3).Text = gVar.ParaBackupStore '����·��
    End With
    
End Sub

Private Sub msSaveParameter(Optional ByVal blnSave As Boolean = True)
    Dim lngRow As Long  '����кż�¼
    Dim tempVal
    
    If Not blnSave Then Exit Sub
    
    '����ֵ��������������
    With Grid1
        '���ڿ��Ʋ���
        lngRow = 2
        gVar.ParaBlnWindowCloseMin = .Cell(lngRow, 1).Text   '�ر�ʱ��С��
        gVar.ParaBlnWindowMinHide = .Cell(lngRow, 5).Text    '��С��ʱ����
        gVar.ParaBlnWindowStartMinS = .Cell(lngRow + 1, 1).Text  '����ʱ��С��
        
        '����˲���
        lngRow = lngRow + 4
        tempVal = Val(.Cell(lngRow, 3).Text)                 '�����˿�
        gVar.TCPSetPort = IIf(tempVal < 10000, gVar.TCPDefaultPort, tempVal)
        gVar.ParaBlnAutoStartupAtBoot = .Cell(lngRow, 5).Text    '�����Զ�����
        
        '���ݿ����������
        lngRow = lngRow + 4
        gVar.ConSource = gfCheckIP(Trim(.Cell(lngRow, 3).Text))    '����������/IP
        gVar.ConDatabase = Trim(.Cell(lngRow, 7).Text)   '���ݿ���
        gVar.ConUserID = Trim(.Cell(lngRow + 2, 3).Text)  '��½��
        gVar.ConPassword = Text1.Text               '��½����
        
        '�ͻ��˿��Ʋ���
        lngRow = lngRow + 5
        gVar.ParaBlnLimitClientConnect = .Cell(lngRow, 1).Text '���ƿͻ�������
        tempVal = Val(.Cell(lngRow, 7).Text)
        gVar.ParaLimitClientConnectTime = IIf(tempVal < 1 Or tempVal > 60, 30, tempVal) '���ƿͻ�������ʱ��
        tempVal = Val(.Cell(lngRow + 1, 3).Text)
        gVar.TCPConnectMax = IIf(tempVal < 1 Or tempVal > 20, 2, tempVal) '���ƿͻ���������
        
        '������ļ����ݲ���
        lngRow = lngRow + 4
        gVar.ParaBackupStore = mfCheckFolder(.Cell(lngRow, 3).Text) '����·��
    End With
    
    '����ֵͨ�����ñ��������ע�����
    With gVar
        '���ڿ��Ʋ���
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, IIf(.ParaBlnWindowCloseMin, 1, 0))    '�ر�ʱ��С��
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, IIf(.ParaBlnWindowMinHide, 1, 0))  '��С��ʱ����
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowStartMinS, IIf(.ParaBlnWindowStartMinS, 1, 0)) '����ʱ��С��
        
        '����˲���
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, .TCPSetPort)  '�����˿�
        If .ParaBlnAutoStartupAtBoot Then   'ע��������������
            .ParaBlnAutoStartupAtBoot = gfStartUpSet(True, RegWrite)
        Else    'ע�����ɾ��������
            Call gfStartUpSet(True, RegDelete)
        End If
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, IIf(.ParaBlnAutoStartupAtBoot, 1, 0)) '�����Զ�����
        
        '���ݿ����������
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .ConSource)
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString(.ConDatabase, .EncryptKey)) '���ݿ���
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString(.ConUserID, .EncryptKey)) '��½��
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString(.ConPassword, .EncryptKey)) '��½����
        
        '�ͻ��˿��Ʋ���
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnect, IIf(.ParaBlnLimitClientConnect, 1, 0)) '���ƿͻ�������
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectTime, .ParaLimitClientConnectTime) '���ƿͻ�������ʱ��
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectNumber, .TCPConnectMax) '���ƿͻ���������
        
        '������ļ����ݲ���
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackStore, .ParaBackupStore) '����·��
    End With
    
    Call msLoadParameter(True)  '�������¼���һ�α�����ֵ
    
    If MsgBox("����������ɣ��Ƿ������˳����ڣ�", vbInformation + vbYesNo, "��ʾ") = vbYes Then Unload Me
    
End Sub


Private Sub Form_Load()
    Dim strFile As String
    Dim K As Long, lngSum As Long
    
    Me.Icon = LoadPicture("")
    strFile = gVar.FolderNameBin & "OptionWindowServer.cel"
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
        
        Call msLoadParameter(True)  '���ز���ֵ
        
        For K = 0 To .Rows - 1  '�������ʵ�ʸ߶�
            lngSum = lngSum + .RowHeight(K) * 15    'FC������ֵ��λΪ���أ�ת��VB���Ҫ*15.
        Next
        .Height = lngSum    '���ñ��߶�
        Me.Height = .Top + lngSum + 220 '���ô��ڸ߶�
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Grid1.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub Grid1_ButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strPath As String
    Dim lngRow As Long, lngCol As Long
    
    On Error Resume Next
    
    lngRow = 19 '������������к�
    lngCol = 3  '������������к�
    
    If Row = lngRow And Col = lngCol Then    'ѡ���ļ�����·��
'''        With CommonDialog1   'Ȩ��֮���÷���������
'''            .DialogTitle = "����·��ѡ��"
'''            .Flags = cdlOFNPathMustExist  '·�������������Ч cdlOFNCreatePrompt=cdlOFNFileMustExist + cdlOFNPathMustExist
'''            .InitDir = IIf(Len(Grid1.Cell(lngRow, lngCol).Text) > 0, Grid1.Cell(lngRow, lngCol).Text, gVar.FolderNameBackup)
'''            .FileName = mconstrTip
'''            .ShowOpen
'''            strPath = mfCheckFolder(.FileName)
'''            If Len(strPath) > 0 Then
'''                If Not Right(strPath, 1) = "\" Then strPath = strPath & "\"
'''                Grid1.Cell(lngRow, lngCol).Text = strPath
'''            End If
'''        End With
        
        strPath = BrowseForFolder(Me, Grid1.Cell(lngRow, lngCol).Text)
        If Len(strPath) > 0 Then
            If Not Right(strPath, 1) = "\" Then strPath = strPath & "\"
            Grid1.Cell(lngRow, lngCol).Text = strPath
        End If
    End If
End Sub

Private Sub Grid1_Click()
    With Grid1.ActiveCell
        If .Row = 12 And .Col = 7 Then  '���뵥Ԫ�����TextBox�ؼ�������Ǻ�*
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
    '��������ֵ
    
    URL = ""
    Changed = True
    If Row <> (Grid1.Rows - 1) Then Exit Sub
    
    If Col = 3 Then '����
        If MsgBox("ȷ���������в���ֵ��", vbQuestion + vbOKCancel, "����ѯ��") = vbOK Then Call msSaveParameter(True)
    ElseIf Col = 7 Then '�˳�
        Unload Me
    End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Dim intRow As Integer, intCol As Integer

    intRow = Grid1.ActiveCell.Row
    intCol = Grid1.ActiveCell.Col
    If intRow = 19 And intCol = 3 Then  '�������룺����·��
        KeyCode = 0
    End If
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer, intCol As Integer

    intRow = Grid1.ActiveCell.Row
    intCol = Grid1.ActiveCell.Col
    If intRow = 19 And intCol = 3 Then  '�������룺����·��
        KeyAscii = 0
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
