VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmSysUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   Icon            =   "frmSysUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5655
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin FrameFileUpdate.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   2640
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1440
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2640
      Top             =   2760
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1995
   End
End
Attribute VB_Name = "frmSysUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnHide As Boolean     '���´��������ش�ģʽ����ʾ��ģʽ
Dim mblnCheckStart As Boolean   '�ѿ�ʼ����ʶ
Dim mblnUpdateFinish As Boolean     '������ɱ�ʶ
Dim mblnUnload As Boolean '�˳������ʶ
Dim mblnFTErr As Boolean '�����쳣�˳���ʶ


Private Function mfCheckUpdate() As Boolean
    '������
    Dim strFileLoc As String, strFileNet As String, strVerLoc As String, strVerNet As String
    
    strFileLoc = gVar.AppPath & gVar.EXENameOfClient
    If Not gfDirFile(strFileLoc) Then Exit Function
    strVerLoc = Trim(gfBackVersion(strFileLoc))
    If Len(strVerLoc) = 0 Then Exit Function
    
    If Me.Winsock1.Item(1).State <> 7 Then Exit Function
    Call msSetText("����������֤�汾�С���", vbBlue)
    Call gfSendInfo(gVar.PTVersionOfClient & strVerLoc, Me.Winsock1.Item(1))
    
End Function

Private Function mfConnect(Optional ByVal blnCon As Boolean = True) As Boolean
    '���������������
    Static lngCount As Long
            
    lngCount = lngCount + 1
    If lngCount >= 2 Then
        Call msSetText("�汾���ʧ�ܣ��޷����ӷ�������" & vbCrLf & _
                       "��ȷ�Ϸ�����IP��ַ�Ƿ���ȷ����������������ų������������и��³���", vbRed)
        If mblnHide Then
            Call gsAlarmAndLogEx("���³����޷���������������ӣ���ȷ��IP��ַ�Ƿ���ȷ���������������", "���¼��ʧ��")
            mblnUnload = True 'Unload Me  '��½�ͻ��˳��򼤻�ĸ��³�����ж��
        End If
        Exit Function    '����[lngCount]�κ���������
    End If
    
    With Me.Winsock1.Item(1)
        If Label1(1).Caption = gVar.ClientStateDisConnected Then
            If .State <> 0 Then .Close  '�ȹر�
            .RemoteHost = gVar.TCPSetIP
            .RemotePort = gVar.TCPSetPort
            .Connect
            If .State = 7 Then gVar.TCPStateConnected = True
        End If
    End With
End Function

Private Function mfShellSetup(ByVal strFile As String) As Boolean
    '�رտͻ��˳���ִ�и��°�װ��
    
    Dim strClient As String
    
    If MsgBox("�Ƿ�����ִ�и��³���", vbQuestion + vbYesNo, "��װѯ��") = vbYes Then
        If gfCloseApp(gVar.EXENameOfClient) Then   '�رտͻ���exe
            If gfShellExecute(strFile) Then     '���а�װ��
                Unload Me
                Exit Function
            End If
        Else
            MsgBox "��ȷ���ѹرտͻ��˳��򣬲��������и��³���", vbInformation, "����"
        End If
    Else
        Rem Call Winsock1_Close(1)
        Rem Unload Me   '��û�ҵ����ʷ��������쳣���˳���������mblnUnload��ʶ��Timer�ؼ������˳���
        MsgBox "��ȡ�����θ����ˡ�", vbInformation, "����ȡ��"
        mblnUnload = True '�˳������ʶ������Unload Me���
    End If
End Function

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    '��ע����м��ز���ֵ�����ñ�����
    
    If Not blnLoad Then Exit Sub
    
    On Error Resume Next    '��/���ܺ������̿������쳣
    With gVar
        .TCPDefaultIP = Me.Winsock1.Item(0).LocalIP '����IP��ַ
        .TCPSetIP = gfCheckIP(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPDefaultIP)) 'Ҫ���ӷ����IP��ַ
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , .TCPDefaultPort, 10000, 65535) 'Ҫ���ӵķ������˿�
        
        .UserComputerName = gfBackComputerInfo(ciComputerName)
        .UserLoginName = gfBackComputerInfo(ciUserName)
        .UserFullName = "UpdateProgram"
    End With
End Sub

Private Sub msSetLabel(ByVal strCaption As String, ByVal BackColor As Long)
    Me.Label1.Item(1).Caption = strCaption
    Me.Label1.Item(1).BackColor = BackColor
End Sub

Private Sub msSetText(ByVal strTxt As String, ByVal ForeColor As Long)
    Me.Text1.Text = strTxt
    Me.Text1.ForeColor = ForeColor
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim strCmd As String, arrCmd() As String
    
    Label1.Item(0).Caption = ""
    Text1.BackColor = Me.BackColor
    Timer1.Interval = 1000
    Timer1.Enabled = True

    ReDim gArr(1)
    
    Call Main
    Call msLoadParameter(True)
    
    '����Ƿ��������в���������û�����˳�����
    strCmd = Command()
    If Len(strCmd) = 0 Then
        GoTo LineUnload '��ֱֹ���������³��򣬱�����������
    Else
        arrCmd = Split(strCmd, gVar.CmdLineSeparator)
        
        If UCase(arrCmd(0)) <> UCase(gVar.EXENameOfClient) Then
            GoTo LineUnload    '��������е�һ���ַ��̶�Ϊexe�ļ�������������Ϊ�Ƿ��������³��򣬲�׼ִ��
        End If
        
        If UBound(arrCmd) > 0 Then  '�ж�����������Ƿ�������ش�������
            If LCase(arrCmd(1)) = LCase(gVar.CmdLineParaOfHide) Then
                mblnHide = True
                Me.Hide
            End If
        End If
    End If
    
    Call msSetLabel(gVar.ClientStateDisConnected, vbRed)
    Call gsLoadSkin(Me, Me.SkinFramework1, sMSVst, True)
    Call mfConnect(True)
    
    Exit Sub
    
LineUnload:
    Unload Me   '�������³�End Sub�����ٸ��κ���Ч����
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж�ش���
    
    On Error Resume Next
    
    mblnHide = False
    mblnCheckStart = False
    mblnUpdateFinish = False
    
    Me.Winsock1.Item(1).Close
    gArr(1) = gArr(0)
    
    Close
    
End Sub

Private Sub Timer1_Timer()
    Const conConn As Byte = 1       '����״̬�����conConn��
    Const conState As Byte = 5      '���ӷ����������conState��
    
    Static byteConn As Byte
    Static byteState As Byte
    Static byteDotCount As Byte
    
    If mblnUnload Then '�˳�����
        Unload Me
        If mblnFTErr Then MsgBox "�������������̷����쳣������ȷ���˳�����", vbExclamation, "�����쳣"
        Exit Sub
    End If
    
    byteConn = byteConn + 1
    byteState = byteState + 1
    
    If byteConn >= conConn Then
        If Me.Winsock1.Item(1).State = 7 Then
            Call msSetLabel(gVar.ClientStateConnected, vbGreen)
            gVar.TCPStateConnected = True
            If Not mblnCheckStart And gArr(1).Connected Then
                mblnCheckStart = True
                Call mfCheckUpdate
            End If
        Else
            Call msSetLabel(gVar.ClientStateDisConnected, vbRed)
            gVar.TCPStateConnected = False
        End If
        byteConn = 0    '��λ��̬����
    End If
    
    If byteState >= conState Then
        If Me.Winsock1.Item(1).State <> 7 Then
            If Not mblnUpdateFinish Then Call mfConnect
        End If
        byteState = 0   '��λ��̬����
    End If
    
    If gArr(1).FileTransmitState Then
        byteDotCount = byteDotCount + 1
        If byteDotCount > 6 Then byteDotCount = 1
        Me.Label1.Item(0).Caption = "����������" & String(byteDotCount, "��")
    End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '���䱻�ر�
    If UBound(gArr) = 1 Then
        gArr(1) = gArr(0)
        Rem Debug.Print "Winsock1_Close trigger all time ?"
    End If
    
    If mblnCheckStart Then
        Call msSetText("�����������жϣ��汾���¼��ʧ�ܣ�", vbRed)
        mblnCheckStart = False
    End If
    Label1.Item(0).Caption = ""
    
    If mblnHide Then mblnUnload = True 'Unload Me  '�쳣ʱж��
End Sub


Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    '���շ������˴�����Ϣ���ļ�
    
    Dim strGet As String    '�����ַ���Ϣ
    Dim byteGet() As Byte   '�����ļ�
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '�ַ���Ϣ����״̬��
            
            Me.Winsock1.Item(Index).GetData strGet
            
            If Not gfRestoreInfo(strGet, Me.Winsock1.Item(Index)) Then '
                
            End If
            
            If InStr(strGet, gVar.PTClientConfirm) Then '�յ�Ҫ�ظ������ȷ�����ӵ���Ϣ
                Call gfSendInfo(gVar.PTClientIsTrue, Me.Winsock1.Item(Index))
                Call gfSendClientInfo(gVar.UpdatePCName, gVar.UpdateAccount, gVar.UpdateUserName, Me.Winsock1.Item(Index))
                .Connected = True
                
            ElseIf InStr(strGet, gVar.PTConnectIsFull) > 0 Then '����˷�������������
                Me.Timer1.Enabled = False
                If Not mblnHide Then
                    MsgBox "�ͻ������������������ޣ��������û��˳������ԣ�", vbCritical, "��������������"
                End If
                mblnUnload = True ' Call Unload(Me)
                
            ElseIf InStr(strGet, gVar.PTConnectTimeOut) > 0 Then '����˷�������ʱ�䵽
                Me.Timer1.Enabled = False
                If Not mblnHide Then
                    MsgBox "���������������ʱ���ѵ���", vbExclamation, "����ʱ��������ʾ"
                End If
                mblnUnload = True 'Call Unload(Me)
                
            ElseIf InStr(strGet, gVar.PTVersionNeedUpdate) > 0 Then '��Ҫ����
                Dim strVer As String
                
                strVer = Mid(strGet, Len(gVar.PTVersionNeedUpdate) + 1)
                Call msSetText("�����°棺" & strVer, vbBlue)
                If Not gfCloseApp(gVar.EXENameOfClient) Then '�رտͻ���
                    Me.Winsock1.Item(Index).Close
                    MsgBox "�޷��رտͻ��˳��򣬵��¸����쳣�����˳����£�", vbCritical, "�ر��쳣����"
                End If
                
            ElseIf InStr(strGet, gVar.PTVersionNotUpdate) > 0 Then '����Ҫ����
                Dim strNot As String
                
                If Len(strGet) = Len(gVar.PTVersionNotUpdate) Then
                    strNot = "����ǰ�İ汾�������°汾������Ҫ���¡�"
                    Call msSetText(strNot, vbBlue)
                    If mblnHide Then mblnUnload = True 'Unload Me  '����ģʽ�򿪸��´���ʱ���޸�����ֱ���˳�
                Else
                    strNot = Mid(strGet, Len(gVar.PTVersionNotUpdate) + 1)
                    strNot = "�汾����쳣��" & strNot
                    Call msSetText(strNot, vbMagenta)
                End If
                
                mblnUpdateFinish = True
                
            End If
            
            Debug.Print "Get Server Info:" & strGet, bytesTotal
            '�ַ���Ϣ����״̬��
            
        Else
            '�ļ�����״̬��
            
            If .FileNumber = 0 Then
                .FileNumber = FreeFile
                
                If gfCloseApp(.FileName) Then
                    Open .FilePath For Binary As #.FileNumber
                End If
                Rem MsgBox .FileNumber & "�ļ���Ϣ��" & .FilePath
                
                LabelProgressBar1.Min = 0
                LabelProgressBar1.Max = .FileSizeTotal
                LabelProgressBar1.Value = 0
            End If
            
            Rem On Error GoTo LineErr
            ReDim byteGet(bytesTotal - 1)
            Me.Winsock1.Item(Index).GetData byteGet, vbArray + vbByte
            Put #.FileNumber, , byteGet
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal
            LabelProgressBar1.Value = .FileSizeCompleted
            
            If .FileSizeCompleted >= .FileSizeTotal Then
                Dim strSetupFile As String
                
                strSetupFile = .FilePath
                Close #.FileNumber
                Call gfSendInfo(gVar.PTFileEnd, Winsock1.Item(Index))
                gArr(Index) = gArr(0)
                Label1.Item(0).Caption = "������ɣ�"
                
                Call mfShellSetup(strSetupFile)
                
                Debug.Print "Received Over"
            End If
            
            '�ļ�����״̬��
        End If
    End With
    
    Exit Sub
LineERR:
    mblnFTErr = True
    mblnUnload = True
End Sub


Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Index <> 0 Then
        If gArr(Index).FileTransmitState Then   '�쳣ʱ����ļ�������Ϣ
            Close #gArr(Index).FileNumber
            gArr(Index) = gArr(0)
        End If
        If mblnHide Then mblnUnload = True  '�쳣ʱж��
    End If
End Sub
