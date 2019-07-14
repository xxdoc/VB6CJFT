VERSION 5.00
Begin VB.Form frmRestartServer 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7800
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "�����ڽ���60���ر�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   6165
   End
End
Attribute VB_Name = "frmRestartServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrAppPath As String, mstrFileExe As String, mstrFileExePath As String
Dim mstrFileNameErrLog As String, mstrFileNameRunLog As String
Private Const mconStrFormaty_M_dH_m_s As String = "yyyy-MM-dd HH:mm:ss"

Private Type gtypeValueAndErr    '���ڷ��ز���ֵ�Ĺ��̣�˳�㷵���쳣����
    Result As Boolean
    ErrNum As Long
End Type

Private Enum genumFileOpenType   '���ļ���ʽ
    udAppend    '��˳���ͷ��ʣ����ַ�׷�ӵ��ļ�
    udBinary    '�Զ����Ʒ���
    udInput     '��˳���ͷ��ʣ����ļ������ַ�
    udOutput    '��˳���ͷ��ʣ����ļ�����ַ�
    udRandom    '�������ʽ
End Enum

Private Enum genumFileWriteType  'д���ļ���ʽ
    udPut       '��Get����.For Binary��Random.
    udWrite     '��Input����
    udPrint     '��Line Input �� Input����
End Enum


Private Sub AlarmAndLog(Optional ByVal strErr As String, Optional ByVal blnMsgBox As Boolean = True, _
        Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    'ϵͳ�쳣��ʾ��д���쳣��־
    Dim strMsg As String
    
    strMsg = "�쳣���ţ�" & Err.Number & vbCrLf & "�쳣������" & Err.Description
    If blnMsgBox Then MsgBox strMsg, MsgButton, strErr
    Call FileWrite(mstrFileNameErrLog, strErr & vbTab & Replace(strMsg, vbCrLf, vbTab))
    
End Sub

Private Function CloseExeFile(ByVal strName As String) As Boolean
    '�ر�ָ��exe�������
    
    Dim winHwnd As Long
    Dim retVal As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    On Error GoTo LineErr
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")    'VB�Դ��ĺ���
    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='" & strName & "' ")
    For Each objProcess In colProcessList
        retVal = objProcess.Terminate
        If retVal <> 0 Then Exit Function   '���۲�=0ʱ�رս��̳ɹ������ɹ�ʱ����ֵ��Ϊ��
    Next
    CloseExeFile = True   'ȫ���رճɹ��򲻴��ڸý�����ʱ
    
LineErr:
    Set objWMIService = Nothing
    Set colProcessList = Nothing
    Set objProcess = Nothing
    If Err.Number > 0 Then
        Call AlarmAndLog("�رս����쳣")
    End If
End Function

Private Sub CloseServer(ByVal strFileExe As String, ByVal strFileExePath As String)
    '�رղ���������
    
    If CloseExeFile(strFileExe) Then
        Call FileWrite(mstrFileNameRunLog, "�ر�" & strFileExe & "����ɹ�")
        If ShellExePath(strFileExePath) = 0 Then
            Call FileWrite(mstrFileNameRunLog, "����" & strFileExe & "����ʧ��")
        Else
            Call FileWrite(mstrFileNameRunLog, "����" & strFileExe & "����ɹ�")
        End If
    Else
        Call FileWrite(mstrFileNameRunLog, "�ر�" & strFileExe & "����ʧ��")
    End If
End Sub

Private Function FileExistEx(ByVal strPath As String) As gtypeValueAndErr
    '��һ�ַ���ֵ��ʽ�����ж��ļ����ļ�Ŀ¼ �Ƿ����
    'ר������Ĺ���FileRepair����
    Dim strBack As String
    
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '���ַ�������
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then
            FileExistEx.Result = True
        Else
            FileExistEx.ErrNum = -1   '�����ڣ�Ҳû�쳣
        End If
    End If

LineErr:
    If Err.Number > 0 Then
        FileExistEx.ErrNum = Err.Number   '�쳣�ˣ�Ҳ������������
        Call AlarmAndLog("�ļ��жϷ����쳣")
    End If
End Function

Private Function FileRepair(ByVal strFile As String, Optional ByVal blnFolder As Boolean) As Boolean
    '��� �ļ�/�ļ��� ������ �򴴽�
    'ǰ����·�����ϲ�Ŀ¼�ɷ���
    '����blnFolderָ�������·��strFile���ļ�����ΪTrue��Ĭ�����ļ�False
    
    Dim strTemp As String
    Dim typBack As gtypeValueAndErr
    Dim lngLoc As Long
    
    If Right(strFile, 1) = "\" Then
        strFile = Left(strFile, Len(strFile) - 1)   'ȥ����ĩ��"\"
    End If
    strTemp = strFile
    If Len(strTemp) = 0 Then Exit Function          '��ֹ������ַ���
    
    On Error GoTo LineErr

    typBack = FileExistEx(strTemp)    '�ж��Ƿ����
    If Not typBack.Result Then          '�ļ�������
        If typBack.ErrNum = -1 Then     '�����쳣
            
            lngLoc = InStrRev(strTemp, "\") '�ж��Ƿ����ϲ�Ŀ¼
            If lngLoc > 0 Then              '���ϲ�Ŀ¼��ݹ�
                strTemp = Left(strTemp, lngLoc - 1) '�ó��ϲ�Ŀ¼�ľ���·��
                Call FileRepair(strTemp, True)    '�ݹ���������Ա�֤�ϲ�Ŀ¼����
            End If

            If blnFolder Then                   '����������ļ���
                MkDir strFile                   '�򴴽��ļ���
            Else                                '����������ļ�
                Close                           '�򴴽��ļ�
                Open strFile For Random As #1
                Close
            End If
            FileRepair = True '�����ɹ�����True
        End If
    Else
        FileRepair = True '·������ֱ�ӷ���True
    End If

LineErr:
    Close
End Function

Private Sub FileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As genumFileOpenType = udAppend, _
    Optional ByVal WriteMode As genumFileWriteType = udPrint)
    '��ָ��������ָ���ķ�ʽд��ָ���ļ���
    
    Dim intNum As Integer
    Dim strTime As String
    
    If Not FileRepair(strFile) Then Exit Sub
    intNum = FreeFile
    
    On Error Resume Next
    
    Select Case OpenMode
        Case udBinary
            Open strFile For Binary As #intNum
        Case udInput
            Open strFile For Input As #intNum
        Case udOutput
            Open strFile For Output As #intNum
        Case Else   '����Ե���udAppend
            Open strFile For Append As #intNum
    End Select
    
    strTime = Format(Now, mconStrFormaty_M_dH_m_s)
    Select Case WriteMode
        Case udWrite
            Write #intNum, strTime, strContent
        Case udPut
            Put #intNum, , strTime & vbTab & strContent
        Case Else   '����Ե���udPrint
            Print #intNum, strTime, strContent
    End Select
    Close #intNum
    
End Sub

Private Function ShellExePath(ByVal strExePath As String) As Long
    'ִ��EXE�ļ�
    
    On Error Resume Next
    
    ShellExePath = Shell(strExePath)
    
End Function


Private Sub Command1_Click()
    Unload Me '�˳�
End Sub

Private Sub Form_Load()
    '���ڼ���
    
    Dim strCmd As String
    Dim arrCmd() As String
        
    Me.Hide  '���ز���ʾ�ô���
        
    Me.Timer1.Interval = 1000
    Me.Timer1.Enabled = True
        
    
    'ģ�������ֵ
    mstrFileExe = "FFS.exe"
    mstrAppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    mstrFileExePath = mstrAppPath & mstrFileExe
    mstrFileNameErrLog = mstrAppPath & "Data\ErrorRecord.LOG"
    mstrFileNameRunLog = mstrAppPath & "Data\RunRecord.LOG"
    
    On Error Resume Next
    
    strCmd = Trim(Command())  '��ȡ�����в���ֵ
    If Len(strCmd) = 0 Then
        GoTo LineUnload
    Else
        arrCmd = Split(strCmd, " / ")
        If UCase(arrCmd(0)) <> UCase(mstrFileExe) Then
            GoTo LineUnload '��������е�һ���ַ��̶�Ϊexe�ļ�������������Ϊ�Ƿ������ó��򣬲�׼ִ��
        Else
            Me.Text1.Text = mstrFileExePath
        End If
        
        If UBound(arrCmd) > 0 Then  '�ж���������еڶ��������Ƿ�Ϊ�ر�ָ��
            If LCase(arrCmd(1)) = "close" Then
                Call CloseServer(mstrFileExe, mstrFileExePath)
            End If
        End If
    End If

LineUnload:
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Const cMax As Long = 60 'cMax����˳�
    Static lngCount As Long
    
    If lngCount > cMax Then  '�ƴ�����
        Call CloseServer(mstrFileExe, mstrFileExePath)
        Call FileWrite(mstrFileNameRunLog, "��ʱ�˳�" & App.EXEName & "����")
        Unload Me
    Else
        Label1.Caption = "�����ڽ���" & CStr(cMax - lngCount) & "���ر�" & String((lngCount Mod 4), "��")
        lngCount = lngCount + 1
    End If
End Sub
