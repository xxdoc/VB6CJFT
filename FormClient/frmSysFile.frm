VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmSysFile 
   Caption         =   "�ļ�����"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   9945
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ϴ�"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   6255
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7646
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�·����"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frmSysFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function mfDeleteFile(ByVal strFID As String) As Boolean
    'ɾ��һ���ļ�
    Dim strSQL As String
    Dim rsDel As ADODB.Recordset
    
    On Error GoTo LineERR
    
    strFID = Trim(strFID)
    If Len(strFID) = 0 Then Exit Function
    strSQL = "SELECT * FROM tb_FT_Lib_File WHERE FileID ='" & strFID & "' "
    Set rsDel = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsDel.State = adStateOpen Then
        If rsDel.RecordCount = 1 Then
            rsDel.Delete
        End If
    End If
    mfDeleteFile = True
    
LineERR:
    If Not rsDel Is Nothing Then If rsDel.State = adStateOpen Then rsDel.Close
    Set rsDel = Nothing
    If Err.Number > 0 Then Call gsAlarmAndLog("�ļ�ɾ���쳣")

End Function

Private Sub msLoadFileList(Optional ByVal blnLD As Boolean = True)
    '�����ļ���Ϣ�����
    Dim strSQL As String
    Dim rsFile As ADODB.Recordset
    Dim K As Long, C As Long
    
    strSQL = "SELECT FileID ,FileClassify ,FileExtension ,FileOldName ,FileSaveName ," & _
             "FileSize ,FileSaveLocation ,FileUploadMen ,FileUploadTime FROM tb_FT_Lib_File "
    Set rsFile = gfBackRecordset(strSQL)
    If rsFile.State = adStateOpen Then
        C = rsFile.RecordCount
        If C > 0 Then
            With Grid1
                .AutoRedraw = False
                .Rows = C + 1
                If C < 20 Then .Rows = 21
                While Not rsFile.EOF
                    K = K + 1
                    Grid1.Cell(K, 0).Text = K
                    Grid1.Cell(K, 1).Text = rsFile.Fields("FileID") & ""
                    Grid1.Cell(K, 2).Text = rsFile.Fields("FileSaveName") & ""
                    Grid1.Cell(K, 3).Text = rsFile.Fields("FileSaveLocation") & ""
                    Grid1.Cell(K, 4).Text = ""
                    Grid1.Cell(K, 5).Text = rsFile.Fields("FileClassify") & ""
                    Grid1.Cell(K, 6).Text = rsFile.Fields("FileExtension") & ""
                    Grid1.Cell(K, 7).Text = rsFile.Fields("FileSize") & ""
                    Grid1.Cell(K, 8).Text = rsFile.Fields("FileUploadMen") & ""
                    Grid1.Cell(K, 9).Text = rsFile.Fields("FileUploadTime") & ""
                    Grid1.Cell(K, 10).Text = "��"
                    Grid1.Cell(K, 11).Text = "ɾ��"
                    Grid1.Cell(K, 12).Text = rsFile.Fields("FileOldName") & ""
                    rsFile.MoveNext
                Wend
                .Range(1, 10, K, 11).ForeColor = vbBlue
                .ReadOnly = True
                .AutoRedraw = True
                .Refresh
            End With
        End If
        rsFile.Close
    End If
    Set rsFile = Nothing
End Sub

Private Sub Command1_Click()
    '���
    
    With CommonDialog1
        .DialogTitle = "ѡ��һ��Ҫ�ϴ����ļ�"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
        Text1.Text = .FileName
    End With
End Sub

Private Sub Command2_Click()
    '�ϴ�
    Const conLngSize As Long = 524288000 '500MB=500*1024*1024=524288000(B)
    Dim sckFile As MSWinsockLib.Winsock
    
    gVar.FTUploadFileNameNew = ""
    gVar.FTUploadFilePath = Trim(Text1.Text)
    If Len(gVar.FTUploadFilePath) = 0 Then
        MsgBox "����ѡ��һ���ļ���", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    If MsgBox("ȷ��Ҫ�ϴ���ѡ�ļ���", vbQuestion + vbOKCancel, "����") = vbCancel Then Exit Sub
    
    If Not gfFileExist(gVar.FTUploadFilePath) Then
        MsgBox "�ļ������ڣ���ȷ�ϻ�����ѡ��", vbCritical, "����"
        Exit Sub
    End If
    
    gVar.FTUploadFileSize = FileLen(gVar.FTUploadFilePath)   '��ȡ�ļ���С����λ�ֽ�
    If gVar.FTUploadFileSize > conLngSize Then
        MsgBox "��ѡ�ļ���С���ܳ���500MB��", vbCritical, "����"
        Exit Sub
    End If
    
    gVar.FTUploadFileNameOld = Mid(gVar.FTUploadFilePath, InStrRev(gVar.FTUploadFilePath, "\") + 1)   '��ȡ����·�����ļ���
    gVar.FTUploadFileExtension = Mid(gVar.FTUploadFilePath, InStrRev(gVar.FTUploadFilePath, ".") + 1)  '��ȡ�ļ�����չ��
    gVar.FTUploadFileNameNew = gfBackFileName(udUpperCase, 30) '�����ļ��ڷ���˱����õ��ļ�����30�������д��ĸ
    gVar.FTUploadFileFolder = gVar.FolderStore   '�����ļ��ڷ���˵Ĵ洢λ�á�ע�ⲻ��·��
    gVar.FTUploadFileClassify = "�����ļ�"
    
    Set sckFile = gWind.Winsock1.Item(1)
    Call gsLoadFileInfo(sckFile.Index, True)      '�����ļ�������Ϣ
    If sckFile.State = 7 Then    '���������������״̬
        If gfSendInfo(gfFileInfoJoin(sckFile.Index, ftSend), sckFile) Then  '�������������ļ���Ϣ
            Debug.Print "Client: �ϴ��ļ�[" & gVar.FTUploadFileNameNew & "]����Ϣ����OK," & Now
            Timer1.Enabled = True
        End If
    Else
        MsgBox "������������ӱ��Ͽ����޷��ϴ���", vbCritical, "����"
    End If
    Set sckFile = Nothing
End Sub

Private Sub Form_Load()
    '���ڼ��أ��������
    
    Timer1.Interval = 100   '100ms
    Timer1.Enabled = False
    Text1.Text = ""
    Text1.Locked = True
    Text1.Font.Size = 11
    With Grid1
        .AutoRedraw = False
        .Rows = 16
        .Cols = 13
        .Cell(0, 0).Text = "���"
        .Cell(0, 1).Text = "�ļ�ID"
        .Cell(0, 2).Text = "�洢����"
        .Cell(0, 3).Text = "�洢λ��"
        .Cell(0, 4).Text = "����λ��"
        .Cell(0, 5).Text = "�ļ�����"
        .Cell(0, 6).Text = "��չ��"
        .Cell(0, 7).Text = "�ļ���С"
        .Cell(0, 8).Text = "�ϴ���"
        .Cell(0, 9).Text = "�ϴ�����"
        .Cell(0, 10).Text = "�鿴"
        .Cell(0, 11).Text = "ɾ��"
        .Cell(0, 12).Text = "�ļ���"
        .Cell(0, 7).Comment = "��λ���ֽ�(B)"
        .Cell(0, 11).Comment = "�ļ�ɾ��֮�󲻿ɻָ�"
        .Column(0).Width = 40
        .Column(1).Width = 0
        .Column(2).Width = 0
        .Column(3).Width = 0
        .Column(4).Width = 0
        .Column(5).Width = 100
        .Column(6).Width = 50
        .Column(7).Width = 70
        .Column(8).Width = 70
        .Column(9).Width = 120
        .Column(10).Width = 50
        .Column(11).Width = 50
        .Column(12).Width = 150
        .ExtendLastCol = True
        .rowHeight(0) = 30
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(9).Alignment = cellCenterCenter
        .Column(10).Alignment = cellCenterCenter
        .Column(11).Alignment = cellCenterCenter
        .Column(10).CellType = cellHyperLink
        .Column(11).CellType = cellHyperLink
        .Column(9).FormatString = gVar.Formaty_M_dH_m_s
        .DisplayRowIndex = True
        .AllowUserSort = True
        .AllowUserResizing = True
        .Appearance = Flat
        .BackColorBkg = Me.BackColor
        .BackColorFixed = RGB(121, 151, 219)
        .BackColor2 = RGB(250, 235, 215)
        .BackColorFixedSel = vbYellow
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    Call msLoadFileList(True)
End Sub

Private Sub Form_Resize()
    '���ڴ�С�ĸı�
    Const conLngW As Long = 10650 + 720
    Const conLngH As Long = 5000 + 720
    Dim lngW As Long, lngH As Long
    
    On Error Resume Next
    
    lngW = Me.Width
    lngH = Me.Height
    If lngW > conLngW Then
        Grid1.Width = lngW - 900
    Else
        Grid1.Width = 10000
    End If
    If lngH > conLngH Then
        Grid1.Height = lngH - 1600
    Else
        Grid1.Height = 5000
    End If
End Sub

Private Sub Grid1_HyperLinkClick(ByVal Row As Long, ByVal Col As Long, URL As String, Changed As Boolean)
    Dim sckFile As MSWinsockLib.Winsock
    
    URL = ""
    Changed = True
    gVar.FTDownloadFilePath = ""
    If Row = 0 Then Exit Sub
    If Col = 11 Then    'ɾ��
        If Len(Trim(Grid1.Cell(Row, 1).Text)) > 0 Then
            If MsgBox("ȷ��Ҫɾ����ѡ�ļ���" & Grid1.Cell(Row, 12).Text & "����", vbQuestion + vbOKCancel, "ѯ��") = vbOK Then
                If Trim(InputBox("������ɾ���ļ�����ʾ���֣�123", "�ļ�ɾ����֤")) = "123" Then
                    Call mfDeleteFile(Grid1.Cell(Row, 1).Text)
                    Call msLoadFileList(True)
                    MsgBox "�ļ�ɾ���ɹ���", vbInformation, "��ʾ"
                End If
            End If
        End If
    ElseIf Col = 10 Then    '�鿴
        Rem Debug.Print Grid1.Cell(Row, 12).Text, Grid1.Cell(Row, 4).Text
        gVar.FTDownloadFilePath = Trim(Grid1.Cell(Row, 4).Text)
        If gfFileExist(gVar.FTDownloadFilePath) Then    '�ļ�����
            If FileLen(gVar.FTDownloadFilePath) = Grid1.Cell(Row, 7).Text Then '��С���
                Call gfFileOpen(gVar.FTDownloadFilePath)    'ֱ�Ӵ��ļ����ļ�����������
                Exit Sub    '�˳�����
            End If
        End If
        
        '�����ļ�
        gVar.FTDownloadFileClassify = Trim(Grid1.Cell(Row, 5).Text)
        gVar.FTDownloadFileExtension = Trim(Grid1.Cell(Row, 6).Text)
        gVar.FTDownloadFileFolder = Trim(Grid1.Cell(Row, 3).Text)
        gVar.FTDownloadFileNameNew = Trim(Grid1.Cell(Row, 2).Text)
        gVar.FTDownloadFileNameOld = Trim(Grid1.Cell(Row, 12).Text)
        gVar.FTDownloadFileSize = Trim(Grid1.Cell(Row, 7).Text)
        gVar.FTDownloadFilePath = gVar.AppPath & gVar.FTDownloadFileFolder & "\" & gVar.FTDownloadFileNameNew
        Set sckFile = gWind.Winsock1.Item(1)
        Call gsLoadFileInfo(sckFile.Index, False) '�����ļ���Ϣ
        If sckFile.State = 7 Then
            If gfSendInfo(gfFileInfoJoin(sckFile.Index, ftReceive), sckFile) Then
                Debug.Print "Client��Ҫ���ص��ļ�[" & Grid1.Cell(Row, 2).Text & "]����Ϣ�ѷ���" & Now
                Timer1.Enabled = True
            End If
        Else
            MsgBox "������������ӱ��Ͽ����޷������ļ���", vbCritical, "����"
        End If
    End If
    Set sckFile = Nothing
End Sub

Private Sub Timer1_Timer()
    '�ж��ϴ��������Ƿ����
    Dim strNewName As String
    
    If Not gArr(1).FileTransmitNotOver Then '�������
        If Not Me.Enabled Then Exit Sub     '����δ���������
        If Not gArr(1).FileTransmitError Then   '����������ɣ�û���쳣
            If gVar.FTUploadOrDownload Then     '�ϴ���������
                If Len(gfSaveFile(Me)) > 0 Then '�ɹ������ļ���Ϣ�����ݿ�
                    Call msLoadFileList(True)   'ˢ�±��
                End If
            Else    '���ؽ�������
                If gfFileExist(gVar.FTDownloadFilePath) Then
                    strNewName = Left(gVar.FTDownloadFilePath, InStrRev(gVar.FTDownloadFilePath, "\")) & gVar.FTDownloadFileNameOld
                    If gfFileReNameEx(gVar.FTDownloadFilePath, strNewName) Then '��ԭ���ϴ�ʱ���ļ���
                        gVar.FTDownloadFilePath = strNewName
                        Call gfFileOpen(gVar.FTDownloadFilePath)    '���ļ�
                        Grid1.Cell(Grid1.ActiveCell.Row, 4).Text = gVar.FTDownloadFilePath
                    End If
                End If
            End If
        End If
        Timer1.Enabled = False  'ֹͣ�ж�
    End If
    
End Sub
