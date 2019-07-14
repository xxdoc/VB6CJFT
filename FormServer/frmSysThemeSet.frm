VERSION 5.00
Begin VB.Form frmSysThemeSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6585
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ָ�Ĭ������"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   3600
      Width           =   1500
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Index           =   2
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   280
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���塢��ɫѡ��"
      Height          =   180
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ѡ��"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����ļ�·����"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "frmSysThemeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    '�ָ�Ĭ������
    If MsgBox("�Ƿ񽫴�������ָ���Ĭ��ֵ��", vbQuestion + vbYesNo, "ȷ����ʾ") = vbNo Then
        Exit Sub
    Else
        Call gsLoadSkin(gWind, gWind.SkinFramework1, -1, False, , , True)
        List1.Item(1).ListIndex = -1
        List1.Item(2).ListIndex = -1
    End If
End Sub

Private Sub Command3_Click()
    '�˳�����
    Unload Me
End Sub

Private Sub Form_Load()
    '���ڼ���
    
    Dim skinDes As SkinDescription
    Dim skinDesAll As SkinDescriptions
    Dim strFPath As String, strFName As String
    Dim strRegRes As String, strRegIni As String
    Dim L As Long, M As Long
    
    Text1.Text = gVar.FolderNameBin
    
    Set skinDesAll = gWind.SkinFramework1.EnumerateSkinDirectory(gVar.FolderNameBin, False) 'ö�ٳ��ļ�����������Դ�ļ�
    If skinDesAll.Count > 0 Then
        List1.Item(1).Clear
        For Each skinDes In skinDesAll  '���������ļ����б���
            strFPath = skinDes.Path
            strFName = Right(strFPath, Len(strFPath) - InStrRev(strFPath, "\"))
            List1.Item(1).AddItem strFName
        Next
        
        strRegRes = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinSvrRes, "")
        strRegIni = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinSvrIni, "")
        strRegRes = Mid(strRegRes, InStrRev(strRegRes, "\") + 1)    'ȥ��·�������ļ���
        strRegIni = Mid(strRegIni, InStrRev(strRegIni, "\") + 1)
        If Len(strRegRes) > 0 Then
            For L = 0 To List1.Item(1).ListCount - 1    '��λ��ǰ��������
                If LCase(strRegRes) = LCase(List1.Item(1).List(L)) Then
                    List1.Item(1).ListIndex = L
                    If Len(strRegIni) > 0 Then
                        If List1.Item(2).ListCount > 0 Then
                            For M = 0 To List1.Item(2).ListCount - 1
                                If LCase(strRegIni) = LCase(List1.Item(2).List(M)) Then
                                    List1.Item(2).ListIndex = M
                                    Exit For    '�˳�ѭ��
                                End If
                            Next
                        End If
                    End If
                    Exit For    '�˳�ѭ��
                End If
            Next
            
        End If
    End If
    
End Sub

Private Sub List1_Click(Index As Integer)
    '������Դѡ��
    Dim skinDes As SkinDescription
    Dim skinIni As SkinIniFile
    Dim strRes As String, strIni As String
        
    If Index = 1 Then
        If List1.Item(1).ListIndex = -1 Then Exit Sub   '�б�Ϊ��ʱ�����Ч
        
        Set skinDes = gWind.SkinFramework1.EnumerateSkinFile(gVar.FolderNameBin & List1.Item(1).Text) 'ö�ٳ��������ļ������������ļ�
        If skinDes.Count > 0 Then
            List1.Item(2).Clear
            For Each skinIni In skinDes 'ö�ٳ���������Դ�ļ������ļ����ڶ����б���
                List1.Item(2).AddItem skinIni.IniFileName
            Next
            List1.Item(2).ListIndex = 0 '��Ĭ��ѡ�е�һ�����ļ�
        End If
        
    End If
    
    If Index = 2 Then
        If Me.Visible Then   '������ѡ
            strRes = gVar.FolderNameBin & List1.Item(1).List(List1.Item(1).ListIndex)   'ȫ·���ļ�����Ч
            strIni = List1.Item(2).List(List1.Item(2).ListIndex)    'ע����ļ���û��·����
            Call gsLoadSkin(gWind, gWind.SkinFramework1, -1, False, strRes, strIni, True)
        End If
    End If
    
    Set skinDes = Nothing
    Set skinIni = Nothing
End Sub
