VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmSysUser 
   Caption         =   "�û�����"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16185
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   16185
   Begin VB.HScrollBar Hsb 
      Height          =   255
      Left            =   14520
      TabIndex        =   23
      Top             =   6720
      Width           =   1455
   End
   Begin VB.VScrollBar Vsb 
      Height          =   1935
      Left            =   15600
      TabIndex        =   22
      Top             =   4680
      Width           =   255
   End
   Begin VB.Frame ctlMove 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   15615
      Begin VB.Frame Frame1 
         Caption         =   "�û���ɫָ��"
         ForeColor       =   &H00FF0000&
         Height          =   6375
         Index           =   1
         Left            =   7680
         TabIndex        =   21
         Top             =   0
         Width           =   7815
         Begin VB.Frame Frame1 
            Caption         =   "�û�ͷ��"
            ForeColor       =   &H00FF00FF&
            Height          =   2895
            Index           =   5
            Left            =   4000
            TabIndex        =   32
            Top             =   2880
            Width           =   3720
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   120
               Top             =   2160
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton Command4 
               Caption         =   "ѡ����Ƭ"
               Height          =   495
               Left            =   720
               TabIndex        =   36
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "*��ʽΪjpg|png|bmp *�ļ�С��5MB"
               ForeColor       =   &H000000FF&
               Height          =   420
               Index           =   32
               Left            =   1845
               TabIndex        =   37
               Top             =   2200
               Width           =   1710
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1815
               Left            =   600
               Top             =   240
               Width           =   1305
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2415
            Index           =   4
            Left            =   4000
            TabIndex        =   31
            Top             =   150
            Width           =   3720
            Begin VB.TextBox Text1 
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Index           =   5
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   34
               Text            =   "Text1"
               Top             =   360
               Width           =   2500
            End
            Begin VB.CommandButton Command3 
               Caption         =   "�û���ɫָ���������"
               Height          =   495
               Left            =   720
               TabIndex        =   33
               Top             =   1560
               Width           =   2415
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "��ѡ�û�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   31
               Left            =   120
               TabIndex        =   35
               Top             =   390
               Width           =   900
            End
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   4095
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7223
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "�û�"
         ForeColor       =   &H00FF0000&
         Height          =   6375
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   7455
         Begin VB.Timer Timer1 
            Left            =   4800
            Top             =   1920
         End
         Begin VB.Frame Frame1 
            Height          =   405
            Index           =   3
            Left            =   675
            TabIndex        =   27
            Top             =   4080
            Width           =   2500
            Begin VB.OptionButton Option1 
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   210
               Index           =   2
               Left            =   120
               TabIndex        =   29
               Top             =   150
               Width           =   855
            End
            Begin VB.OptionButton Option1 
               Caption         =   "ͣ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   3
               Left            =   1200
               TabIndex        =   28
               Top             =   150
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            Height          =   405
            Index           =   2
            Left            =   720
            TabIndex        =   24
            Top             =   2100
            Width           =   2500
            Begin VB.OptionButton Option1 
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   1
               Left            =   1200
               TabIndex        =   26
               Top             =   150
               Width           =   855
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Ů"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   150
               Width           =   855
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "�޸��û���Ϣ"
            Height          =   495
            Left            =   1800
            TabIndex        =   7
            Top             =   4800
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "����û�"
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   4800
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   720
            TabIndex        =   1
            Text            =   "Text2"
            Top             =   720
            Width           =   2500
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Text            =   "Combo2"
            Top             =   3240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000003&
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   240
            Width           =   2500
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2640
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   720
            PasswordChar    =   "*"
            TabIndex        =   2
            Text            =   "Text2"
            Top             =   1200
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   720
            TabIndex        =   3
            Text            =   "Text2"
            Top             =   1680
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   4
            Left            =   720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   3120
            Width           =   2500
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4095
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7223
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   1
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "״̬"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   4200
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   165
            TabIndex        =   20
            Top             =   1260
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�˺�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   165
            TabIndex        =   19
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��ʶ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   165
            TabIndex        =   18
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   165
            TabIndex        =   17
            Top             =   1740
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�Ա�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   165
            TabIndex        =   16
            Top             =   2220
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   165
            TabIndex        =   15
            Top             =   2700
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��ע"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   165
            TabIndex        =   14
            Top             =   3180
            Width           =   450
         End
         Begin VB.Label Label1 
            Caption         =   "*** ����ֻ�ܰ������ֻ��С��ĸ���ҳ�����20λ����"
            ForeColor       =   &H000000FF&
            Height          =   420
            Index           =   21
            Left            =   240
            TabIndex        =   13
            Top             =   5640
            Width           =   3060
         End
      End
   End
End
Attribute VB_Name = "frmSysUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngID As Long  'ѭ������
Private Const mKeyDept As String = "k"
Private Const mKeyUser As String = "u"
Private Const mOtherKey As String = "kOther"
Private Const mOtherText As String = "������Ա"
Private Const mKeyRole As String = "r"
Private Const mOtherKeyRole As String = "kOtherRole"
Private Const mOtherTextRole As String = "������ɫ"
Private Const mTwoBar As String = "--"


Private Function mfCheckPhoto(ByVal strPhoPath As String) As Boolean
    '���ͼƬ�Ƿ�����
    
    gVar.FTUploadFilePath = Trim(strPhoPath)    '��ȡ��Ƭ·��
    If Len(gVar.FTUploadFilePath) = 0 Then Exit Function
    
    If Not gfFileExist(gVar.FTUploadFilePath) Then    'ԭ�ļ�������
        MsgBox "ͷ����Ƭ��ԭ�ļ��Ѷ�ʧ�������б��棡", vbCritical, "��Ƭ����"
        Exit Function
    End If
    
    gVar.FTUploadFileNameOld = Mid(gVar.FTUploadFilePath, InStrRev(gVar.FTUploadFilePath, "\") + 1)    '��ȡ���ļ���
    If Len(gVar.FTUploadFileNameOld) > 50 Then
        MsgBox "��Ƭ�ļ������ܳ���50���ַ� ������������", vbExclamation, "���ļ�������"
        Exit Function
    End If
    
    gVar.FTUploadFileExtension = Mid(gVar.FTUploadFileNameOld, InStrRev(gVar.FTUploadFileNameOld, ".") + 1)      '��ȡ���ļ�����չ��
    If Len(gVar.FTUploadFileExtension) > 20 Then
        MsgBox "��Ƭ�ļ�����չ�����ܳ���20���ַ� ������������", vbExclamation, "��չ������"
        Exit Function
    End If
    
    gVar.FTUploadFileSize = FileLen(gVar.FTUploadFilePath)     '��ȡ�ļ���С
    If gVar.FTUploadFileSize > 5242880 Then
        MsgBox "��Ƭ��С���ܳ���5MB��", vbExclamation, "�ļ���С����"
        Exit Function
    End If
    
    gVar.FTUploadFileNameNew = gfBackFileName(udUpperCase, 30)   '��ȡ���ļ�����30�������д��ĸ��
    gVar.FTUploadFileFolder = gVar.FolderStore  '����˴洢λ�á�ע�ⲻ��·�������ļ�������
    gVar.FTUploadFileClassify = "����ͷ��"      '�ļ��洢���
    
    mfCheckPhoto = True '������ֵ
End Function

Private Function mfSavePhoto(Optional ByVal blnPho As Boolean) As Boolean
    '�ϴ�����ͷ����Ƭ�ļ���������
    Dim sckPho As MSWinsockLib.Winsock
    
    Set sckPho = gWind.Winsock1.Item(1)
    If gfFileExist(gVar.FTUploadFilePath) Then
        Call gsLoadFileInfo(sckPho.Index, True) '�����ļ�����������Ϣ
        If gWind.Winsock1.Item(1).State = 7 Then    'ʹ��MDI�����ϵ�Winsock�ؼ������ļ���Ϣ
            If gfSendInfo(gfFileInfoJoin(gWind.Winsock1.Item(1).Index, ftSend), gWind.Winsock1.Item(1)) Then
                Debug.Print "Client���ѷ���[ͷ����Ƭ]���ļ���Ϣ�������," & Now
                Timer1.Enabled = True
            End If
        End If
    End If
    Set sckPho = Nothing
End Function

Private Sub msLoadDept(ByRef tvwDept As MSComctlLib.TreeView)
    '���ز�����TreeView�ؼ���
    'Ҫ��1�����ݿ��в�����Ϣ��Dept����DeptID(Not Null)��DeptName(Not Null)��ParentID(Null)�����ֶΡ�
    'Ҫ��2�����ű���ֻ���������ţ�����Ϊ��˾���ƣ���ParentIDΪNull���������ŵ�ParentID������ΪNull��
    
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim arrDept() As String 'ע���±�Ҫ��0��ʼ
    Dim I As Long, lngCount As Long, lngOneCompany As Long
    Dim blnLoop As Boolean
        
    strSQL = "SELECT t1.DeptID ,t1.DeptName ,t1.ParentID ,t2.DeptName AS [ParentName] " & _
             "FROM tb_FT_Sys_Department AS [t1] " & _
             "LEFT JOIN tb_FT_Sys_Department AS [t2] " & _
             "ON t1.ParentID = t2.DeptID " & _
             "ORDER BY t1.ParentID ,t1.DeptName"    'ע���ֶ�˳�򲻿ɱ�
    Set rsDept = gfBackRecordset(strSQL)
    If rsDept.State = adStateClosed Then Exit Sub
    If rsDept.RecordCount > 0 Then
        
        tvwDept.Nodes.Clear
        Combo1.Item(0).Clear
        Combo1.Item(1).Clear
        
        While Not rsDept.EOF
            If IsNull(rsDept.Fields(3).Value) Then
                lngOneCompany = lngOneCompany + 1
                tvwDept.Nodes.Add , , mKeyDept & rsDept.Fields(0).Value, rsDept.Fields(1).Value, "SysCompany"
                tvwDept.Nodes.Item(mKeyDept & rsDept.Fields(0).Value).Expanded = True
            Else
                ReDim Preserve arrDept(3, lngCount)
                For I = 0 To 3
                    arrDept(I, lngCount) = rsDept.Fields(I).Value
                Next
                lngCount = lngCount + 1
                blnLoop = True
            End If
            
            Combo1.Item(0).AddItem rsDept.Fields(1).Value
            Combo1.Item(1).AddItem rsDept.Fields(0).Value
            
            rsDept.MoveNext
        Wend
        
    End If
    rsDept.Close
    Set rsDept = Nothing
    
    If blnLoop Then Call msLoadDeptTree(tvwDept, arrDept)

End Sub

Private Sub msLoadDeptTree(ByRef tvwTree As MSComctlLib.TreeView, ByRef arrLoad() As String)
    '������msLoadDept�������ʹ�������ز����б�
    
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    Static C As Long
    
    With tvwTree
        For J = LBound(arrLoad, 2) To UBound(arrLoad, 2)
            For I = 1 To .Nodes.Count   'ע��˴��±��1��ʼ
                If .Nodes.Item(I).Key = mKeyDept & arrLoad(2, J) Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, mKeyDept & arrLoad(0, J), arrLoad(1, J), "threemen"
                    .Nodes.Item(mKeyDept & arrLoad(0, J)).Expanded = True
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(3, lngCount)
                For K = 0 To 3
                    arrOther(K, lngCount) = arrLoad(K, J)
                Next
                lngCount = lngCount + 1
            End If
            
        Next
    End With
    
    C = C + 1
    If C > 64 Then Exit Sub '��ֹ�ݹ����̫��¶�ջ������������
    
    If blnOther Then
        Call msLoadDeptTree(tvwTree, arrOther)
    End If

End Sub

Private Sub msLoadRole(ByRef tvwUser As MSComctlLib.TreeView)
    '���ؽ�ɫ��ǰ�����Ѽ��غò���
    
    Dim strSQL As String
    Dim rsRole As ADODB.Recordset
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    
    If tvwUser.Nodes.Count = 0 Then Exit Sub
    
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID FROM tb_FT_Sys_Role "
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateClosed Then GoTo LineEnd
    If rsRole.RecordCount = 0 Then GoTo LineEnd
    
    With tvwUser
        While Not rsRole.EOF
            For I = 1 To .Nodes.Count
                If .Nodes(I).Key = mKeyDept & rsRole.Fields("DeptID").Value Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, _
                        mKeyRole & rsRole.Fields("RoleAutoID").Value, rsRole.Fields("RoleName").Value, "SysRole"
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(2, lngCount)
                For K = 0 To 2
                    arrOther(K, lngCount) = rsRole.Fields(K).Value & ""
                Next
                lngCount = lngCount + 1
            End If
            
            rsRole.MoveNext
        Wend
        
        If blnOther Then
            .Nodes.Add 1, tvwChild, mOtherKey, mOtherText, "unknown"
            .Nodes(mOtherKey).Expanded = True
            For I = LBound(arrOther, 2) To UBound(arrOther, 2)
                .Nodes.Add mOtherKey, tvwChild, mKeyRole & arrOther(0, I), _
                    arrOther(1, I), "SysRole"
            Next
        End If

    End With
    
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
End Sub

Private Sub msLoadUser(ByRef tvwUser As MSComctlLib.TreeView)
    '�����û���ǰ�����Ѽ��غò���
    
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    
    If tvwUser.Nodes.Count = 0 Then Exit Sub
    
    strSQL = "SELECT UserAutoID ,UserFullName ,UserSex ,DeptID FROM tb_FT_Sys_User " & _
             "WHERE UserLoginName <>'" & gVar.AccountAdmin & "' AND UserLoginName <>'" & gVar.AccountSystem & "' "
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then GoTo LineEnd

    With tvwUser
        While Not rsUser.EOF
            For I = 1 To .Nodes.Count
                If .Nodes(I).Key = mKeyDept & rsUser.Fields("DeptID").Value Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, _
                        mKeyUser & rsUser.Fields("UserAutoID").Value, rsUser.Fields("UserFullName").Value, _
                        IIf(rsUser.Fields("UserSex") = "��", "man", "woman")
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(3, lngCount)
                For K = 0 To 3
                    arrOther(K, lngCount) = rsUser.Fields(K).Value & ""
                Next
                lngCount = lngCount + 1
            End If
            
            rsUser.MoveNext
        Wend
        
        If blnOther Then
            .Nodes.Add 1, tvwChild, mOtherKey, mOtherText, "unknown"
            .Nodes(mOtherKey).Expanded = True
            For I = LBound(arrOther, 2) To UBound(arrOther, 2)
                .Nodes.Add mOtherKey, tvwChild, mKeyUser & arrOther(0, I), _
                    arrOther(1, I), IIf(arrOther(2, I) = "��", "man", "woman")
            Next
        End If

    End With
    
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
End Sub

Private Sub msLoadUserRole(ByVal strUID As String)
    
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    Dim I As Long
    
    strSQL = "SELECT UserAutoID ,RoleAutoID FROM tb_FT_Sys_UserRole WHERE UserAutoID =" & strUID
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateOpen Then
        With TreeView2.Nodes
            For I = 2 To .Count
                If Left(.Item(I).Key, Len(mKeyRole)) = mKeyRole Then
                    If rsUser.RecordCount > 0 Then rsUser.MoveFirst
                    Do While Not rsUser.EOF
                        If .Item(I).Key = mKeyRole & rsUser.Fields("RoleAutoID") Then
                            .Item(I).Checked = True
                            Exit Do
                        End If
                        rsUser.MoveNext
                    Loop
                    If rsUser.EOF Then .Item(I).Checked = False
                Else
                     .Item(I).Checked = False
                End If
            Next
        End With
        rsUser.Close
    End If
    
    Set rsUser = Nothing
    
End Sub


Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            Combo1.Item(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub Command1_Click()
    '���
    
    Dim strLoginName As String, strPWD As String, strFullName As String
    Dim strSex As String, strMemo As String, strState As String
    Dim strDept As Variant, strPho As String, strPhotoID As String
    Dim strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset, rsPhoto As ADODB.Recordset
    
    strLoginName = Trim(Text1.Item(1).Text)
    strPWD = Trim(Text1.Item(2).Text)
    strFullName = Trim(Text1.Item(3).Text)
    strMemo = Trim(Text1.Item(4).Text)
    
    strLoginName = Left(strLoginName, 50)
    strPWD = Left(strPWD, 20)
    strFullName = Left(strFullName, 50)
    strMemo = Left(strMemo, 500)
    
    Text1.Item(1).Text = strLoginName
    Text1.Item(2).Text = strPWD
    Text1.Item(3).Text = strFullName
    Text1.Item(4).Text = strMemo
    
    If Option1.Item(0).Value Then strSex = Option1.Item(0).Caption
    If Option1.Item(1).Value Then strSex = Option1.Item(1).Caption
    If Option1.Item(2).Value Then strState = Option1.Item(2).Caption
    If Option1.Item(3).Value Then strState = Option1.Item(3).Caption
    strDept = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    If Len(strLoginName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strLoginName)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strLoginName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strLoginName)
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox Label1.Item(2).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPWD)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strPWD)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(2).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPWD)
        Exit Sub
    End If
    
    If Len(strFullName) = 0 Then
        MsgBox Label1.Item(3).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(strFullName)
        Exit Sub
    End If
    
    If Len(strSex) = 0 Then
        Option1.Item(0).Value = True
        strSex = Option1.Item(0).Caption
    End If
    If Len(strState) = 0 Then
        Option1.Item(3).Value = True
        strState = Option1.Item(3).Caption
    End If
    
    If Len(strDept) = 0 Then strDept = Null
    
    strPho = Trim(CommonDialog1.FileName)
    If Len(strPho) > 0 Then
        If Not mfCheckPhoto(strPho) Then Exit Sub
    End If
    
    If MsgBox("�Ƿ�����û���" & strLoginName & "����" & strFullName & "����", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    On Error GoTo LineERR
    
    strSQL = "SELECT UserAutoID ,UserLoginName ,UserPassword ," & _
             "UserFullName ,UserSex ,UserState ,DeptID ,UserMemo ,FileID " & _
             "From tb_FT_Sys_User " & _
             "WHERE UserLoginName = '" & strLoginName & "'"
    Set rsUser = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount > 0 Then
        strMsg = "�˺��Ѵ��ڣ��������"
        GoTo LineBrk
    Else
        '�Ȼ�ȡͷ����ϢID
        If mfCheckPhoto(strPho) Then   '������ϴ�ͼ����Ƭ������б���(��ȡͼƬID)
            strSQL = "SELECT FileID ,FileClassify ,FileExtension ,FileOldName ,FileSaveName ,FileSize ," & _
                     "FileSaveLocation ,FileUploadMen ,FileUploadTime FROM tb_FT_Lib_File   " & _
                     "WHERE FileSaveName ='" & gVar.FTUploadFileNameNew & "' AND FileSaveLocation ='" & gVar.FTUploadFileFolder & "' "
            Set rsPhoto = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
            If rsPhoto.State = adStateClosed Then GoTo LineEnd
            If rsPhoto.RecordCount > 0 Then
                strMsg = "ͷ����Ƭ��Ϣ�ڿ����Ѵ��ڣ����ٴ��ϴ���"
                GoTo LineBrk
            Else
                rsPhoto.AddNew
                rsPhoto.Fields("FileClassify") = gVar.FTUploadFileClassify
                rsPhoto.Fields("FileExtension") = gVar.FTUploadFileExtension
                rsPhoto.Fields("FileOldName") = gVar.FTUploadFileNameOld
                rsPhoto.Fields("FileSaveName") = gVar.FTUploadFileNameNew
                rsPhoto.Fields("FileSize") = gVar.FTUploadFileSize
                rsPhoto.Fields("FileSaveLocation") = gVar.FTUploadFileFolder
                rsPhoto.Fields("FileUploadMen") = gVar.UserFullName
                rsPhoto.Fields("FileUploadTime") = Now
                rsPhoto.Update
                strPhotoID = rsPhoto.Fields("FileID")    '��ȡID
                rsPhoto.Close
                strMsg = "Ϊ�û���" & strLoginName & "����" & strFullName & "������ͷ����Ƭ[" & strPhotoID & "][" & gVar.FTUploadFileNameNew & "]"
                Call gsLogAdd(Me, udInsert, "tb_FT_Lib_File", strMsg)
                Call mfSavePhoto(True)  '�ϴ�ͼƬ
            End If
        End If
        
        'Ȼ������û���Ϣ
        rsUser.AddNew
        rsUser.Fields("UserLoginName") = strLoginName
        rsUser.Fields("UserPassword") = EncryptString(strPWD, gVar.EncryptKey)
        rsUser.Fields("UserFullName") = strFullName
        rsUser.Fields("UserSex") = strSex
        rsUser.Fields("UserState") = strState
        rsUser.Fields("DeptID") = strDept
        rsUser.Fields("UserMemo") = strMemo
        rsUser.Fields("FileID") = strPhotoID
        rsUser.Update
        strMsg = rsUser.Fields("UserAutoID").Value
        Text1.Item(0).Text = strMsg
        rsUser.Close
        strMsg = "����û���" & strMsg & "����" & strLoginName & "����" & strFullName & "��"
        Call gsLogAdd(Me, udInsert, "tb_FT_Sys_User", strMsg)
        MsgBox "�û���" & strLoginName & "����" & strFullName & "����ӳɹ���", vbInformation
        Call msLoadDept(TreeView1)
        Call msLoadUser(TreeView1)
    End If
    
    GoTo LineEnd
    
LineBrk:
    If Not rsPhoto Is Nothing Then If rsPhoto.State = adStateOpen Then rsPhoto.Close
    If Not rsUser Is Nothing Then If rsUser.State = adStateOpen Then rsUser.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineERR:
    Call gsAlarmAndLog("����û��쳣")
LineEnd:
    If Not rsPhoto Is Nothing Then If rsPhoto.State = adStateOpen Then rsPhoto.Close
    If Not rsUser Is Nothing Then If rsUser.State = adStateOpen Then rsUser.Close
    Set rsPhoto = Nothing
    Set rsUser = Nothing
End Sub

Private Sub Command2_Click()
    '�޸�
    
    Dim strUID As String, strLoginName As String, strPWD As String, strState As String
    Dim strFullName As String, strSex As String, strDept As String, strMemo As String
    Dim blnLoginName As Boolean, blnPwd As Boolean, blnFullName As Boolean, blnPhoto As Boolean, blnNewPho As Boolean
    Dim blnSex As Boolean, blnDept As Boolean, blnMemo As Boolean, blnState As Boolean
    Dim strSQL As String, strMsg As String, strPho As String, strPhotoID As String, strWhrPho As String
    Dim rsUser As ADODB.Recordset, rsPhoto As ADODB.Recordset
    
    strUID = Trim(Text1.Item(0).Text)
    strLoginName = Trim(Text1.Item(1).Text)
    strPWD = Trim(Text1.Item(2).Text)
    strFullName = Trim(Text1.Item(3).Text)
    strMemo = Trim(Text1.Item(4).Text)
    strPho = Trim(CommonDialog1.FileName)
    
    strLoginName = Left(strLoginName, 50)
    strPWD = Left(strPWD, 20)
    strFullName = Left(strFullName, 50)
    strMemo = Left(strMemo, 500)
    
    Text1.Item(1).Text = strLoginName
    Text1.Item(2).Text = strPWD
    Text1.Item(3).Text = strFullName
    Text1.Item(4).Text = strMemo
    
    If Option1.Item(0).Value Then strSex = Option1.Item(0).Caption
    If Option1.Item(1).Value Then strSex = Option1.Item(1).Caption
    If Option1.Item(2).Value Then strState = Option1.Item(2).Caption
    If Option1.Item(3).Value Then strState = Option1.Item(3).Caption
    strDept = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    If Len(strLoginName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strLoginName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox Label1.Item(2).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(Text1.Item(2).Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strPWD)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(2).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPWD)
        Exit Sub
    End If
    
    If Len(strFullName) = 0 Then
        MsgBox Label1.Item(3).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(Text1.Item(3).Text)
        Exit Sub
    End If
    
    If Len(strDept) = 0 Then
        MsgBox Label1.Item(5).Caption & " ����Ϊ�գ�", vbExclamation
        Combo1.Item(0).SetFocus
        Exit Sub
    End If
    
    If Not mfCheckPhoto(strPho) Then Exit Sub
    
    If Len(strSex) = 0 Then
        Option1.Item(0).Value = True
        strSex = Option1.Item(0).Caption
    End If
    If Len(strState) = 0 Then
        Option1.Item(3).Value = True
        strState = Option1.Item(3).Caption
    End If
    
    strSQL = "SELECT UserAutoID ,UserLoginName ,UserPassword ," & _
             "UserFullName ,UserSex ,UserState ,DeptID ,UserMemo ,FileID " & _
             "From tb_FT_Sys_User " & _
             "WHERE UserAutoID = '" & strUID & "'"
    Set rsUser = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then
        strMsg = "���˺������Ϣ�Ѷ�ʧ������ϵ����Ա��"
        GoTo LineBrk
    ElseIf rsUser.RecordCount > 1 Then
        strMsg = "���˺������Ϣ�쳣������ϵ����Ա��"
        GoTo LineBrk
    Else
        If strLoginName <> rsUser.Fields("UserLoginName") Then blnLoginName = True
        If IsNull(rsUser.Fields("UserPassword")) Or strPWD <> DecryptString(rsUser.Fields("UserPassword"), gVar.EncryptKey) Then blnPwd = True
        If IsNull(rsUser.Fields("UserFullName")) Or strFullName <> rsUser.Fields("UserFullName") Then blnFullName = True
        If IsNull(rsUser.Fields("UserSex")) Or strSex <> rsUser.Fields("UserSex") Then blnSex = True
        If IsNull(rsUser.Fields("DeptID")) Or strDept <> rsUser.Fields("DeptID") Then blnDept = True
        If IsNull(rsUser.Fields("UserMemo")) Or strMemo <> rsUser.Fields("UserMemo") Then blnMemo = True
        If IsNull(rsUser.Fields("UserState")) Or strState <> rsUser.Fields("UserState") Then blnState = True
        If Len(strPho) > 0 Then blnPhoto = True
        If Not (blnLoginName Or blnPwd Or blnFullName Or blnSex Or blnState Or blnDept Or blnMemo Or blnPhoto) Then
            strMsg = "û��ʵ���ԵĸĶ����������޸ģ�"
            GoTo LineBrk
        End If
        
        strMsg = "ȷ��Ҫ�޸�" & Label1.Item(0).Caption & "��" & strUID & "�����û���Ϣ��"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then GoTo LineEnd
        
        On Error GoTo LineERR
        
        If blnLoginName Then rsUser.Fields("UserLoginName") = strLoginName
        If blnPwd Then rsUser.Fields("UserPassword") = EncryptString(strPWD, gVar.EncryptKey)
        If blnFullName Then rsUser.Fields("UserFullName") = strFullName
        If blnSex Then rsUser.Fields("UserSex") = strSex
        If blnDept Then rsUser.Fields("DeptID") = strDept
        If blnMemo Then rsUser.Fields("UserMemo") = strMemo
        If blnState Then rsUser.Fields("UserState") = strState
        
        If blnPhoto And mfCheckPhoto(strPho) Then   '������ϴ�ͼ����Ƭ������б���(��ȡͼƬID)
            strPhotoID = "" & rsUser.Fields("FileID")
            blnNewPho = IIf(Len(strPhotoID) > 0, False, True)
            strWhrPho = IIf(blnNewPho, _
                " FileSaveName ='" & gVar.FTUploadFileNameNew & "' AND FileSaveLocation ='" & gVar.FTUploadFileFolder & "' ", _
                " FileID='" & strPhotoID & "' ")
            strSQL = "SELECT FileID ,FileClassify ,FileExtension ,FileOldName ,FileSaveName ,FileSize ," & _
                     "FileSaveLocation ,FileUploadMen ,FileUploadTime FROM tb_FT_Lib_File   " & _
                     "WHERE " & strWhrPho
            Set rsPhoto = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
            If rsPhoto.State = adStateClosed Then GoTo LineEnd
            If rsPhoto.RecordCount > 1 Then
                strMsg = "ͷ����Ƭ��Ϣ�ڿ��д��ڶ��������ϵ����Ա��"
                GoTo LineBrk
            Else
                If rsPhoto.RecordCount = 0 Then blnNewPho = True    '�������ͼƬ���δ����û���Ϣ�е�ͼƬ
                If blnNewPho Then rsPhoto.AddNew
                If blnNewPho Then rsPhoto.Fields("FileClassify") = gVar.FTUploadFileClassify
                rsPhoto.Fields("FileExtension") = gVar.FTUploadFileExtension
                rsPhoto.Fields("FileOldName") = gVar.FTUploadFileNameOld
                rsPhoto.Fields("FileSaveName") = gVar.FTUploadFileNameNew
                rsPhoto.Fields("FileSize") = gVar.FTUploadFileSize
                If blnNewPho Then rsPhoto.Fields("FileSaveLocation") = gVar.FTUploadFileFolder
                rsPhoto.Fields("FileUploadMen") = gVar.UserFullName
                rsPhoto.Fields("FileUploadTime") = Now
                rsPhoto.Update
                strPhotoID = rsPhoto.Fields("FileID")    '��ȡID
                rsPhoto.Close
                If blnNewPho Then rsUser.Fields("FileID") = strPhotoID  '�������ͼ������ӵ��û���Ϣ��
                strMsg = "Ϊ�û���" & strLoginName & "����" & strFullName & "��" & IIf(blnNewPho, "����", "�޸�") & "ͷ����Ƭ[" & strPhotoID & "][" & gVar.FTUploadFileNameNew & "]"
                Call gsLogAdd(Me, IIf(blnNewPho, udInsert, udUpdate), "tb_FT_Lib_File", strMsg)
                Call mfSavePhoto(True)  '�ϴ�ͼƬ
            End If
        End If
        
        rsUser.Update
        rsUser.Close
        
        strMsg = "�޸�ID��" & strUID & "����"
        If blnLoginName Then strMsg = strMsg & "��" & Label1.Item(1).Caption & "��"
        If blnPwd Then strMsg = strMsg & "��" & Label1.Item(2).Caption & "��"
        If blnFullName Then strMsg = strMsg & "[" & Label1.Item(3).Caption & "��"
        If blnSex Then strMsg = strMsg & "��" & Label1.Item(4).Caption & "��"
        If blnDept Then strMsg = strMsg & "��" & Label1.Item(5).Caption & "��"
        If blnMemo Then strMsg = strMsg & "��" & Label1.Item(6).Caption & "��"
        If blnState Then strMsg = strMsg & "��" & Label1.Item(7).Caption & "��"
        If blnPhoto Then strMsg = strMsg & "��ͷ����Ƭ��"
        Call gsLogAdd(Me, udUpdate, "tb_FT_Sys_User", strMsg)
        
        MsgBox "�ѳɹ�" & strMsg & "��", vbInformation
        
        If blnFullName Or blnSex Or blnDept Then
            Call msLoadDept(TreeView1)
            Call msLoadUser(TreeView1)
        End If
        
    End If
    
    GoTo LineEnd
    
LineBrk:
    If Not rsPhoto Is Nothing Then If rsPhoto.State = adStateOpen Then rsPhoto.Close
    If Not rsUser Is Nothing Then If rsUser.State = adStateOpen Then rsUser.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineERR:
    Call gsAlarmAndLog("�û���Ϣ�޸��쳣")
LineEnd:
    If Not rsPhoto Is Nothing Then If rsPhoto.State = adStateOpen Then rsPhoto.Close
    If Not rsUser Is Nothing Then If rsUser.State = adStateOpen Then rsUser.Close
    Set rsPhoto = Nothing
    Set rsUser = Nothing
End Sub

Private Sub Command3_Click()
    '����
    
    Dim strUID As String, strTemp As String, strMsg As String, strSQL As String
    Dim cnUser As ADODB.Connection
    Dim rsUser As ADODB.Recordset
    Dim blnTran As Boolean
    Dim I As Long
    
    If (TreeView1.Nodes.Count = 0) Or (TreeView2.Nodes.Count = 0) Then
        MsgBox "���ȱ�֤���š��û�����ɫ���Ѿ����úã���Ϊ�û�ָ����ɫ��", vbExclamation
        Exit Sub
    End If
    strTemp = Trim(Text1.Item(5).Text)
    If (TreeView1.SelectedItem Is Nothing) Or (Len(strTemp) = 0) Then
        MsgBox "����ѡ��һ���û���", vbExclamation
        Exit Sub
    End If
    strUID = Left(strTemp, InStr(strTemp, mTwoBar) - 1)
    If Trim(strUID) <> Trim(Text1.Item(0).Text) Then
        MsgBox "�û�����쳣��������ѡ��һ���û���", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("ȷ���ԡ�" & strTemp & "������" & Command3.Caption & "��", vbQuestion + vbOKCancel, Command3.Caption & "ѯ��") = vbCancel Then Exit Sub
    
    
    Set cnUser = New ADODB.Connection
    Set rsUser = New ADODB.Recordset
    cnUser.CursorLocation = adUseClient
    
    On Error GoTo LineERR
    
    cnUser.Open gVar.ConString
    strSQL = "DELETE FROM tb_FT_Sys_UserRole WHERE UserAutoID =" & strUID
    cnUser.BeginTrans
    blnTran = True
    cnUser.Execute strSQL
    
    strSQL = "SELECT UserAutoID ,RoleAutoID FROM tb_FT_Sys_UserRole WHERE UserAutoID =" & strUID
    rsUser.Open strSQL, cnUser, adOpenStatic, adLockBatchOptimistic
    If rsUser.RecordCount > 0 Then
        strMsg = strTemp & " �ĺ�̨���ݼ���쳣�������Ի���ϵ����Ա��"
        GoTo LineERR
    Else
        With TreeView2.Nodes
            For I = 2 To .Count
                If Left(.Item(I).Key, Len(mKeyRole)) = mKeyRole Then
                    If .Item(I).Checked Then
                        rsUser.AddNew
                        rsUser.Fields("UserAutoID") = strUID
                        rsUser.Fields("RoleAutoID") = Right(.Item(I).Key, Len(.Item(I).Key) - Len(mKeyRole))
                    End If
                End If
            Next
        End With
        rsUser.UpdateBatch
        cnUser.CommitTrans
       
    End If
    
    rsUser.Close
    cnUser.Close
    Set rsUser = Nothing
    Set cnUser = Nothing
    
    Call gsLogAdd(Me, udInsertBatch, "tb_FT_Sys_UserRole", "Ϊ��" & strTemp & "��ָ����ɫ")
    MsgBox "��" & strTemp & "��" & Command3.Caption & " �ɹ���", vbInformation
    
    Exit Sub
    
LineERR:
    If blnTran Then cnUser.RollbackTrans
    If rsUser.State = adStateOpen Then rsUser.Close
    If cnUser.State = adStateOpen Then cnUser.Close
    Set rsUser = Nothing
    Set cnUser = Nothing
    If Len(strMsg) = 0 Then
        Call gsAlarmAndLog(Command3.Caption & "�쳣")
    Else
        MsgBox strMsg, vbExclamation
    End If
    
End Sub

Private Sub Command4_Click()
    'ѡ����Ƭ
    Dim strPhoto As String
    Dim lngFiveMB As Long
    
    If Len(Text1.Item(0).Text) = 0 Then
        MsgBox "��ѡ��һ���û���", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    On Error GoTo LineERR
    
    With CommonDialog1
        .DialogTitle = "ѡ��һ����Ƭ"
        .Filter = "ͼƬ(*.jpg;*.png;*.bmp)|*.jpg;*.png;*.bmp;*.gif;*.jpeg"
        .Flags = cdlOFNFileMustExist '�ļ�Ҫ����
        .ShowOpen   '�����򿪴���
        strPhoto = .FileName    '����ѡͼƬ·���ŵ�������
    End With

    lngFiveMB = 5242880     ''' 5 * 1024 * 1024 B
    If FileLen(strPhoto) > lngFiveMB Then   '�ļ����ܴ���5MB
        MsgBox "��ѡ���ͼƬ����5MB�ˣ�", vbExclamation, "�ļ�����"
        CommonDialog1.FileName = ""
        Exit Sub
    End If
    Image1.Picture = LoadPicture(strPhoto)  '����ͼƬ
    
LineERR:
    If Err.Number > 0 Then  '���쳣����
        CommonDialog1.FileName = "" '�����ЧͼƬ��·��
        Image1.Picture = LoadPicture("")
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "ͼƬ�����쳣"
    End If
End Sub

Private Sub Form_Load()

    Set Me.Icon = gWind.ImageList1.ListImages("SysUser").Picture
    Me.Caption = gWind.CommandBars1.Actions(gID.SysAuthUser).Caption
    Frame1.Item(0).Caption = Me.Caption
    Command4.ToolTipText = "��Ƭ��ͨ����" & Command1.Caption & "����" & Command2.Caption & "����ť���б���"
    Me.Timer1.Interval = 100   '100ms
    Me.Timer1.Enabled = False
    Me.Image1.Stretch = True
    
    For mlngID = Text1.LBound To Text1.UBound
        Text1.Item(mlngID).Text = ""
    Next
    
    Text1.Item(2).ToolTipText = "����ֻ�ܰ������ֻ��С��ĸ���ҳ�����20λ����"
    
    TreeView1.Nodes.Clear
    TreeView1.ImageList = gWind.ImageList1
    TreeView2.Nodes.Clear
    TreeView2.ImageList = gWind.ImageList1
    
    Call msLoadDept(TreeView1)  '�����ȼ���
    Call msLoadUser(TreeView1)  '��Ա�����
    
    Call msLoadDept(TreeView2)
    Call msLoadRole(TreeView2)
    
    Call gfLoadAuthority(Me, Command1)
    Call gfLoadAuthority(Me, Command2)
    Call gfLoadAuthority(Me, Command3)
    Call gfLoadAuthority(Me, TreeView1)
    
End Sub

Private Sub Form_Resize()
    
    Const conHeight As Long = 6500
    Const conEdge As Long = 120
    Const conTB As Long = 400
    
    If Me.WindowState <> vbMinimized Then
        If Me.Height > conHeight Then
            If Me.ScaleHeight > conEdge * 2 Then
                Frame1.Item(0).Height = Me.ScaleHeight - conEdge * 2
                Frame1.Item(1).Height = Frame1.Item(0).Height
                TreeView1.Height = Frame1.Item(0).Height - conTB
                TreeView2.Height = TreeView1.Height
                ctlMove.Height = Frame1.Item(0).Height
            End If
        End If
    End If
    
    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 16000, 9000)  'ע�ⳤ������޸�
    
End Sub

Private Sub Hsb_Change()
    ctlMove.Left = -Hsb.Value
End Sub

Private Sub Hsb_Scroll()
    Call Hsb_Change    '�������������еĻ���ʱ��ͬʱ���¶�Ӧ���ݣ�����ͬ��
End Sub

Private Sub Timer1_Timer()
    '��Ƭ�ϴ����������Ķ���
    Dim strNew As String
    
    If Not Me.Enabled Then Exit Sub '����δ���˵�����ڴ���״̬
    If Not (gArr(1).FileTransmitNotOver Or gArr(1).FileTransmitError) Then  '�����������
        If gVar.FTUploadOrDownload Then '�ϴ�״̬
            '��ҳ��������������
        Else    '����״̬
            If gfFileExist(gVar.FTDownloadFilePath) Then  'ȷ���ļ�����
                strNew = Left(gVar.FTDownloadFilePath, InStrRev(gVar.FTDownloadFilePath, "\")) & gVar.FTDownloadFileNameOld
                If gfFileReNameEx(gVar.FTDownloadFilePath, strNew) Then
                    gVar.FTDownloadFilePath = strNew
                    Call gfLoadPicture(Me.Image1, gVar.FTDownloadFilePath)
                End If
            End If
        End If
        Timer1.Enabled = False  '���ټ���ϴ�����״̬
    End If
End Sub

Private Sub Vsb_Change()
    ctlMove.Top = -Vsb.Value
End Sub

Private Sub Vsb_Scroll()
    Call Vsb_Change
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim lngLen As Long, I As Long
    Dim strKey As String, strUID As String, strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset, rsPhoto As ADODB.Recordset
    Dim strPhotoID As String
    Dim sckPho As MSWinsockLib.Winsock
    
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyUser) Then Exit Sub
    If Left(strKey, Len(mKeyDept)) = mKeyDept Then
        For mlngID = Text1.LBound To Text1.UBound
            Text1.Item(mlngID).Text = ""
        Next
        Option1.Item(1).Value = True
        Combo1.Item(0).ListIndex = -1
        Exit Sub
    End If
    If Left(strKey, Len(mKeyUser)) <> mKeyUser Then Exit Sub
    
    CommonDialog1.FileName = ""    '�����ͷ��ͼƬ��Ϣ
    Image1.Picture = LoadPicture("") '�����ͷ��ͼƬ
    
    strUID = Right(Node.Key, lngLen - Len(mKeyUser))
    strSQL = "EXEC sp_FT_Sys_UserInfo '" & strUID & "'"
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then
        strMsg = "�û���Ϣ��ʧ�ˣ�����ϵ����Ա��"
        GoTo LineBreak
    ElseIf rsUser.RecordCount > 1 Then
        strMsg = "�û���Ϣ�쳣������ϵ����Ա��"
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strUID
        Text1.Item(1).Text = rsUser.Fields("UserLoginName").Value & ""
        Text1.Item(2).Text = DecryptString(rsUser.Fields("UserPassword").Value & "", gVar.EncryptKey)
        Text1.Item(3).Text = rsUser.Fields("UserFullName").Value & ""
        Text1.Item(4).Text = rsUser.Fields("UserMemo").Value & ""
        Text1.Item(5).Text = strUID & mTwoBar & rsUser.Fields("UserFullName")
        
        Option1.Item(0).Value = IIf(rsUser.Fields("UserSex") = Option1.Item(0).Caption, True, False)
        Option1.Item(1).Value = IIf(rsUser.Fields("UserSex") = Option1.Item(1).Caption, True, False)
        Option1.Item(2).Value = IIf(rsUser.Fields("UserState") = Option1.Item(2).Caption, True, False)
        Option1.Item(3).Value = IIf(rsUser.Fields("UserState") = Option1.Item(3).Caption, True, False)
        
        If IsNull(rsUser.Fields("DeptID").Value) Then
            Combo1.Item(0).ListIndex = -1
        Else
            For I = 0 To Combo1.Item(1).ListCount - 1
                If rsUser.Fields("DeptID").Value = Combo1.Item(1).List(I) Then
                    Combo1.Item(0).ListIndex = I
                    Exit For
                End If
            Next
            If I = Combo1.Item(1).ListCount Then Combo1.Item(0).ListIndex = -1
        End If
        strPhotoID = rsUser.Fields("FileID") & ""   'ȡ����ƬID
        
        Node.SelectedImage = "SelectedMen"
        rsUser.Close
        Call msLoadUserRole(strUID) '���ؽ�ɫ�б�
        
        '����ͷ��
        If Len(strPhotoID) > 0 Then
            strSQL = "SELECT FileID ,FileClassify ,FileExtension ,FileOldName ,FileSaveName ," & _
                     "FileSize ,FileSaveLocation FROM tb_FT_Lib_File WHERE FileID ='" & strPhotoID & "' "
            Set rsPhoto = gfBackRecordset(strSQL)
            If rsPhoto.State = adStateClosed Then GoTo LineEnd
            If rsPhoto.RecordCount = 0 Then
                strMsg = "ͷ����Ƭ��Ϣ��ʧ������ϵ����Ա��"
                GoTo LineBreak
            ElseIf rsPhoto.RecordCount > 1 Then
                strMsg = "ͷ����Ƭ��Ϣ�쳣������ϵ����Ա��"
                GoTo LineBreak
            Else
                gVar.FTDownloadFileNameNew = rsPhoto.Fields("FileSaveName") & ""   '��ȡ����������е�ͼƬ�ļ���
                gVar.FTDownloadFileFolder = rsPhoto.Fields("FileSaveLocation") & ""   '��ȡ����������е��ļ�����
                gVar.FTDownloadFileSize = rsPhoto.Fields("FileSize") & ""             '��ȡ����������е��ļ���С
                gVar.FTDownloadFileNameOld = rsPhoto.Fields("FileOldName") & ""
                gVar.FTDownloadFileExtension = rsPhoto.Fields("FileExtension") & ""
                gVar.FTDownloadFileClassify = rsPhoto.Fields("FileClassify") & ""
                gVar.FTDownloadFilePath = gVar.AppPath & gVar.FTDownloadFileFolder & "\" & gVar.FTDownloadFileNameNew
            End If
            rsPhoto.Close
            If Len(gVar.FTDownloadFileNameNew) > 0 And Len(gVar.FTDownloadFileFolder) > 0 And gVar.FTDownloadFileSize > 0 Then
                '�ļ������ơ�����λ�á��ļ���С����Ϣ����ʱ����Ҫ�����ع���
                Set sckPho = gWind.Winsock1.Item(1)
                Call gsLoadFileInfo(sckPho.Index, False)    '������Ƭ������Ϣ��gArr()������
                If Not gfFolderRepair(gVar.FolderNameStore) Then GoTo LineEnd
                If sckPho.State = 7 Then    'ʹ��MDI�����ϵ�Winsock�ؼ������ļ���Ϣ
                    If gfSendInfo(gfFileInfoJoin(sckPho.Index, ftReceive), sckPho) Then
                        Debug.Print "Client���ѷ�����Ҫ[ͷ����Ƭ]��������Ϣ�������," & Now
                        Timer1.Enabled = True
                    End If
                End If
            Else
                strMsg = "����ͷ����Ƭ��Ϣ�쳣������ϵ����Ա��"
                GoTo LineBreak
            End If
        End If
    End If
    
    GoTo LineEnd
    
LineBreak:
    If Not rsPhoto Is Nothing Then If rsPhoto.State = adStateOpen Then rsPhoto.Close
    If Not rsUser Is Nothing Then If rsUser.State = adStateOpen Then rsUser.Close
    MsgBox strMsg, vbExclamation
LineEnd:
    If Not rsPhoto Is Nothing Then If rsPhoto.State = adStateOpen Then rsPhoto.Close
    If Not rsUser Is Nothing Then If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
    Set rsPhoto = Nothing
    Set sckPho = Nothing
End Sub

Private Sub TreeView2_NodeCheck(ByVal Node As MSComctlLib.Node)
    '
    Call gsNodeCheckCascade(Node, Node.Checked)
    
End Sub
