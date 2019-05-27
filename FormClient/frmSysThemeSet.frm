VERSION 5.00
Begin VB.Form frmSysThemeSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "窗体主题设置"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6585
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "恢复默认主题"
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
      Caption         =   "字体、颜色选择："
      Height          =   180
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主题选择："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主题文件路径："
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
    '恢复默认主题
    If MsgBox("是否将窗体主题恢复成默认值？", vbQuestion + vbYesNo, "确认提示") = vbNo Then
        Exit Sub
    Else
        Call gsLoadSkin(gWind, gWind.SkinFramework1, -1)
        List1.Item(1).ListIndex = -1
        List1.Item(2).ListIndex = -1
    End If
End Sub

Private Sub Command3_Click()
    '退出窗口
    Unload Me
End Sub

Private Sub Form_Load()
    '窗口加载
    
    Dim skinDes As SkinDescription
    Dim skinDesAll As SkinDescriptions
    Dim strFPath As String, strFName As String
    Dim strRegRes As String, strRegIni As String
    Dim L As Long, M As Long
    
    Text1.Text = gVar.FolderNameBin
    
    Set skinDesAll = gWind.SkinFramework1.EnumerateSkinDirectory(gVar.FolderNameBin, False) '枚举出文件夹下所有资源文件
    If skinDesAll.Count > 0 Then
        List1.Item(1).Clear
        For Each skinDes In skinDesAll  '加载主题文件到列表中
            strFPath = skinDes.Path
            strFName = Right(strFPath, Len(strFPath) - InStrRev(strFPath, "\"))
            List1.Item(1).AddItem strFName
        Next
        
        strRegRes = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinRes, "")
        strRegIni = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinIni, "")
        strRegRes = Mid(strRegRes, InStrRev(strRegRes, "\") + 1)    '去掉路径保留文件名
        strRegIni = Mid(strRegIni, InStrRev(strRegIni, "\") + 1)
        If Len(strRegRes) > 0 Then
            For L = 0 To List1.Item(1).ListCount - 1    '定位当前窗口主题
                If LCase(strRegRes) = LCase(List1.Item(1).List(L)) Then
                    List1.Item(1).ListIndex = L
                    If Len(strRegIni) > 0 Then
                        If List1.Item(2).ListCount > 0 Then
                            For M = 0 To List1.Item(2).ListCount - 1
                                If LCase(strRegIni) = LCase(List1.Item(2).List(M)) Then
                                    List1.Item(2).ListIndex = M
                                    Exit For    '退出循环
                                End If
                            Next
                        End If
                    End If
                    Exit For    '退出循环
                End If
            Next
            
        End If
    End If
    
End Sub

Private Sub List1_Click(Index As Integer)
    '主题资源选择
    Dim skinDes As SkinDescription
    Dim skinIni As SkinIniFile
    Dim strRes As String, strIni As String
        
    If Index = 1 Then
        If List1.Item(1).ListIndex = -1 Then Exit Sub   '列表为空时点击无效
        
        Set skinDes = gWind.SkinFramework1.EnumerateSkinFile(gVar.FolderNameBin & List1.Item(1).Text) '枚举出该主题文件下所有配置文件
        If skinDes.Count > 0 Then
            List1.Item(2).Clear
            For Each skinIni In skinDes '枚举出该主题资源文件的子文件至第二个列表中
                List1.Item(2).AddItem skinIni.IniFileName
            Next
            List1.Item(2).ListIndex = 0 '并默认选中第一个子文件
        End If
        
    End If
    
    If Index = 2 Then
        If Me.Visible Then   '保存所选
            strRes = gVar.FolderNameBin & List1.Item(1).List(List1.Item(1).ListIndex)
            strIni = gVar.FolderNameBin & List1.Item(2).List(List1.Item(2).ListIndex)
            Call gsLoadSkin(gWind, gWind.SkinFramework1, -1, False, strRes, strIni, False)
        End If
    End If
    
    Set skinDes = Nothing
    Set skinIni = Nothing
End Sub
