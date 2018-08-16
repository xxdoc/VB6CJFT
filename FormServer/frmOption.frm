VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
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
   StartUpPosition =   1  '所有者中心
   Begin FlexCell.Grid Grid1 
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3201
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
    
    '从公共变量或注册表中加载配置信息
    With Me.Grid1
        .Cell(2, 1).Text = gVar.ParaBlnWindowCloseMin   '关闭时最小化
        .Cell(2, 5).Text = gVar.ParaBlnWindowMinHide    '最小化时隐藏
        
    
    End With
    
End Sub

Private Sub msSaveParameter(Optional ByVal blnSave As Boolean = True)
    
    If Not blnSave Then Exit Sub
    
    '参数值更新至公共变量
    With Grid1
        gVar.ParaBlnWindowCloseMin = .Cell(2, 1).Text
        gVar.ParaBlnWindowMinHide = .Cell(2, 5).Text

    End With
    
    '参数值通过公用变量保存进注册表中
    With gVar
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, IIf(.ParaBlnWindowMinHide, 1, 0))
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, IIf(.ParaBlnWindowCloseMin, 1, 0))
        
    End With
    
    If MsgBox("参数保存完成！是否现在退出窗口？", vbInformation + vbYesNo, "提示") = vbYes Then Unload Me
    
End Sub


Private Sub Form_Load()
    Dim strFile As String
    
    strFile = gVar.FolderNameBin & "OptionWindow.cel"
    If Not gfFileExist(strFile) Then
        MsgBox "以下配置文件加载失败，请解决后再重新打开窗口。" & vbCrLf & strFile, vbCritical, "异常提示"
        Exit Sub
    End If
    With Grid1
        .AutoRedraw = False
        .OpenFile (strFile)
        .Appearance = Flat
        .Column(0).Width = 0
        .RowHeight(0) = 0
        .ExtendLastCol = True
        .GridColor = vbWhite
        .BorderColor = Me.BackColor
        .BackColorBkg = Me.BackColor
        
        Call msLoadParameter(True)
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub Form_Resize()
    Grid1.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub Grid1_HyperLinkClick(ByVal Row As Long, ByVal Col As Long, URL As String, Changed As Boolean)
    '保存设置值
    
    URL = ""
    Changed = True
    If Row <> (Grid1.Rows - 1) Then Exit Sub
    
    If Col = 1 Then '保存
        If MsgBox("确定保存所有参数值吗？", vbQuestion + vbOKCancel, "保存询问") = vbOK Then Call msSaveParameter(True)
    ElseIf Col = 5 Then '退出
        Unload Me
    End If
End Sub
