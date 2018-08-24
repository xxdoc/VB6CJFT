VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSysUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5835
   StartUpPosition =   3  '窗口缺省
   Begin FrameFileUpdate.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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



Private Function mfConnect() As Boolean
    Dim strIP As String, strPort As String
    Static lngCount As Long
            
    lngCount = lngCount + 1
    If lngCount = 2 Then
        Call mfSetText("版本检测失败！无法连接服务器。" & vbCrLf & _
                       "请确认服务器已启动，并重新运行更新程序！", vbRed)
        Exit Function    '尝试百次后不再连接了
    End If
    
    With Winsock1.Item(1)
        If Label1(1).Caption = gVar.DisConnected Then
            strIP = GetSetting(gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyIP, gVar.TCPIP)
            strIP = gfCheckIP(strIP)

            strPort = GetSetting(gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyPort, gVar.TCPPort)
            strPort = CStr(CLng(Val(strPort)))
            If Val(strPort) > 65535 Or Val(strPort) < 0 Then strPort = gVar.TCPPort

            If .State <> 0 Then .Close
            .RemoteHost = strIP
            .RemotePort = strPort
            .Connect
            If .State = 7 Then gVar.TCPConnected = True
        End If
    End With
End Function


Private Sub Form_Load()
    
    Dim strCmd As String, arrCmd() As String
    
    Label1.Item(0).Caption = ""
    ReDim gArr(0 To 1)
    Call Main
    
    '检测是否传入命令行参数进来，没有则退出程序
    strCmd = Command
    If Len(strCmd) = 0 Then
        GoTo LineUnload '禁止直接启动更新程序，必须带命令参数
    Else
        arrCmd = Split(strCmd, gVar.CmdLineSeparator)
        
        If UCase(arrCmd(0)) <> UCase(gVar.EXENameOfClient) Then
            GoTo LineUnload    '命令参数中第一串字符固定为exe文件名，不是则认为非法启动更新程序，不准执行
        End If
        
        If UBound(arrCmd) > 0 Then  '判断命令参数中是否带否隐藏窗口命令
            If LCase(arrCmd(1)) = LCase(gVar.CmdLineParaOfHide) Then
                mblnHide = True
                Me.Hide
            End If
        End If
    End If
    
    
    Text1.BackColor = Me.BackColor
    Call mfSetLabel(gVar.ClientStateDisConnected, vbRed)
    Call mfConnect
    Timer1.Interval = 1000
    Timer1.Enabled = True

    Exit Sub
    
LineUnload:
    Unload Me   '此行以下除End Sub不可再跟任何有效代码
End Sub
