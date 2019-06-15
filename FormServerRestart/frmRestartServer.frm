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
   StartUpPosition =   1  '所有者中心
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
      Caption         =   "退出"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "本窗口将在60秒后关闭"
      BeginProperty Font 
         Name            =   "宋体"
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

Private Type gtypeValueAndErr    '用于返回布尔值的过程，顺便返回异常代号
    Result As Boolean
    ErrNum As Long
End Type

Private Enum genumFileOpenType   '打开文件方式
    udAppend    '以顺序型访问，把字符追加到文件
    udBinary    '以二进制访问
    udInput     '以顺序型访问，从文件输入字符
    udOutput    '以顺序型访问，向文件输出字符
    udRandom    '以随机方式
End Enum

Private Enum genumFileWriteType  '写入文件方式
    udPut       '用Get读出.For Binary、Random.
    udWrite     '用Input读出
    udPrint     '用Line Input 或 Input读出
End Enum


Private Sub AlarmAndLog(Optional ByVal strErr As String, Optional ByVal blnMsgBox As Boolean = True, _
        Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    '系统异常提示并写下异常日志
    Dim strMsg As String
    
    strMsg = "异常代号：" & Err.Number & vbCrLf & "异常描述：" & Err.Description
    If blnMsgBox Then MsgBox strMsg, MsgButton, strErr
    Call FileWrite(mstrFileNameErrLog, strErr & vbTab & Replace(strMsg, vbCrLf, vbTab))
    
End Sub

Private Function CloseExeFile(ByVal strName As String) As Boolean
    '关闭指定exe程序进程
    
    Dim winHwnd As Long
    Dim retVal As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    On Error GoTo LineErr
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")    'VB自带的函数
    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='" & strName & "' ")
    For Each objProcess In colProcessList
        retVal = objProcess.Terminate
        If retVal <> 0 Then Exit Function   '经观察=0时关闭进程成功，不成功时返回值不为零
    Next
    CloseExeFile = True   '全部关闭成功或不存在该进程名时
    
LineErr:
    Set objWMIService = Nothing
    Set colProcessList = Nothing
    Set objProcess = Nothing
    If Err.Number > 0 Then
        Call AlarmAndLog("关闭进程异常")
    End If
End Function

Private Sub CloseServer(ByVal strFileExe As String, ByVal strFileExePath As String)
    '关闭并重启程序
    
    If CloseExeFile(strFileExe) Then
        Call FileWrite(mstrFileNameRunLog, "关闭" & strFileExe & "程序成功")
        If ShellExePath(strFileExePath) = 0 Then
            Call FileWrite(mstrFileNameRunLog, "重启" & strFileExe & "程序失败")
        Else
            Call FileWrite(mstrFileNameRunLog, "重启" & strFileExe & "程序成功")
        End If
    Else
        Call FileWrite(mstrFileNameRunLog, "关闭" & strFileExe & "程序失败")
    End If
End Sub

Private Function FileExistEx(ByVal strPath As String) As gtypeValueAndErr
    '另一种返回值方式：来判断文件、文件目录 是否存在
    '专供后面的过程FileRepair调用
    Dim strBack As String
    
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '空字符串不算
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then
            FileExistEx.Result = True
        Else
            FileExistEx.ErrNum = -1   '不存在，也没异常
        End If
    End If

LineErr:
    If Err.Number > 0 Then
        FileExistEx.ErrNum = Err.Number   '异常了，也当作不存在了
        Call AlarmAndLog("文件判断返回异常")
    End If
End Function

Private Function FileRepair(ByVal strFile As String, Optional ByVal blnFolder As Boolean) As Boolean
    '如果 文件/文件夹 不存在 则创建
    '前提是路径的上层目录可访问
    '参数blnFolder指明传入的路径strFile是文件夹则为True，默认是文件False
    
    Dim strTemp As String
    Dim typBack As gtypeValueAndErr
    Dim lngLoc As Long
    
    If Right(strFile, 1) = "\" Then
        strFile = Left(strFile, Len(strFile) - 1)   '去掉最末的"\"
    End If
    strTemp = strFile
    If Len(strTemp) = 0 Then Exit Function          '防止传入空字符串
    
    On Error GoTo LineErr

    typBack = FileExistEx(strTemp)    '判断是否存在
    If Not typBack.Result Then          '文件不存在
        If typBack.ErrNum = -1 Then     '且无异常
            
            lngLoc = InStrRev(strTemp, "\") '判断是否有上层目录
            If lngLoc > 0 Then              '有上层目录则递归
                strTemp = Left(strTemp, lngLoc - 1) '得出上层目录的具体路径
                Call FileRepair(strTemp, True)    '递归调用自身，以保证上层目录存在
            End If

            If blnFolder Then                   '传入参数是文件夹
                MkDir strFile                   '则创建文件夹
            Else                                '传入参数是文件
                Close                           '则创建文件
                Open strFile For Random As #1
                Close
            End If
            FileRepair = True '创建成功返回True
        End If
    Else
        FileRepair = True '路径完整直接返回True
    End If

LineErr:
    Close
End Function

Private Sub FileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As genumFileOpenType = udAppend, _
    Optional ByVal WriteMode As genumFileWriteType = udPrint)
    '将指定内容以指定的方式写入指定文件中
    
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
        Case Else   '其余皆当作udAppend
            Open strFile For Append As #intNum
    End Select
    
    strTime = Format(Now, mconStrFormaty_M_dH_m_s)
    Select Case WriteMode
        Case udWrite
            Write #intNum, strTime, strContent
        Case udPut
            Put #intNum, , strTime & vbTab & strContent
        Case Else   '其余皆当作udPrint
            Print #intNum, strTime, strContent
    End Select
    Close #intNum
    
End Sub

Private Function ShellExePath(ByVal strExePath As String) As Long
    '执行EXE文件
    
    On Error Resume Next
    
    ShellExePath = Shell(strExePath)
    
End Function


Private Sub Command1_Click()
    Unload Me '退出
End Sub

Private Sub Form_Load()
    '窗口加载
    
    Dim strCmd As String
    Dim arrCmd() As String
        
    Me.Hide  '隐藏不显示该窗口
        
    Me.Timer1.Interval = 1000
    Me.Timer1.Enabled = True
        
    
    '模块变量赋值
    mstrFileExe = "FFS.exe"
    mstrAppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    mstrFileExePath = mstrAppPath & mstrFileExe
    mstrFileNameErrLog = mstrAppPath & "Data\ErrorRecord.LOG"
    mstrFileNameRunLog = mstrAppPath & "Data\RunRecord.LOG"
    
    On Error Resume Next
    
    strCmd = Trim(Command())  '获取命令行参数值
    If Len(strCmd) = 0 Then
        GoTo LineUnload
    Else
        arrCmd = Split(strCmd, " / ")
        If UCase(arrCmd(0)) <> UCase(mstrFileExe) Then
            GoTo LineUnload '命令参数中第一串字符固定为exe文件名，不是则认为非法启动该程序，不准执行
        Else
            Me.Text1.Text = mstrFileExePath
        End If
        
        If UBound(arrCmd) > 0 Then  '判断命令参数中第二个命令是否为关闭指令
            If LCase(arrCmd(1)) = "close" Then
                Call CloseServer(mstrFileExe, mstrFileExePath)
            End If
        End If
    End If

LineUnload:
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Const cMax As Long = 60 'cMax秒后退出
    Static lngCount As Long
    
    If lngCount > cMax Then  '计次已满
        Call CloseServer(mstrFileExe, mstrFileExePath)
        Call FileWrite(mstrFileNameRunLog, "限时退出" & App.EXEName & "程序")
        Unload Me
    Else
        Label1.Caption = "本窗口将在" & CStr(cMax - lngCount) & "秒后关闭" & String((lngCount Mod 4), "·")
        lngCount = lngCount + 1
    End If
End Sub
