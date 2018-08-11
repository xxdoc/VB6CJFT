Attribute VB_Name = "modSub"
Option Explicit


Public Sub Main()
    
    Dim strTemp As String
    
    '主窗体CommandBars的ID值初始化
    With gID
        .Sys = 1000
        
        .SysLoginOut = 1101
        .SysLoginAgain = 1102
        .SysAuthChangePassword = 1103
        .SysAuthDepartment = 1104
        .SysAuthUser = 1105
        .SysAuthRole = 1106
        .SysAuthFunc = 1107
        .SysAuthLog = 1108
        
        .SysExportToExcel = 1201
        .SysExportToPDF = 1202
        .SysExportToText = 1203
        .SysExportToWord = 1204
        .SysExportToXML = 1205
        
        .SysPrint = 1303
        .SysPrintPageSet = 1301
        .SysPrintPreview = 1302
                
        
        .Wnd = 2000
        
        .WndResetLayout = 2050
        
        .TabWorkspacePopupMenu = 2051
        
        .WndThemeCommandBars = 2100
        .WndThemeCommandBarsOffice2000 = 2101
        .WndThemeCommandBarsOffice2003 = 2102
        .WndThemeCommandBarsOfficeXp = 2103
        .WndThemeCommandBarsResource = 2104
        .WndThemeCommandBarsRibbon = 2105
        .WndThemeCommandBarsVS2008 = 2106
        .WndThemeCommandBarsVS2010 = 2107
        .WndThemeCommandBarsVS6 = 2108
        .WndThemeCommandBarsWhidbey = 2109
        .WndThemeCommandBarsWinXP = 2110

        .WndThemeTaskPanel = 2200
        .WndThemeTaskPanelListView = 2201
        .WndThemeTaskPanelListViewOffice2003 = 2202
        .WndThemeTaskPanelListViewOfficeXP = 2203
        .WndThemeTaskPanelNativeWinXP = 2204
        .WndThemeTaskPanelNativeWinXPPlain = 2205
        .WndThemeTaskPanelOffice2000 = 2206
        .WndThemeTaskPanelOffice2000Plain = 2207
        .WndThemeTaskPanelOffice2003 = 2208
        .WndThemeTaskPanelOffice2003Plain = 2209
        .WndThemeTaskPanelOfficeXPPlain = 2210
        .WndThemeTaskPanelResource = 2211
        .WndThemeTaskPanelShortcutBarOffice2003 = 2212
        .WndThemeTaskPanelToolbox = 2213
        .WndThemeTaskPanelToolboxWhidbey = 2214
        .WndThemeTaskPanelVisualStudio2010 = 2215
        
        .WndSon = 2300
        .WndSonCloseAll = 2301
        .WndSonCloseCurrent = 2302
        .WndSonCloseLeft = 2303
        .WndSonCloseOther = 2304
        .WndSonCloseRight = 2305
        .WndSonVbAllBack = 2306
        .WndSonVbAllMin = 2307
        .WndSonVbArrangeIcons = 2308
        .WndSonVbCascade = 2309
        .WndSonVbTileHorizontal = 2310
        .WndSonVbTileVertical = 2311
        
        
        .WndThemeSkin = 2400
        .WndThemeSkinCodejock = 2401
        .WndThemeSkinOffice2007 = 2402
        .WndThemeSkinOffice2010 = 2403
        .WndThemeSkinVista = 2404
        .WndThemeSkinWinXPLuna = 2405
        .WndThemeSkinWinXPRoyale = 2406
        .WndThemeSkinZune = 2407
        
        .WndThemeSkinSet = 2450
        
        
        .Help = 3000
        .HelpAbout = 3101
        .HelpDocument = 3102
        .HelpUpdate = 3103
        
        
        
        '''***请将所有菜单栏中的【菜单】的CommandBrs的ID值设置在20000以下*******************
        
        
        .Pane = 21000
        
        .PaneIDFirst = 21101
        .PaneTitleFirst = 21102
        
        .PanePopupMenu = 21103
        .PanePopupMenuAutoFoldOther = 21104
        .PanePopupMenuExpandALL = 21105
        .PanePopupMenuFoldALL = 21106
        
        
        .StatusBarPane = 22000
        
        .StatusBarPaneConnectButton = 22101
        .StatusBarPaneConnectState = 22102
        .StatusBarPaneProgress = 22103
        .StatusBarPaneProgressText = 22104
        .StatusBarPaneServerButton = 22105
        .StatusBarPaneServerState = 22106
        .StatusBarPaneTime = 22107
        .StatusBarPaneUserInfo = 22108
        
        .IconPopupMenu = 23000
        .IconPopupMenuMaxWindow = 23101
        .IconPopupMenuMinWindow = 23102
        .IconPopupMenuShowWindow = 23103
        
    End With
    
    '公用变量值初始化
    With gVar
        
        .TCPSetConnectMax = 20
        .TCPSetIP = "127.0.0.1"
        .TCPSetPort = 9898
        
        .FTChunkSize = 5734
        .FTWaitTime = 5
        
        .ServerClose = "关闭服务"
        .ServerError = "异常"
        .ServerNotStarted = "未启动"
        .ServerStart = "开启服务"
        .ServerStarted = "已启动"
        
        .StateConnected = "已连接"
        .StateDisConnected = "未连接"
        .StateConnectError = "连接异常"
        .StateConnectToServer = "建立连接"
        .StateDisConnectFromServer = "断开连接"
        
        .PTFileName = "<FileName>"
        .PTFileSize = "<FileSize>"
        .PTFileFolder = "<FileFolder>"
        .PTFileStart = "<FileStart>"
        .PTFileEnd = "<FileEnd>"
        .PTFileSend = "<FileSend>"
        .PTFileReceive = "<FileReceive>"
        .PTFileExist = "<FileExist>"
        .PTFileNoExist = "<FileNoExist>"
        
        .PTVersionNeedUpdate = "<VersionNeedUpdate>"
        .PTVersionNotUpdate = "<VersionNotUpdate>"
        .PTVersionOfClient = "<VersionOfClient>"
        
        .PTClientConfirm = "<ClientConfirm>"
        .PTClientIsTrue = "<ClientIsTrue>"
        .PTWaitTime = 2
        
        .EXENameOfClient = "FFC.exe"
        .EXENameOfServer = "FFS.exe"
        .EXENameOfSetup = "FFSetup.exe"
        .EXENameOfUpdate = "FFU.exe"
        
        .CmdLineParaOfHide = "Hide"
        .CmdLineSeparator = " / "
        
        .RegAppName = "FF"
        .RegKeyTCPIP = "IP"
        .RegKeyTCPPort = "Port"
        .RegSectionTCP = "TCP"
        
        .RegSectionSkin = "SkinFile"
        .RegKeySkinFile = "SkinRes"
        
        .RegSectionServer = "Server"
        .RegKeyServerAccount = "ServerAccount"
        .RegKeyServerIP = "ServerIP"
        .RegKeyServerPassword = "ServerPassword"
        
        .RegSectionUser = "UserInfo"
        .RegKeyUserLast = "LastLoginUser"
        .RegKeyUserList = "LoginUserList"
        
        .RegSectionSettings = "Settings"
        .RegKeyCommandBars = "cbs"
        .RegKeyWindowHeight = "WindowHeight"
        .RegKeyWindowLeft = "WindowLeft"
        .RegKeyWindowTop = "WindowTop"
        .RegKeyWindowWidth = "WindowWidth"
        .RegKeyCommandbarsTheme = "cbsTheme"
        
        .RegTrailPath = "SoftWare\Common\Section"   'HKEY_CURRENT_USER\SoftWare\……
        .RegTrailKey = "Key"
        .TrailPeriod = 15
        
        .RegKeyParaWindowMinHide = "WindowMinHide"
        
        .AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
        
        .FolderNameBin = .AppPath & "Bin\"
        .FolderNameData = .AppPath & "Data\"
        .FolderNameTemp = .AppPath & "Temp\"
        
        .FileNameErrLog = .FolderNameData & "ErrorRecord.LOG"
        .FileNameSkin = ""
        .FileNameSkinIni = ""
        
        .AccountAdmin = "Admin"     '两个特殊用户
        .AccountSystem = "System"   '两个特殊用户
        
        .FuncButton = "按钮"
        .FuncControl = "其它"
        .FuncForm = "窗口"
        .FuncMainMenu = "主菜单"
        
        .WindowHeight = 8700
        .WindowWidth = 15800
        
        '''*****在注册表中保存服务器地址、访问的账号与密码****
        strTemp = GetSetting(.RegAppName, .RegSectionServer, .RegKeyServerIP)
        .ConSource = gfCheckIP(strTemp)
        
        strTemp = GetSetting(.RegAppName, .RegSectionServer, .RegKeyServerAccount, "")
        If Len(strTemp) > 0 Then strTemp = gfDecryptSimple(strTemp)
        .ConUserID = strTemp
        
        strTemp = GetSetting(.RegAppName, .RegSectionServer, .RegKeyServerPassword, "")
        If Len(strTemp) > 0 Then strTemp = gfDecryptSimple(strTemp)
        .ConPassword = strTemp
        
        .ConDatabase = "db_Test"    '暂仅限连接SQLServer2008 OR 2012 数据库
        .ConString = "Provider=SQLOLEDB;Persist Security Info=False;Data Source=" & .ConSource & _
                    ";UID=" & .ConUserID & ";PWD=" & .ConPassword & _
                    ";DataBase=" & .ConDatabase & ";"   '''在64位系统上Data Source中间要空格隔开才能建立连接
        
    End With
    
End Sub


Public Sub gsAlarmAndLog(Optional ByVal strErr As String, Optional ByVal blnMsgBox As Boolean = True, Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    '异常提示并写下异常日志
    
    Dim strMsg As String
    
    strMsg = "异常代号：" & Err.Number & vbCrLf & "异常描述：" & Err.Description
    If blnMsgBox Then MsgBox strMsg, MsgButton, strErr
    Call gsFileWrite(gVar.FileNameErrLog, strErr & vbTab & Replace(strMsg, vbCrLf, vbTab))
    
End Sub


Public Sub gsFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As genumFileOpenType = udAppend, _
    Optional ByVal WriteMode As genumFileWriteType = udPrint)
    '将指定内容以指定的方式写入指定文件中
    
    Dim intNum As Integer
    Dim strTime As String
    
    If Not gfFileRepair(strFile) Then Exit Sub
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
    
    strTime = Format(Now, "yyyy-MM-dd hh:mm:ss")
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


Public Sub gsFormScrollBar(ByRef frmCur As Form, ByRef ctlMv As Control, _
    ByRef Hsb As HScrollBar, ByRef Vsb As VScrollBar, _
    Optional ByVal lngMW As Long = 12000, _
    Optional ByVal lngMH As Long = 9000, _
    Optional ByVal lngHV As Long = 255)
    
    'frmCur：滚动条所在的窗体
    'ctlMv：窗体中的控件（除滚动条以外）都在此容器控件中
    'Hsb：窗体frmCur中水平滚动条控件
    'Vsb：窗体frmCur中垂直滚动条控件
    'lngMW：窗体不出现滚动条的宽度
    'lngMH：窗体不出现滚动条的高度
    'lngHV：滚动条的窄边宽度或高度。
    '***注意注意注意：滚动条控件需最后添加至窗体中，且不能放在容器控件ctlMv中*******
    
    Dim lngW As Long
    Dim lngH As Long
    Dim lngSW As Long
    Dim lngSH As Long
    Dim lngMin As Long
    
    lngW = frmCur.Width
    lngH = frmCur.Height
    lngSW = frmCur.ScaleWidth
    lngSH = frmCur.ScaleHeight
    lngMin = -120
    
    On Error Resume Next
    
    If lngW >= lngMW Then
        Hsb.Visible = False
        ctlMv.Left = -lngMin
    Else
        With Hsb
            .Move 0, lngSH - lngHV, lngSW, lngHV
            .Min = lngMin
            .Max = lngMW - lngW + lngHV
            .SmallChange = 10
            .LargeChange = 50
            .Visible = True
        End With
    End If
    
    If lngH >= lngMH Then
        Vsb.Visible = False
        ctlMv.Top = -lngMin
    Else
        With Vsb
            .Move lngSW - lngHV, 0, lngHV, IIf(Hsb.Visible, lngSH - lngHV, lngSH)
            .Min = lngMin
            .Max = lngMH - lngH + lngHV
            .SmallChange = 10
            .LargeChange = 50
            .Visible = True
        End With
    End If
    
'    '在窗体中添加窗口控件ctlMove，将所有其它控件放入此容器中，然
'    '后添加名称分别为Hsb\Vsb的水平\垂直滚动条在窗体中，最好留到最后放入窗体中
'    '然后在窗体中添加以下事件调用即可
'Private Sub Form_Resize()
'    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 12000, 9000)  '注意长、宽的修改
'End Sub
'Private Sub Hsb_Change()
'    ctlMove.Left = -Hsb.Value
'End Sub
'
'Private Sub Hsb_Scroll()
'    Call Hsb_Change    '当滑动滚动条中的滑块时会同时更新对应内容，以下同。
'End Sub
'
'Private Sub Vsb_Change()
'    ctlMove.Top = -Vsb.Value
'End Sub
'
'Private Sub Vsb_Scroll()
'    Call Vsb_Change
'End Sub

End Sub

Public Sub gsFormSizeLoad(ByRef frmLoad As Form)
    '从注册表中加载窗口的位置与大小信息
    Dim Left As Long, Top As Long, Width As Long, Height As Long
    
    Left = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowLeft, 0))
    If Left < 0 Or Left > Screen.Width Then Left = 0
    Top = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowTop, 0))
    If Top < 0 Or Left > Screen.Height Then Top = 0
    Width = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowWidth, gVar.WindowWidth))
    If Width <= 0 Or Width > Screen.Width Then Width = gVar.WindowWidth
    Height = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowHeight, gVar.WindowHeight))
    If Height <= 0 Or Height > Screen.Height Then Height = gVar.WindowHeight
    frmLoad.Move Left, Top, Width, Height
    
End Sub

Public Sub gsFormSizeSave(ByRef frmSave As Form)
    '保存窗口的位置与大小信息至注册表中
    Dim Left As Long, Top As Long, Width As Long, Height As Long
    
    With frmSave
        Left = .Left
        Top = .Top
        Width = .Width
        Height = .Height
        If Left < 0 Or Left > Screen.Width Then Left = 0
        If Top < 0 Or Top > Screen.Height Then Top = 0
        If Width > Screen.Width Then Width = gVar.WindowWidth
        If Height > Screen.Height Then Height = gVar.WindowHeight
    End With
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowLeft, CStr(Left))
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowTop, CStr(Top))
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowWidth, CStr(Width))
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowHeight, CStr(Height))
    
End Sub

Public Sub gsGridPageSet()
    '打印页面设置
    
    Dim gridControl As Control
    Dim blnFlexCell As Boolean
    
    If Screen.ActiveForm Is Nothing Then GoTo LineBreak
    If Screen.ActiveForm.ActiveControl Is Nothing Then GoTo LineBreak
    
    Set gridControl = Screen.ActiveForm.ActiveControl
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    If blnFlexCell Then
'''        frmSysPageSet.Show vbModal   '内容较多暂不设置
        gridControl.PrintDialog
    Else
        GoTo LineBreak
    End If
        
    Exit Sub

LineBreak:
    MsgBox "页面设置检测异常，请重试！", vbExclamation
    
End Sub

Public Sub gsGridPrint()
    '打印表格内容
    
    Call gsGridPrintPreview
    
End Sub

Public Sub gsGridPrintPreview()
    '预览表格内容
    
    Dim gridControl As Control
    Dim blnFlexCell As Boolean
    
    If Screen.ActiveForm Is Nothing Then GoTo LineBreak
    If Screen.ActiveForm.ActiveControl Is Nothing Then GoTo LineBreak
    
    Set gridControl = Screen.ActiveForm.ActiveControl
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True

    If blnFlexCell Then
        With gridControl
            With .PageSetup
                .PrintFixedColumn = True
                .PrintFixedRow = True
                .PrintGridlines = True
                .Footer = "第 &P 页 共 &N 页"
                .FooterAlignment = cellCenter
            End With
            .PrintPreview
        End With
    Else
        GoTo LineBreak
    End If
        
    Exit Sub

LineBreak:
    MsgBox "预览页面检测异常，请重试！", vbExclamation
    
End Sub

Public Sub gsGridToExcel(ByRef gridControl As Control, Optional ByVal TimeCol As Long = -1, Optional ByVal TimeStyle As String = "yyyy-MM-dd HH:mm:ss")  '导出至Excel
    '将表格控件中的内容导出至Excel中
    '参数TimeCol：为控件中的时间列的列号，TimeStyle设定格式
    '最好引用Excel对象。运行时电脑上应有MSOFFICE软件。
    
'    Dim xlsOut As Excel.Application    '用这个申明好编程但要引用，编完后改为Object
    Dim xlsOut As Object
'    Dim sheetOut As Excel.Worksheet
    Dim sheetOut  As Object
    Dim blnFlexCell As Boolean
    Dim R As Long, C As Long, I As Long, J As Long
    
    On Error Resume Next
    Screen.MousePointer = 13
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    Set xlsOut = CreateObject("Excel.Application")
    xlsOut.Workbooks.Add
    Set sheetOut = xlsOut.ActiveSheet
    
    With gridControl
        R = .Rows
        C = .Cols
        '表格内容复制到Excel中
        If blnFlexCell Then
            For I = 0 To R - 1
                For J = 0 To C - 1
                    sheetOut.Cells(I + 1, J + 1) = .Cell(I, J).Text
                Next
            Next
        Else
            For I = 0 To R - 1
                For J = 0 To C - 1
                    sheetOut.Cells(I + 1, J + 1) = .TextMatrix(I, J)
                Next
            Next
        End If
    End With
    
    With sheetOut
        If TimeCol > -1 Then
            .Columns(TimeCol + 1).NumberFormatLocal = TimeStyle
        End If
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Bold = True '加粗显示(第一行默认标题行)
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Size = 12   '第一行12号字大小
        .Range(.Cells(2, 1), .Cells(R, C)).Font.Size = 10   '第二行以后10号字大小
        .Range(.Cells(1, 1), .Cells(R, C)).HorizontalAlignment = -4108  'xlCenter= -4108(&HFFFFEFF4)   '居中显示
        .Range(.Cells(1, 1), .Cells(R, C)).Borders.Weight = 2   'xlThin=2  '单元格显示黑色线宽
        .Columns.EntireColumn.AutoFit   '自动列宽
        .Rows(1).RowHeight = 23 '第一行行高
    End With
    
    xlsOut.Visible = True   '显示Excel文档
    
    Set sheetOut = Nothing
    Set xlsOut = Nothing
    Screen.MousePointer = 0
    
End Sub


Public Sub gsGridToText(ByRef gridControl As Control)
    '将传入的表格控件中的内容导出为文本文件
    
    Dim strFileName As String
    Dim blnFlexCell As Boolean
    Dim intFree As Integer
    Dim R As Long, C As Long, I As Long, J As Long
    Dim strTxt As String
    
    For I = 1 To 8
        strFileName = strFileName & gfBackOneChar(udNumber + udUpperCase) '文件名中的8个随机字符，不含小写字母
    Next
    strFileName = gVar.FolderNameData & Format(Now, "yyyyMMddHHmmss_") & strFileName & ".txt"
    If Not gfFileRepair(strFileName) Then
        MsgBox "创建文件失败，请重试！", vbExclamation, "文件生成警告"
        Exit Sub
    End If
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    intFree = FreeFile
    Open strFileName For Output As #intFree
    With gridControl
        R = .Rows - 1
        C = .Cols - 1
        If blnFlexCell Then
            For I = 0 To R
                strTxt = ""
                For J = 0 To C
                    strTxt = strTxt & .Cell(I, J).Text & vbTab
                Next
                Print #intFree, strTxt
            Next
        End If
    End With
    
    Close
    
    Call gfFileOpen(strFileName)    '打开
    
End Sub


Public Sub gsGridToWord(ByRef gridControl As Control)
    '将指定表格中的内容导出至Word文档中
    
'    Dim wordApp As Word.Application
    Dim wordApp As Object
'    Dim docOut As Word.Document
    Dim docOut As Object
'    Dim tbOut As Word.Table
    Dim tbOut As Object
    Dim lngRows As Long, lngCols As Long
    Dim I As Long, J As Long
    Dim blnFlexCell As Boolean
    
    lngRows = gridControl.Rows
    lngCols = gridControl.Cols
    
    On Error Resume Next
'    Set wordApp = New Word.Application
    Set wordApp = CreateObject("Word.Application")
    Set docOut = wordApp.Documents.Add()
    Set tbOut = docOut.Tables.Add(docOut.Range, lngRows, lngCols, True)
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    If blnFlexCell Then
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.Cell(I, J).Text
            Next
        Next
    Else
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.TextMatrix(I, J)
            Next
        Next
    End If
    tbOut.Rows(1).Range.Bold = True             '第一行内容加粗
    tbOut.Range.ParagraphFormat.Alignment = 1   '表格内容居中显示
    Call tbOut.AutoFitBehavior(1)               '根据内容自动调整列宽
    
    wordApp.Visible = True
    
    Set tbOut = Nothing
    Set docOut = Nothing
    Set wordApp = Nothing
    
End Sub

Public Sub gsLoadAuthority(ByRef frmCur As Form, ByRef ctlCur As Control)
    '加载窗口中的控制权限
    
    Dim strUser As String, strForm As String, strCtlName As String
    
    strUser = LCase(gVar.UserLoginName)
    strForm = LCase(frmCur.Name)
    strCtlName = LCase(ctlCur.Name)
    
    If strUser = LCase(gVar.AccountAdmin) Or strUser = LCase(gVar.AccountSystem) Then Exit Sub
    ctlCur.Enabled = False
    
    With gVar.rsURF
        If .State = adStateOpen Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If strForm = LCase(.Fields("FuncFormName")) Then
                        If strCtlName = LCase(.Fields("FuncName")) Then
                            ctlCur.Enabled = True
                            Exit Do
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    
End Sub

Public Sub gsLoadSkin(ByRef frmCur As Form, ByRef skFRM As XtremeSkinFramework.SkinFramework, _
    Optional ByVal lngResource As genumSkinResChoose, Optional ByVal blnFromReg As Boolean)
    '加载主题
    Dim lngReg As Long, strRes As String, strIni As String
    
    lngReg = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinFile, 0)
    If blnFromReg Then  '如果从注册表中获取资源文件，则按注册表中值修改lngResource的值
        If lngReg > sMSO10 Or lngReg < sNone Then lngReg = sNone
        lngResource = lngReg
    End If
    
    Select Case lngResource '选择窗口风格资源文件
        Case sMSO7
            strRes = gVar.FolderNameBin & "cjstylesO7.dll"
            strIni = "NormalBlue.ini"   'NormalBlue LightBlue NormalBlack NormalSilver NormalAqua
        Case sMSO10
            strRes = gVar.FolderNameBin & "cjstylesO10.dll"
            strIni = "NormalBlue.ini"   'NormalBlue NormalBlack NormalSilver
        Case sMSVst
            strRes = gVar.FolderNameBin & "cjstylesOvst.dll"
            strIni = "NormalBlue.ini"   'NormalBlue NormalBlack NormalSilver NormalBlack2
        Case Else
    End Select
    
    With skFRM
        .LoadSkin strRes, strIni
'''        .ApplyOptions = .ApplyOptions Or xtpSkinApplyMetrics Or xtpSkinApplyMenus   '全部应用
        .ApplyOptions = xtpSkinApplyMenus Or xtpSkinApplyColors Or xtpSkinApplyMetrics  '如果添加xtpSkinApplyFrame，鼠标滚轮不能控制FC表格滚动条
        .ApplyWindow frmCur.hwnd
    End With
    
    If lngReg <> lngResource Then Call SaveSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinFile, lngResource)
    
End Sub

Public Sub gsLogAdd(ByRef frmCur As Form, Optional ByVal LogType As genumLogType = udSelect, _
    Optional ByVal strTable As String = "", Optional ByVal strContent As String = "")
    '添加操作日志
    
    Dim strType As String
    Dim strSQL As String
    Dim rsLog As ADODB.Recordset
    
    strType = gfBackLogType(LogType)
    
    strSQL = "EXEC sp_Test_Sys_LogAdd '" & strType & "','" & frmCur.Name & "," & frmCur.Caption & "','" & strTable & _
             "','" & strContent & "','" & gVar.UserLoginName & "," & gVar.UserFullName & "','" & gVar.UserLoginIP & "','" & gVar.UserComputerName & "'"
'Debug.Print strSQL
    Set rsLog = gfBackRecordset(strSQL, , adLockOptimistic)
    If rsLog.State = adStateOpen Then rsLog.Close
    Set rsLog = Nothing
    
End Sub

Public Sub gsNodeCheckCascade(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '结点的Checked属性级联变化
    
    If blnCheck Then    '=False时不处理
        Call gsNodeCheckUp(nodeCheck)
    End If
    
    Call gsNodeCheckDown(nodeCheck, blnCheck)
    
End Sub

Public Sub gsNodeCheckDown(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '不/勾选结点的所有子结点
    
    Dim nodeSon As MSComctlLib.Node
    Dim C As Long, K As Long
    
    C = nodeCheck.Children
    If C > 0 Then
        For K = 1 To C
            If K = 1 Then
                Set nodeSon = nodeCheck.Child
            Else
                Set nodeSon = nodeSon.Next
            End If
            If nodeSon.Checked <> blnCheck Then nodeSon.Checked = blnCheck
            If nodeSon.Children > 0 Then
                Call gsNodeCheckDown(nodeSon, blnCheck)
            End If
        Next
    End If
    
End Sub

Public Sub gsNodeCheckUp(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean = True)
    '勾选结点的所有父结点
    
    Dim nodeDad As MSComctlLib.Node
    
    If Not nodeCheck.Parent Is Nothing Then
        Set nodeDad = nodeCheck.Parent
        If Not nodeDad.Checked Then nodeDad.Checked = blnCheck
        If Not nodeDad.Parent Is Nothing Then
            Call gsNodeCheckUp(nodeDad)
        End If
    End If
    
End Sub


Public Sub gsOpenTheWindow(ByVal strFormName As String, _
    Optional ByVal OpenMode As FormShowConstants = vbModeless, _
    Optional ByVal FormWndState As FormWindowStateConstants = vbMaximized)
    '以指定窗口模式OpenMode与窗口FormWndState状态来打开指定窗体strFormName
    
    Dim frmOpen As Form
    Dim C As Long
    
    strFormName = LCase(strFormName)
    If gfFormLoad(strFormName) Then
        For C = 0 To Forms.Count - 1
            If LCase(Forms(C).Name) = strFormName Then
                Set frmOpen = Forms(C)
                Exit For
            End If
        Next
    Else
        Set frmOpen = Forms.Add(strFormName)
    End If
    
    frmOpen.WindowState = FormWndState
    frmOpen.Show OpenMode               '此句放最后，不能放上句前面，否则退出程序时MDI窗体不能完全关闭，可能因为CommandBars控件的原因。
        
End Sub

Public Sub gsSaveCommandbarsTheme(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '保存CommandBars的风格主题
    Dim lngID As Long
    
    For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        If cbsBars.Actions(lngID).Checked Then Exit For
    Next
    If lngID > gID.WndThemeCommandBarsWinXP Then lngID = gID.WndThemeCommandBarsRibbon
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyCommandbarsTheme, lngID)
    
End Sub

Public Sub gsStartUpSet(Optional ByVal blnSet As Boolean = True)
    
    '开机自启动设置
    Dim strReg As String, strCur As String
    Dim blnReg As Boolean
    
    If Not blnSet Then Exit Sub
    strCur = Chr(34) & gVar.AppPath & App.EXEName & ".exe" & Chr(34) & "-s"
    blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strReg, RegRead)
    If blnReg Then
        If LCase(strCur) <> LCase(strReg) Then
            blnReg = False
'''Debug.Print LCase(strCur),LCase(strReg)
        End If
    End If
    If Not blnReg Then
        blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strCur, RegWrite)
        If Not blnReg Then
            '记录设置开机自动启动失败
            Call gsAlarmAndLog("设置开机自动启动失败！")
        End If
    End If
    
End Sub

Public Sub gsThemeCommandBar(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars风格设置
    Dim lngTheme As Long, lngID As Long
    Dim blnChangeSkin As Boolean
    
    Select Case CID
        Case gID.WndThemeCommandBarsOffice2000
            lngTheme = xtpThemeOffice2000
        Case gID.WndThemeCommandBarsOffice2003
            lngTheme = xtpThemeOffice2003
            blnChangeSkin = True
        Case gID.WndThemeCommandBarsOfficeXp
            lngTheme = xtpThemeOfficeXP
        Case gID.WndThemeCommandBarsResource
            lngTheme = xtpThemeResource
            blnChangeSkin = True
        Case gID.WndThemeCommandBarsRibbon
            lngTheme = xtpThemeRibbon: blnChangeSkin = True
        Case gID.WndThemeCommandBarsVS2008
            lngTheme = xtpThemeVisualStudio2008
        Case gID.WndThemeCommandBarsVS2010
            lngTheme = xtpThemeVisualStudio2010
        Case gID.WndThemeCommandBarsVS6
            lngTheme = xtpThemeVisualStudio6
        Case gID.WndThemeCommandBarsWhidbey
            lngTheme = xtpThemeWhidbey
        Case Else   'gID.WndThemeCommandBarsWinXP
            lngTheme = xtpThemeNativeWinXP
    End Select
    
    cbsBars.VisualTheme = lngTheme
    
    For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        cbsBars.Actions(lngID).Checked = False
    Next
    cbsBars.Actions(CID).Checked = True
    
    If blnChangeSkin Then   '更改对应窗口主题使颜色统一
        Call gsLoadSkin(gWind, gWind.SkinFramework1, sMSO7)
    Else
        Call gsLoadSkin(gWind, gWind.SkinFramework1, sMSVst)
    End If
    
End Sub

Public Sub gsUnCheckedAction(ByVal strFormName As String, cbsBars As XtremeCommandBars.CommandBars)
    '当窗口关闭时，去掉主窗体中CommandBars控件中被勾选的对应Action
    
    Dim actionCur As XtremeCommandBars.CommandBarAction
    
    strFormName = LCase(strFormName)
    For Each actionCur In cbsBars.Actions
        If Len(actionCur.Key) > 0 Then
            If LCase(actionCur.Key) = strFormName Then
                actionCur.Checked = False
                Exit For
            End If
        End If
    Next
    
End Sub


