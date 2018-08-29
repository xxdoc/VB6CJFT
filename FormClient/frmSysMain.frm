VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.MDIForm frmSysMain 
   BackColor       =   &H8000000C&
   Caption         =   "FFC"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   Icon            =   "frmSysMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   2400
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1680
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   68
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":0442
            Key             =   "cNativeWinXP"
            Object.Tag             =   "2110"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":0799
            Key             =   "SysPDF"
            Object.Tag             =   "1204"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":0E6D
            Key             =   "SysXML"
            Object.Tag             =   "1205"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":142E
            Key             =   "cOffice2000"
            Object.Tag             =   "2101"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1725
            Key             =   "cOffice2003"
            Object.Tag             =   "2102"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1CB9
            Key             =   "cOfficeXP"
            Object.Tag             =   "2103"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1FE5
            Key             =   "cResource"
            Object.Tag             =   "2104"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":254C
            Key             =   "cRibbon"
            Object.Tag             =   "2105"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":298D
            Key             =   "cVisualStudio6.0"
            Object.Tag             =   "2108"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":2CD0
            Key             =   "cVisualStudio2008"
            Object.Tag             =   "2106"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":31ED
            Key             =   "cVisualStudio2010"
            Object.Tag             =   "2107"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":3604
            Key             =   "cWhidbey"
            Object.Tag             =   "2109"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":3A4F
            Key             =   "tListView"
            Object.Tag             =   "841"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":3D0D
            Key             =   "tListViewOffice2003"
            Object.Tag             =   "842"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":400D
            Key             =   "tListViewOfficeXP"
            Object.Tag             =   "843"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":42CA
            Key             =   "tNativeWinXP"
            Object.Tag             =   "844"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":4876
            Key             =   "tNativeWinXPPlain"
            Object.Tag             =   "845"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":4CF7
            Key             =   "tOffice2000"
            Object.Tag             =   "846"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5106
            Key             =   "tOffice2000Plain"
            Object.Tag             =   "847"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":551B
            Key             =   "tOffice2003"
            Object.Tag             =   "848"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5825
            Key             =   "tOffice2003Plain"
            Object.Tag             =   "849"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5B28
            Key             =   "tOfficeXPPlain"
            Object.Tag             =   "850"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5DE4
            Key             =   "tResource"
            Object.Tag             =   "851"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":60DE
            Key             =   "tShortcutBarOffice2003"
            Object.Tag             =   "852"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":63FE
            Key             =   "tToolbox"
            Object.Tag             =   "853"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":66BB
            Key             =   "tToolboxWhidbey"
            Object.Tag             =   "854"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":6A78
            Key             =   "tVisualStudio2010"
            Object.Tag             =   "855"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":6E40
            Key             =   "sCodejock"
            Object.Tag             =   "871"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":7E92
            Key             =   "sOffice2007"
            Object.Tag             =   "872"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":8EE4
            Key             =   "sOffice2010"
            Object.Tag             =   "873"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":9F36
            Key             =   "sOrangina"
            Object.Tag             =   "878"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":AF88
            Key             =   "sVista"
            Object.Tag             =   "874"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":BFDA
            Key             =   "sWinXPLuna"
            Object.Tag             =   "875"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":D02C
            Key             =   "sWinXPRoyale"
            Object.Tag             =   "876"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":E07E
            Key             =   "sZune"
            Object.Tag             =   "877"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":F0D0
            Key             =   ""
            Object.Tag             =   "901"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":F76A
            Key             =   "SysWord"
            Object.Tag             =   "1207"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":10444
            Key             =   "SysText"
            Object.Tag             =   "1206"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1111E
            Key             =   "SysExcel"
            Object.Tag             =   "1202"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":11DF8
            Key             =   "SysSearch"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":11F0A
            Key             =   "SysPageSet"
            Object.Tag             =   "1301"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1265C
            Key             =   "SysPreview"
            Object.Tag             =   "1302"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":132AE
            Key             =   "SysPrint"
            Object.Tag             =   "1303"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":13F00
            Key             =   "SysGo"
            Object.Tag             =   "116"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":14BDA
            Key             =   "SysLoginOut"
            Object.Tag             =   "1101"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":158B4
            Key             =   "SysLoginAgain"
            Object.Tag             =   "1102"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1658E
            Key             =   "SysCompany"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":171E0
            Key             =   "SysDepartment"
            Object.Tag             =   "104"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":17E32
            Key             =   "threemen"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":18A84
            Key             =   "SysUser"
            Object.Tag             =   "105"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":196D6
            Key             =   "man"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1A328
            Key             =   "woman"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1AF7A
            Key             =   "SysPassword"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1BBCC
            Key             =   ""
            Object.Tag             =   "902"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1BF1E
            Key             =   "themes"
            Object.Tag             =   "801"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1CB70
            Key             =   "SelectedMen"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1D7C2
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1D8CC
            Key             =   "SysLog"
            Object.Tag             =   "106"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1E51E
            Key             =   "SysRole"
            Object.Tag             =   "107"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1F170
            Key             =   "RoleSelect"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1FDC2
            Key             =   "SysFunc"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":20A14
            Key             =   "FuncHead"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":21666
            Key             =   "FuncSelect"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":222B8
            Key             =   "FuncControl"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":22F0A
            Key             =   "FuncButton"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":2361C
            Key             =   "FuncForm"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":2426E
            Key             =   "FuncMainMenu"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":25DC0
            Key             =   "themeSet"
            Object.Tag             =   "802"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   3480
      Top             =   2880
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   3000
      Top             =   2880
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSysMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngID As Long  '循环变量ID
Dim WithEvents mXtrStatusBar As XtremeCommandBars.StatusBar  '状态栏控件
Attribute mXtrStatusBar.VB_VarHelpID = -1
Dim mcbsPopupIcon As XtremeCommandBars.CommandBar    '托盘图标Pupup菜单



Private Sub msAddAction(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建CommandBars的Action
    
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    cbsBars.EnableActions   '启用CommandBars的Actions集合
    
'    cbsActions.Add "Id", "Caption", "TooltipText", "DescriptionText", "Category"   '范例
    With cbsActions
        .Add gID.Sys, "系统", "", "", "系统"
        
        .Add gID.SysLoginOut, "退出", "", "", ""
        .Add gID.SysLoginAgain, "重启", "", "", ""
        
        .Add gID.SysExportToCSV, "导出至CSV", "", "", ""
        .Add gID.SysExportToExcel, "导出至Excel", "", "", ""
        .Add gID.SysExportToHTML, "导出至HTML", "", "", ""
        .Add gID.SysExportToPDF, "导出至PDF", "", "", ""
        .Add gID.SysExportToText, "导出至txt", "", "", ""
        .Add gID.SysExportToWord, "导出至Word", "", "", ""
        .Add gID.SysExportToXML, "导出至XML", "", "", ""
        
        .Add gID.SysPrint, "打印", "", "", ""
        .Add gID.SysPrintPageSet, "打印页面设置", "", "", ""
        .Add gID.SysPrintPreview, "打印预览", "", "", ""
        
        .Add gID.Wnd, "窗口", "", "", "窗口"
        
        .Add gID.WndResetLayout, "重置窗口布局", "", "", ""
        
        .Add gID.WndThemeCommandBars, "工具栏主题", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2000, "Office2000", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2003, "Office2003", "", "", ""
        .Add gID.WndThemeCommandBarsOfficeXp, "OfficeXP", "", "", ""
        .Add gID.WndThemeCommandBarsResource, "Resource", "", "", ""
        .Add gID.WndThemeCommandBarsRibbon, "Ribbon", "", "", ""
        .Add gID.WndThemeCommandBarsVS2008, "VisualStudio2008", "", "", ""
        .Add gID.WndThemeCommandBarsVS2010, "VisualStudio2010", "", "", ""
        .Add gID.WndThemeCommandBarsVS6, "VisualStudio6", "", "", ""
        .Add gID.WndThemeCommandBarsWhidbey, "Whidbey", "", "", ""
        .Add gID.WndThemeCommandBarsWinXP, "WinXP", "", "", ""
        
        .Add gID.Help, "帮助", "", "", "帮助"
        .Add gID.HelpAbout, "关于…", "", "", ""
        
        .Add gID.StatusBarPane, "状态栏", "", "", ""
        .Add gID.StatusBarPaneProgress, "进度条", "", "", ""
        .Add gID.StatusBarPaneUserInfo, "用户信息", "", "", ""
        .Add gID.StatusBarPaneTime, "本机时间", "", "", ""
        .Add gID.StatusBarPaneProgressText, "进度条百分比值", "", "", ""
        .Add gID.StatusBarPaneServerButton, "服务开启/断开按钮", "", "", ""
        .Add gID.StatusBarPaneServerState, "服务状态", "", "", ""
        .Add gID.StatusBarPaneTime, "系统时间", "", "", ""
        .Add gID.StatusBarPaneIP, "本机IP地址", "", "", ""
        .Add gID.StatusBarPanePort, "连接服务器端口", "", "", ""
        .Add gID.StatusBarPaneConnectState, "连接服务器状态", "", "", ""
        .Add gID.StatusBarPaneConnectButton, "与服务器建立连接按钮", "", "", ""
        .Add gID.StatusBarPaneReStartButton, "服务自动/手动重启模式切换按钮", "", "", ""
        
        .Add gID.IconPopupMenu, "托盘图标菜单", "", "", ""
        .Add gID.IconPopupMenuMaxWindow, "最大化窗口", "", "", ""
        .Add gID.IconPopupMenuMinWindow, "最小化窗口", "", "", ""
        .Add gID.IconPopupMenuShowWindow, "显示窗口", "", "", ""
        
        .Add gID.Tool, "工具", "", "", "工具"
        .Add gID.toolOptions, "选项", "", "", "frmOption"
        
        
'        .Add gID, "", "", "", ""
        
    End With
    
    '填充cbsActions的其它属性ToolTipText、DescriptionText、Key、Category
    For Each cbsAction In cbsActions
        With cbsAction
            If .ID < 20000 Then
                .ToolTipText = .Caption
                .DescriptionText = .ToolTipText
                .Key = .Category    '为菜单时有特殊用，创建Action时窗体名保存在Category中
                .Category = cbsActions((.ID \ 1000) * 1000).Category
            End If
        End With
    Next
    
    '风格系列的cbsActions的两个属性的描述补充
    With cbsActions
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Action(mlngID).DescriptionText = .Action(gID.WndThemeCommandBars).Caption & "设置为：" & .Action(mlngID).DescriptionText
            .Action(mlngID).ToolTipText = .Action(mlngID).DescriptionText
        Next
    End With
    
    Set cbsAction = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddDesignerControls(ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars自定义对话框中内容项的添加
    
    Dim cbsControls As XtremeCommandBars.CommandBarControls
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    Set cbsControls = cbsBars.DesignerControls
    For Each cbsAction In cbsActions
        If cbsAction.ID < 20000 Then
            cbsControls.Add xtpControlButton, cbsAction.ID, ""
        End If
    Next
    
    Set cbsControls = Nothing
    Set cbsAction = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddKeyBindings(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建快捷键
    
    With cbsBars.KeyBindings
        .AddShortcut gID.SysLoginOut, "F10"
    End With
    
End Sub

Private Sub msAddMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建菜单栏
    
    Dim cbsMenuBar As XtremeCommandBars.MenuBar
    Dim cbsMenuMain As XtremeCommandBars.CommandBarPopup
    Dim cbsMenuCtrl As XtremeCommandBars.CommandBarControl
    
    Set cbsMenuBar = cbsBars.ActiveMenuBar
    cbsMenuBar.ShowGripper = False  '不显示可拖动的那个点点标记
    cbsMenuBar.EnableDocking xtpFlagStretched     '菜单栏独占一行且不能主动拖动
    
    '系统主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Sys, "")
    With cbsMenuMain.CommandBar.Controls
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysExportToCSV, "")
        cbsMenuCtrl.BeginGroup = True
        For mlngID = gID.SysExportToExcel To gID.SysExportToWord
            .Add xtpControlButton, mlngID, ""
        Next
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysPrintPageSet, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysPrintPreview, ""
        .Add xtpControlButton, gID.SysPrint, ""
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysLoginAgain, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysLoginOut, ""
        
    End With
    
    '窗口主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    With cbsMenuMain.CommandBar.Controls
        '重置布局
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.WndResetLayout, "")
        cbsMenuCtrl.BeginGroup = True
        
        '特殊ID35001自定义工具栏
        Set cbsMenuCtrl = .Add(xtpControlButton, XTP_ID_CUSTOMIZE, "自定义工具栏...")
        cbsMenuCtrl.BeginGroup = True
    
        '特殊ID59392工具栏列表
        Set cbsMenuCtrl = .Add(xtpControlPopup, 0, "工具栏列表")
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
        
        'CommandBars工具栏主题子菜单
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeCommandBars, "")
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
                .Add xtpControlButton, mlngID, ""
            Next
        End With
    End With
    
    '工具菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Tool, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.toolOptions, ""
    
    '帮助主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Help, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.HelpAbout, ""
    
    Set cbsMenuBar = Nothing
    Set cbsMenuMain = Nothing
    Set cbsMenuCtrl = Nothing
End Sub

Private Sub msAddPopupMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建托盘图标右键弹出式菜单
        
    Set mcbsPopupIcon = cbsBars.Add(cbsBars.Actions(gID.IconPopupMenu).Caption, xtpBarPopup)
    With mcbsPopupIcon.Controls
        .Add xtpControlButton, gID.IconPopupMenuMaxWindow, ""
        .Add xtpControlButton, gID.IconPopupMenuMinWindow, ""
        .Add xtpControlButton, gID.IconPopupMenuShowWindow, ""
        .Add xtpControlButton, gID.SysLoginAgain, ""
        .Add xtpControlButton, gID.SysLoginOut, ""
    End With
End Sub

Private Sub msAddToolBar(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建工具栏
    
    Dim cbsBar As XtremeCommandBars.CommandBar
    Dim cbsCtr As XtremeCommandBars.CommandBarControl
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    
    '系统操作工具栏
    Set cbsBar = cbsBars.Add(cbsActions(gID.Sys).Caption, xtpBarTop)
    With cbsBar.Controls
        For mlngID = gID.SysLoginOut To gID.SysLoginAgain
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
        For mlngID = gID.SysExportToExcel To gID.SysExportToWord
            If mlngID <> gID.SysExportToHTML Then
                Set cbsCtr = .Add(xtpControlButton, mlngID, "")
                cbsCtr.BeginGroup = True
            End If
        Next
        For mlngID = gID.SysPrintPageSet To gID.SysPrint
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '工具栏主题
    Set cbsBar = cbsBars.Add(cbsActions(gID.WndThemeCommandBars).Caption, xtpBarTop)
    With cbsBar.Controls
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    Set cbsBar = Nothing
    Set cbsCtr = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddXtrStatusBar(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建状态栏
    
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    Dim BarPane As XtremeCommandBars.StatusBarPane
    
    Set cbsActions = cbsBars.Actions
    Set mXtrStatusBar = cbsBars.StatusBar
    With mXtrStatusBar
        .AddPane 0      '系统Pane，显示CommandBarActions的Description
        .SetPaneStyle 0, SBPS_STRETCH
        .SetPaneText 0, "Hello"
        .IdleText = "Hello"
        
        .AddPane gID.StatusBarPaneUserInfo
        .SetPaneText gID.StatusBarPaneUserInfo, "中华人民"
        .FindPane(gID.StatusBarPaneUserInfo).Width = 60
        
        .AddPane gID.StatusBarPaneIP
        .SetPaneText gID.StatusBarPaneIP, Me.Winsock1.Item(1).LocalIP  'gVar.TCPSetIP
        .FindPane(gID.StatusBarPaneIP).Width = 90
        
        .AddPane gID.StatusBarPanePort
        .SetPaneText gID.StatusBarPanePort, gVar.TCPSetPort
        .FindPane(gID.StatusBarPanePort).Width = 60
        
        .AddPane gID.StatusBarPaneConnectState
        .SetPaneText gID.StatusBarPaneConnectState, gVar.ClientStateDisConnected
        .FindPane(gID.StatusBarPaneConnectState).Width = 60
        
'''        .AddPane gID.StatusBarPaneReStartButton
'''        .SetPaneText gID.StatusBarPaneReStartButton, IIf(gVar.ParaBlnAutoReStartServer, "自", "手") & "动重启服务模式"
'''        .FindPane(gID.StatusBarPaneReStartButton).Width = 120
'''        .FindPane(gID.StatusBarPaneReStartButton).BackgroundColor = vbCyan
'''        .FindPane(gID.StatusBarPaneReStartButton).Button = True
        
'''        .AddPane gID.StatusBarPaneServerState
'''        .FindPane(gID.StatusBarPaneServerState).Text = gVar.ServerStateNotStarted
'''        .FindPane(gID.StatusBarPaneServerState).Width = 60
        
'''        .AddPane gID.StatusBarPaneServerButton
'''        .FindPane(gID.StatusBarPaneServerButton).Text = gVar.ServerButtonStart
'''        .FindPane(gID.StatusBarPaneServerButton).Width = 60
'''        .FindPane(gID.StatusBarPaneServerButton).Button = True
        
        .AddProgressPane gID.StatusBarPaneProgress
                
        .AddPane gID.StatusBarPaneProgressText
        .SetPaneText gID.StatusBarPaneProgressText, "0%"
        .FindPane(gID.StatusBarPaneProgressText).Width = 60
        
        .AddPane 59137  'CapsLock键的状态
        .AddPane 59138  'NumLK键的状态
        .AddPane 59139  'ScrLK键的状态
        .FindPane(0).Caption = "Idle Text"
        .FindPane(59137).Caption = "Caps Lock键状态"
        .FindPane(59138).Caption = "Num LocK键状态"
        .FindPane(59139).Caption = "Scroll LocK键状态"
        
        .Visible = True
        .EnableCustomization True
    End With
    
    For Each BarPane In mXtrStatusBar     '设置Caption、ToolTip、Alignment属性
        If Len(BarPane.Caption) = 0 Then BarPane.Caption = cbsActions(BarPane.ID).Caption
        BarPane.ToolTip = BarPane.Caption
        If BarPane.ID <> 0 Then BarPane.Alignment = xtpAlignmentCenter
    Next
    
    Set cbsActions = Nothing
    Set BarPane = Nothing
End Sub

Private Sub msConnectToServer(ByRef sckCon As MSWinsockLib.Winsock, Optional ByVal blnConnect As Boolean = False)
    '启动与服务器的连接
    
    If Not blnConnect Then Exit Sub
    With sckCon
        If .State <> 0 Then .Close
        .RemoteHost = gVar.TCPSetIP
        .RemotePort = gVar.TCPSetPort
        .Connect
    End With
End Sub

Private Sub msLeftClick(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars单击命令响应公共过程
    
    Dim strKey As String
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    With gID
        Select Case CID
            Case .WndThemeCommandBarsOffice2000 To .WndThemeCommandBarsWinXP
                Call gsThemeCommandBar(CID, cbsBars)
            Case .WndResetLayout
                Call msResetLayout(cbsBars)
                
            Case .SysLoginAgain
                If MsgBox("确定重新启动客户端程序吗？", vbQuestion + vbOKCancel, "重启主程序询问") = vbOK Then
                    Call msUnloadMe(True)
                    Load Me
                End If
            Case .SysLoginOut
                If MsgBox("确定退出客户端程序吗？", vbQuestion + vbOKCancel, "关闭主程序询问") = vbOK Then
                    Call msUnloadMe(True)
                End If
                
            Case .IconPopupMenuMaxWindow
                Me.WindowState = vbMaximized
                Me.Show
            Case .IconPopupMenuMinWindow
                Me.WindowState = vbMinimized
            Case .IconPopupMenuShowWindow
                Me.WindowState = vbNormal
                Me.Show
                
            Case .HelpAbout
                Dim strAbout As String
                strAbout = "名称：" & App.Title & vbCrLf & _
                           "版本：" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                           "版权所有：XMH"
                MsgBox strAbout, vbInformation, "关于" & App.Title
            
            Case .SysExportToCSV To .SysExportToWord, .SysPrintPageSet To .SysPrint
                If Screen.ActiveControl Is Nothing Then Exit Sub
                If Not (TypeOf Screen.ActiveControl Is FlexCell.Grid) Then Exit Sub
                Select Case CID
                    Case .SysExportToCSV To .SysExportToXML
                        Call gsGridExportTo(Screen.ActiveControl, CID)
                    Case .SysExportToText
                        If MsgBox("是否将当前表格内容导出至txt文本文档？", vbQuestion + vbYesNo, "询问") = vbYes Then Call gsGridToText(Screen.ActiveControl)
                    Case .SysExportToWord
                        If MsgBox("是否将当前表格内容导出至Word文档？", vbQuestion + vbYesNo, "询问") = vbYes Then Call gsGridToWord(Screen.ActiveControl)
                        
                    Case .SysPrint
                        If MsgBox("确定打印当前表格内容吗？", vbQuestion + vbOKCancel, "打印询问") = vbOK Then Call gsGridPrint
                    Case .SysPrintPreview
                        Call gsGridPrintPreview
                    Case .SysPrintPageSet
                        Call gsGridPageSet
                End Select
                
            Case Else
                strKey = LCase(cbsActions.Action(CID).Key)
                If Left(strKey, 3) = "frm" Then
                    If cbsActions.Action(CID).Enabled Then
                        Select Case CID
                            Case .toolOptions
                                Call gsOpenTheWindow(strKey, vbModal, vbNormal)
                            Case Else
                                Call gsOpenTheWindow(strKey)
                        End Select
                    End If
                Else
                    MsgBox "【" & cbsActions(CID).Caption & "】命令未定义！", vbExclamation, "命令警告"
                End If
        End Select
    End With
    
    Set cbsActions = Nothing
End Sub

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    '从注册表中加载参数值至公用变量中
    Dim tempVal
    
    If Not blnLoad Then Exit Sub
    
    On Error Resume Next    '加/解密函数过程可能有异常
    With gVar
        .ParaBlnWindowCloseMin = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, 1))    '关闭时最小化
        .ParaBlnWindowMinHide = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, 1))  '最小化时隐藏
        
        .TCPDefaultIP = Me.Winsock1.Item(0).LocalIP '本机IP地址
        .TCPSetIP = gfCheckIP(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPDefaultIP)) '要连接服务端IP地址
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , .TCPDefaultPort, 10000, 65535) '要连接的服务器端口
        
        .ParaBlnAutoReStartServer = Val(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyParaAutoReStartServer, 1))   '手动/自动重启服务模式
        .ParaBlnAutoStartupAtBoot = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, 0))  '开机自动启动
        
        .UserComputerName = VBA.Environ("ComputerName")
        .UserLoginName = VBA.Environ("UserName") '"XiaoMing"
        .UserFullName = "小明"
        
'''        '由服务端发过来给客户端
'''        .ConSource = gfCheckIP(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .TCPSetIP))   '服务器名称/IP
'''        .ConDatabase = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString("dbTest", .EncryptKey)), .EncryptKey)    '数据库名
'''        .ConUserID = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString("123", .EncryptKey)), .EncryptKey)  '登陆名
'''        .ConPassword = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString("888888", .EncryptKey)), .EncryptKey)    '登陆密码
        
        
    End With
End Sub

Private Sub msResetLayout(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '重置窗口布局：CommandBars与Dockingpane控件重置
    
    Dim cBar As XtremeCommandBars.CommandBar
    Dim L As Long, T As Long, R As Long, b As Long

    For Each cBar In cbsBars
    Debug.Print cBar.BarID, cBar.Title, cBar.Type
        cBar.Reset
        cBar.Visible = True
    Next
    
    For mlngID = 2 To cbsBars.Count
        cbsBars.GetClientRect L, T, R, b
        cbsBars.DockToolBar cbsBars(mlngID), 0, b, xtpBarTop
    Next
    
    Set cBar = Nothing
End Sub

Private Sub msSetClientState(ByVal ColorSet As Long)
    '设置状态栏中客户端连接服务端的状态
    
    Dim paneState As XtremeCommandBars.StatusBarPane
    
    Set paneState = mXtrStatusBar.FindPane(gID.StatusBarPaneConnectState)
    With paneState
        If ColorSet = vbGreen Then  '已连接
            .Text = gVar.ClientStateConnected
            .BackgroundColor = ColorSet
            gVar.TCPStateConnected = True   '用变量记录状态
        Else    '其它
            gVar.TCPStateConnected = False
            If ColorSet = vbRed Then    '连接异常
                .Text = gVar.ClientStateConnectError
                .BackgroundColor = ColorSet
            Else    '未连接等
                .Text = gVar.ClientStateDisConnected
                .BackgroundColor = vbYellow
            End If
        End If
    End With
    Set paneState = Nothing
    
End Sub

Private Sub msUnloadMe(Optional ByVal blnUnload As Boolean = True)
    '卸载窗体
    If Not blnUnload Then Exit Sub
    gVar.CloseWindow = True
    Unload Me
End Sub


Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '命令单击事件
    Call msLeftClick(Control.ID, Me.CommandBars1)
End Sub

Private Sub CommandBars1_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'CommandBars控件的Action状态的切换
    
    Dim blnFC As Boolean    '判断是否为FC表格
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = Me.CommandBars1.Actions
    If Screen.ActiveControl Is Nothing Then
        blnFC = False
    Else
        blnFC = TypeOf Screen.ActiveControl Is FlexCell.Grid    '当前活动控件是FC表格
    End If
    With gID
        For mlngID = .SysExportToCSV To .SysExportToWord
            cbsActions(mlngID).Enabled = blnFC  '活动控件是FC表格则激活对应Action，否则使其不可用
        Next
        For mlngID = .SysPrintPageSet To .SysPrint
            cbsActions(mlngID).Enabled = blnFC
        Next
    End With
End Sub

Private Sub MDIForm_Load()
    '窗体加载
    
    Dim cbsBars As XtremeCommandBars.CommandBars
    Dim strUpdate As String
    
    ReDim gArr(1)   '初始化数组。客户端统一使用下标1，包括Timer1控件与Winsocket控件
    Timer1.Item(1).Interval = 1000  '计时器循环时间
    Set gWind = Me  '指定主窗体给全局引用对象
    
    XtremeCommandBars.CommandBarsGlobalSettings.App = App '一个默认设置
    Set cbsBars = Me.CommandBars1
    
    Call Main   '初始化全局公用变量
    Call msLoadParameter(True)  '加载配置参数
    Call msAddAction(cbsBars)   '创建Actions集合
    Call msAddMenu(cbsBars)     '创建菜单栏
    Call msAddToolBar(cbsBars)  '创建工具栏
    Call msAddPopupMenu(cbsBars)    '创建托盘图标的菜单
    Call msAddXtrStatusBar(cbsBars) '创建状态栏
    Call msAddKeyBindings(cbsBars)  '添加快捷键,放到LoadCommandBars方法后面才能生效？？？
    Call msAddDesignerControls(cbsBars) 'CommandBars自定义对话框中使用的
    
    cbsBars.AddImageList ImageList1         '使CommandBars控件匹配ImageList控件中图标
    cbsBars.EnableCustomization True        '允许CommandBars自定义，此属性最好放在所有CommandBars设定之后
    cbsBars.Options.UpdatePeriod = 250      '更改CommandBars的Update事件的执行周期，默认100ms
    
    Call gsLoadSkin(Me, Me.SkinFramework1, sMSO7, True)  '加载窗口主题
    
    '加载工具栏主题
    Call gsThemeCommandBar(Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientCommandbarsTheme, gID.WndThemeCommandBarsRibbon)), cbsBars)
    
    '注册表信息加载-CommandBars设置
    Call cbsBars.LoadCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    
    Set cbsBars = Nothing   '销毁使用完的对象
    
    Call gsFormSizeLoad(Me, False) '注册表信息加载-窗口位置大小
    
    Call msConnectToServer(Me.Winsock1.Item(1), True)      '与务器建立连接
    
    strUpdate = gVar.AppPath & gVar.EXENameOfUpdate & " " & gVar.EXENameOfClient & _
            gVar.CmdLineSeparator & gVar.CmdLineParaOfHide      '隐式打开更新检测程序
    If Not gfShell(strUpdate) Then
        Rem MsgBox "更新程序启动异常！", vbExclamation, "警告"
        Rem Debug.Print strUpdate
        Call gsAlarmAndLog("更新程序启动异常", False)
    End If
    
    '检查是否为试用版*******************************
    '==============================================
    
    
    '==============================================
    
    Call gfNotifyIconAdd(Me)    '添加托盘图标
    Me.Hide
    
    If LCase(App.EXEName & ".exe") <> LCase(gVar.EXENameOfClient) Then
        MsgBox "不可擅自修改可执行的应用程序文件名！", vbCritical, "严重警报"
        Call msUnloadMe(True)    '防止exe文件名被改
    End If
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '响应托盘图标左键、右键动作，托盘菜单
    Dim sngMsg As Single
    
    If Y <> 0 Then Exit Sub    '似乎此句可限制住鼠标一定是在托盘图标上，不是在窗体上
    sngMsg = X / Screen.TwipsPerPixelX
    Select Case sngMsg
        Case WM_RBUTTONUP
            mcbsPopupIcon.ShowPopup  '右键弹出Popup菜单
        Case WM_LBUTTONDBLCLK   '左键双击托盘图标时 窗口最显示/最小化 切换
            Rem If Button <> vbLeftButton Then Exit Sub '窗口不明原因地偶尔自动最小化了，似乎此句可限制住
            With Me
                If .WindowState = vbMinimized Then
                    .WindowState = vbNormal
                    .Show
                    .SetFocus
                Else
                    .WindowState = vbMinimized
                End If
            End With
        Case Else
    End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '判断是否真正要关闭窗口
    
    If gVar.ParaBlnWindowCloseMin Then
        If Not gVar.CloseWindow Then
            Cancel = True
            Me.WindowState = vbMinimized
        End If
        gVar.CloseWindow = False
    Else
        If Not gVar.CloseWindow Then
            If MsgBox("是否最小化窗口？", vbQuestion + vbYesNo, "关闭或最小化") = vbYes Then
                Cancel = True
                Me.WindowState = vbMinimized
            End If
        End If
    End If
End Sub

Private Sub MDIForm_Resize()
    '窗口最小化提示
    If Me.Visible And Me.WindowState = vbMinimized Then
        If gVar.ParaBlnWindowMinHide Then
            Me.Hide
            Call gfNotifyIconBalloon(Me, "最小化到系统托盘图标啦", "最小化提示")
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    '卸载窗体时保存信息
    Dim resetNotifyIconData As gtypeNOTIFYICONDATA
    
    '保存注册表信息-CommandBars设置
    Call Me.CommandBars1.SaveCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    
    Call gsFormSizeSave(Me, False) '保存注册表信息-窗口位置大小
    Call gsSaveCommandbarsTheme(Me.CommandBars1, False)   '保存CommandBars的风格主题
    
    gVar.CloseWindow = False    '清除状态-关闭窗口
    gVar.ClientLoginShow = False '清除状态-显示登陆窗口
    Set gVar.rsURF = Nothing    '清除权限信息
    Call SkinFramework1.LoadSkin("", "")    '清空皮肤
    Set mXtrStatusBar = Nothing  '清除状态栏
    Set mcbsPopupIcon = Nothing '清除Popup菜单
    Call gfNotifyIconDelete(Me) '删除托盘图标
    gNotifyIconData = resetNotifyIconData   '清空托盘气泡信息。否则重启程序时会自动弹出？而且只能放上句删除托盘图标语句的后面?
    gArr(1) = gArr(0)
'    Erase gArr
    Set gWind = Nothing '清除全局窗体引用
    
End Sub

Private Sub Timer1_Timer(Index As Integer)
    Const conCon As Byte = 1    '连接状态检测间隔conConn秒
    Static byteCon As Byte
    Static byteChk As Byte
    
    byteCon = byteCon + 1
    byteChk = byteChk + 1
    
    If byteCon >= conCon Then
        With Me.Winsock1.Item(1)
            If .State = 7 Then  '已连接
                Call msSetClientState(vbGreen)
                If Not gVar.ClientLoginShow Then '弹出登陆窗口
                    If gVar.TCPStateConnected Then
                        If gArr(Index).Connected Then
                            frmSysLogin.Show vbModeless, Me
                            gVar.ClientLoginShow = True
                        End If
                    End If
                End If
            ElseIf .State = 9 Then  '连接异常
                Call msSetClientState(vbRed)
            Else    '未连接等
                Call msSetClientState(vbYellow)
            End If
        End With
        byteCon = 0 '清零静态累积变量
    End If
    
    If byteChk > (gVar.TCPWaitTime + 2) Then  '因为服务器端也是等待gVar.TCPWaitTime才断开连接，这里延迟一点
        byteChk = 0
        If Not gVar.TCPStateConnected Then
            Call gsAlarmAndLogEx("与服务器建立连接失败，请确认服务端程序已启动", "连接警示")
            Call msUnloadMe(True)
        End If
    End If
    
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '连接关闭时清空传输信息
    If UBound(gArr) = 1 Then gArr(1) = gArr(0)
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strGet As String
    Dim byteGet() As Byte
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '字符信息传输状态↓
            
            Me.Winsock1.Item(Index).GetData strGet  '接收字符
            
            If InStr(strGet, gVar.PTClientConfirm) > 0 Then '收到要回复服务端确认连接的信息
                Call gfSendInfo(gVar.PTClientIsTrue, Me.Winsock1.Item(Index))
                gArr(Index).Connected = True
            ElseIf InStr(strGet, gVar.PTConnectIsFull) Then '收到服务端发来的连接数已满
                Me.Timer1.Item(Index).Enabled = False
                MsgBox "客户端与服务端连接数受限，请其他用户退出后再试！", vbCritical, "连接数已满警告"
                Call msUnloadMe(True)
                End
                
            ElseIf InStr(strGet, gVar.PTConnectTimeOut) Then '连续连接时间已到
                Me.Timer1.Item(Index).Enabled = False
                MsgBox "与服务器连续连接时间已到，请重新登陆！", vbExclamation, "连接时间限制提示"
                Call msUnloadMe(True)
                Load Me
                
            ElseIf InStr(strGet, gVar.PTFileStart) > 0 Then '可以发送文件给服务端了的状态
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index)) '发送文件给服务端
                Call gsFormEnable(Me, False)    '禁止客户端再操作
                
            ElseIf InStr(strGet, gVar.PTFileExist) > 0 Then '服务端发来客户端想要的文件存在
                Dim strSize As String, lngInstrSize As Long
                
                lngInstrSize = InStr(strGet, gVar.PTFileSize) '获取客户端想要的存在于服务端的文件的大小
                If lngInstrSize > 0 Then
                    strSize = Mid(strGet, lngInstrSize + Len(gVar.PTFileSize))
                    If IsNumeric(strSize) Then
                        .FileSizeTotal = Val(strSize)
                        Call gfSendInfo(gVar.PTFileSend, Me.Winsock1.Item(Index)) '通知服务端可以发送过来了
                        '初始化进度条
                        
                        .FileTransmitState = True
                        Call gsFormEnable(Me, False)    '禁止客户端再操作
                    End If
                End If
                
            ElseIf InStr(strGet, gVar.PTFileNoExist) > 0 Then   '
                MsgBox "需要文件<" & .FileName & ">在服务端不存在！", vbExclamation, "文件警告"
                gArr(Index) = gArr(0)
                
            End If
            
            Debug.Print "Client GetInfo:" & strGet, bytesTotal
            '字符信息传输状态↑
        Else
            '文件传输状态↓
            
            If .FileNumber = 0 Then '申请文件号
                .FileNumber = FreeFile
                Open .FilePath For Binary As #.FileNumber
            End If
            
            ReDim byteGet(bytesTotal - 1)   '重定义数组大小
            Me.Winsock1.Item(Index).GetData byteGet, vbArray + vbByte   '接收文件信息并放入数组
            Put #.FileNumber, , byteGet '保存进文件中
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal    '记录已传输大小
            '更新进度条
            
            If .FileSizeCompleted >= .FileSizeTotal Then    '传输完成后的一些处理
                Close #.FileNumber
                Call gsFormEnable(Me, True) '解除客户端的限制
                gArr(Index) = gArr(0)
                Call gfSendInfo(gVar.PTFileEnd, Me.Winsock1.Item(Index)) '发送结束标志
                Debug.Print "Client Received Over"
            End If
            
            '文件传输状态↑
        End If
    End With
    
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '连接异常处理
    
    If Index <> 0 Then
        If gArr(Index).FileTransmitState Then   '异常时清空文件传输信息
            Debug.Print "ClientWinsockError:" & Index & "--" & Err.Number & "  " & Err.Description
            Close
            gArr(Index) = gArr(0)
            Call gsFormEnable(Me, True)
            Call gsAlarmAndLog("与服务器连接发生异常", False)
        End If
    End If
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    '发送完处理
    
    If Index = 0 Then Exit Sub
    With gArr(Index)
        If .FileTransmitState Then
            If .FileSizeCompleted < .FileSizeTotal Then
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index))
            Else
                gArr(Index) = gArr(0)
                Call gsFormEnable(Me, True)
                Debug.Print "Client Send File Over"
            End If
        End If
    End With
End Sub
