VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.DockingPane.v15.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.TaskPanel.v15.3.1.ocx"
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
   Begin VB.PictureBox Picture0 
      Align           =   1  'Align Top
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   10095
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10155
      Begin VB.PictureBox Picture1 
         Height          =   855
         Left            =   1920
         ScaleHeight     =   795
         ScaleWidth      =   1755
         TabIndex        =   1
         Top             =   120
         Width           =   1815
         Begin XtremeTaskPanel.TaskPanel TaskPanel1 
            Height          =   615
            Left            =   480
            TabIndex        =   2
            Top             =   120
            Width           =   735
            _Version        =   983043
            _ExtentX        =   1296
            _ExtentY        =   1085
            _StockProps     =   64
            ItemLayout      =   2
            HotTrackStyle   =   1
         End
      End
   End
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
            Object.Tag             =   "2201"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":3D0D
            Key             =   "tListViewOffice2003"
            Object.Tag             =   "2202"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":400D
            Key             =   "tListViewOfficeXP"
            Object.Tag             =   "2203"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":42CA
            Key             =   "tNativeWinXP"
            Object.Tag             =   "2204"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":4876
            Key             =   "tNativeWinXPPlain"
            Object.Tag             =   "2205"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":4CF7
            Key             =   "tOffice2000"
            Object.Tag             =   "2206"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5106
            Key             =   "tOffice2000Plain"
            Object.Tag             =   "2207"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":551B
            Key             =   "tOffice2003"
            Object.Tag             =   "2208"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5825
            Key             =   "tOffice2003Plain"
            Object.Tag             =   "2209"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5B28
            Key             =   "tOfficeXPPlain"
            Object.Tag             =   "2210"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5DE4
            Key             =   "tResource"
            Object.Tag             =   "2211"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":60DE
            Key             =   "tShortcutBarOffice2003"
            Object.Tag             =   "2212"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":63FE
            Key             =   "tToolbox"
            Object.Tag             =   "2213"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":66BB
            Key             =   "tToolboxWhidbey"
            Object.Tag             =   "2214"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":6A78
            Key             =   "tVisualStudio2010"
            Object.Tag             =   "2215"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":6E40
            Key             =   "sCodejock"
            Object.Tag             =   "2401"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":7E92
            Key             =   "sOffice2007"
            Object.Tag             =   "2402"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":8EE4
            Key             =   "sOffice2010"
            Object.Tag             =   "2403"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":9F36
            Key             =   "sOrangina"
            Object.Tag             =   "878"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":AF88
            Key             =   "sVista"
            Object.Tag             =   "2404"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":BFDA
            Key             =   "sWinXPLuna"
            Object.Tag             =   "2405"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":D02C
            Key             =   "sWinXPRoyale"
            Object.Tag             =   "2406"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":E07E
            Key             =   "sZune"
            Object.Tag             =   "2407"
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
            Object.Tag             =   "1403"
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
            Object.Tag             =   "1406"
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
            Object.Tag             =   "1104"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":17E32
            Key             =   "threemen"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":18A84
            Key             =   "SysUser"
            Object.Tag             =   "1105"
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
            Object.Tag             =   "1103"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1BBCC
            Key             =   "helpDoc"
            Object.Tag             =   "3102"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1BF1E
            Key             =   "themeSet"
            Object.Tag             =   "2052"
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
            Object.Tag             =   "1108"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1E51E
            Key             =   "SysRole"
            Object.Tag             =   "1106"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1F170
            Key             =   "RoleSelect"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1FDC2
            Key             =   "SysFunc"
            Object.Tag             =   "1107"
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
            Key             =   "ResetSet"
            Object.Tag             =   "2050"
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
   Begin XtremeDockingPane.DockingPane DockingPane1 
      Left            =   6000
      Top             =   3000
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
Dim mcbsPopupIcon As XtremeCommandBars.CommandBar   '托盘图标Pupup菜单
Dim mcbsPopupNavi As XtremeCommandBars.CommandBar   '导航菜单标题行Popup菜单
Dim mcbsPopupTab As XtremeCommandBars.CommandBar    '多标签右键Popup菜单
Dim WithEvents mTabWorkspace As XtremeCommandBars.TabWorkspace '多标签窗口控件
Attribute mTabWorkspace.VB_VarHelpID = -1




Private Sub msAddAction(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建CommandBars的Action
    
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    cbsBars.EnableActions   '启用CommandBars的Actions集合
    
'    cbsActions.Add "Id", "Caption", "TooltipText", "DescriptionText", "Category"   '范例
    With cbsActions
        .Add gID.Sys, "系统", "", "", "系统"
        
        .Add gID.SysAuthChangePassword, "密码修改", "", "", "frmSysAlterPWD"
        .Add gID.SysAuthDepartment, "部门管理", "", "", "frmSysDepartment"
        .Add gID.SysAuthRole, "角色管理", "", "", "frmSysRole"
        .Add gID.SysAuthUser, "用户管理", "", "", "frmSysUser"
        .Add gID.SysAuthFunc, "权限管理", "", "", "frmSysFunc"
        .Add gID.SysAuthLog, "日志管理", "", "", "frmSysLog"
        
        .Add gID.SysFileManage, "文件管理", "", "", "frmSysFile"
        
        .Add gID.SysLoginOut, "退出", "", "", ""
        .Add gID.SysLoginAgain, "重启", "", "", ""
        
        .Add gID.SysExportMain, "导出", "", "", ""
        .Add gID.SysExportToCSV, "导出至CSV", "", "", ""
        .Add gID.SysExportToExcel, "导出至Excel", "", "", ""
        .Add gID.SysExportToHTML, "导出至HTML", "", "", ""
        .Add gID.SysExportToPDF, "导出至PDF", "", "", ""
        .Add gID.SysExportToText, "导出至txt", "", "", ""
        .Add gID.SysExportToWord, "导出至Word", "", "", ""
        .Add gID.SysExportToXML, "导出至XML", "", "", ""
        
        .Add gID.SysPrintMain, "打印", "", "", ""
        .Add gID.SysPrint, "打印", "", "", ""
        .Add gID.SysPrintPageSet, "打印页面设置", "", "", ""
        .Add gID.SysPrintPreview, "打印预览", "", "", ""
        
        .Add gID.SysSearch, "窗口检索", "", "", ""
        .Add gID.SysSearch1Label, "输入窗口名称关键字", "", "", ""
        .Add gID.SysSearch2TextBox, "关键字输入框", "", "", ""
        .Add gID.SysSearch3Button, "检索窗口", "", "", ""
        .Add gID.SysSearch4ListBoxCaption, "检索到的窗口标题列表", "", "", ""
        .Add gID.SysSearch4ListBoxFormID, "检索到的窗体名称列表", "", "", ""
        .Add gID.SysSearch5Go, "跳转至选定窗口", "", "", ""
        
        
        .Add gID.Wnd, "窗口", "", "", "窗口"
        
        .Add gID.WndThemeSkinSet, "窗口主题设置...", "", "", "frmSysThemeSet"
        .Add gID.WndResetLayout, "重置窗口布局", "", "", ""
        .Add gID.WndToolBarCustomize, "自定义工具栏…", "自定义工具栏", "自定义工具栏", ""
        .Add gID.WndToolBarList, "工具栏列表", "工具栏列表", "工具栏列表", ""
        .Add gID.WndOpenListCaption, "已打开窗口列表", "", "", ""
        .Add gID.WndOpenListID, "", "", "", ""
        
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
        
        .Add gID.WndThemeTaskPanel, "任务面板主题", "", "", ""
        .Add gID.WndThemeTaskPanelListView, "ListView", "", "", ""
        .Add gID.WndThemeTaskPanelListViewOffice2003, "ListViewOffice2003", "", "", ""
        .Add gID.WndThemeTaskPanelListViewOfficeXP, "ListViewOfficeXP", "", "", ""
        .Add gID.WndThemeTaskPanelNativeWinXP, "NativeWinXP", "", "", ""
        .Add gID.WndThemeTaskPanelNativeWinXPPlain, "NativeWinXPPlain", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2000, "Office2000", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2000Plain, "Office2000Plain", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2003, "Office2003", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2003Plain, "Office2003Plain", "", "", ""
        .Add gID.WndThemeTaskPanelOfficeXPPlain, "OfficeXPPlain", "", "", ""
        .Add gID.WndThemeTaskPanelResource, "Resource", "", "", ""
        .Add gID.WndThemeTaskPanelShortcutBarOffice2003, "ShortcutBarOffice2003", "", "", ""
        .Add gID.WndThemeTaskPanelToolbox, "Toolbox", "", "", ""
        .Add gID.WndThemeTaskPanelToolboxWhidbey, "ToolboxWhidbey", "", "", ""
        .Add gID.WndThemeTaskPanelVisualStudio2010, "VisualStudio2010", "", "", ""
        
        .Add gID.WndSon, "子窗口控制", "", "", ""
        .Add gID.WndSonCloseAll, "关闭所有子窗口", "", "", ""
        .Add gID.WndSonCloseCurrent, "关闭当前子窗口", "", "", ""
        .Add gID.WndSonCloseLeft, "关闭当前标签左侧子窗口", "", "", ""
        .Add gID.WndSonCloseOther, "关闭其它子窗口", "", "", ""
        .Add gID.WndSonCloseRight, "关闭当前标签右侧子窗口", "", "", ""
        .Add gID.WndSonVbAllBack, "恢复子窗口", "", "", ""
        .Add gID.WndSonVbAllMin, "最小化所有子窗口", "", "", ""
        .Add gID.WndSonVbArrangeIcons, "重新排列最小化图标", "", "", ""
        .Add gID.WndSonVbCascade, "子窗口层叠", "", "", ""
        .Add gID.WndSonVbTileHorizontal, "子窗口水平平铺", "", "", ""
        .Add gID.WndSonVbTileVertical, "子窗口垂直平铺", "", "", ""
        
        .Add gID.WndThemeSkin, "窗口主题", "", "", ""
        .Add gID.WndThemeSkinCodejock, "Codejock", "", "", ""
        .Add gID.WndThemeSkinOffice2007, "Office2007", "", "", ""
        .Add gID.WndThemeSkinOffice2010, "Office2010", "", "", ""
        .Add gID.WndThemeSkinVista, "Vista", "", "", ""
        .Add gID.WndThemeSkinWinXPLuna, "WinXPLuna", "", "", ""
        .Add gID.WndThemeSkinWinXPRoyale, "WinXPRoyale", "", "", ""
        .Add gID.WndThemeSkinZune, "Zune", "", "", ""
        
        
        .Add gID.Tool, "工具", "", "", "工具"
        .Add gID.toolOptions, "选项…", "选项", "选项", "frmOption"
        
        .Add gID.Help, "帮助", "", "", "帮助"
        .Add gID.HelpAbout, "关于…", "", "", ""
        .Add gID.HelpDocument, "帮助文档", "", "", ""
        .Add gID.HelpUpdate, "更新检查", "", "", ""
                
        
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

        .Add gID.Pane, "任务面板", "", "", ""
        .Add gID.PaneNavi, "导航菜单", "", "导航菜单显示/隐藏", ""
        
        .Add gID.PanePopupMenuNavi, "导航菜单标题行Popup菜单", "", "", ""
        .Add gID.PanePopupMenuNaviAutoFoldOther, "自动收拢", "", "点击某子菜单时收拢导航菜单中所有其它主菜单", ""
        .Add gID.PanePopupMenuNaviExpandALL, "全部展开", "", "展开导航菜单中所有主菜单", ""
        .Add gID.PanePopupMenuNaviFoldALL, "全部收拢", "", "收拢导航菜单中所有主菜单", ""
        
        .Add gID.TabWorkspacePopupMenu, "多标签右键菜单", "", "", ""
        
'        .Add gID, "", "", "", ""
        
    End With
    
    '填充cbsActions的其它属性ToolTipText、DescriptionText、Key、Category
    For Each cbsAction In cbsActions
        With cbsAction
            If .ID < 20000 Then
                .ToolTipText = .Caption
                .DescriptionText = .ToolTipText
                .Key = .Category    '为菜单时有特殊用，创建Action时窗体名保存在Category中
                If LCase(Left(.Key, 3)) = "frm" Then
                    Select Case .ID
                        Case gID.toolOptions, gID.SysAuthChangePassword, gID.WndThemeSkinSet
                            '一些不需要受权限控制的窗口
                        Case Else '受控制窗口
                            cbsAction.Enabled = False '先禁需要权限控制的窗口，加载权限时再解锁
                    End Select
                End If
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

Private Sub msAddDockingPane(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建浮动面板
    
    Dim paneNavigation As XtremeDockingPane.Pane
    Dim cbsActions As XtremeCommandBars.CommandBarActions
    
    Set cbsActions = cbsBars.Actions
'    Me.Picture1.Appearance = 0
'    Me.Picture1.BackColor = Me.BackColor
    
    With Me.DockingPane1
        .SetCommandBars cbsBars '若这两种控制同时使用必需这么设置，且CommandBars控件在DockingPane控件的顶层
        With .Options
            .AlphaDockingContext = True
            .ShowDockingContextStickers = True
            .StickerStyle = StickerStyleVisualStudio2008 '必须使AlphaDockingContext、ShowDockingContextStickers都为True
        End With
        Set paneNavigation = .CreatePane(gID.PaneNavi, 260, 240, DockLeftOf)
        cbsBars.Actions(gID.PaneNavi).Checked = True
    End With
    With paneNavigation
        .Title = cbsActions(gID.PaneNavi).Caption
        .TitleToolTip = .Title
        .TabCaption = .Title
        .Options = PaneHasMenuButton
        .Handle = Me.Picture1.hwnd
    End With
    
    Set paneNavigation = Nothing
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
    Dim cbsMenuCtrlTemp As XtremeCommandBars.CommandBarControl
    
    Set cbsMenuBar = cbsBars.ActiveMenuBar
    cbsMenuBar.ShowGripper = False  '不显示可拖动的那个点点标记
    cbsMenuBar.EnableDocking xtpFlagStretched     '菜单栏独占一行且不能主动拖动
    
    '系统主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Sys, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.SysAuthChangePassword, ""
                        
        Set cbsMenuCtrlTemp = .Add(xtpControlButton, gID.SysFileManage, "")
        cbsMenuCtrlTemp.BeginGroup = True
          
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysAuthDepartment, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysAuthRole, ""
        .Add xtpControlButton, gID.SysAuthUser, ""
        .Add xtpControlButton, gID.SysAuthFunc, ""
        .Add xtpControlButton, gID.SysAuthLog, ""
                      
        Set cbsMenuCtrlTemp = .Add(xtpControlButtonPopup, gID.SysExportMain, "导出")
        cbsMenuCtrlTemp.BeginGroup = True
        With cbsMenuCtrlTemp.CommandBar.Controls
            Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysExportToCSV, "")
            cbsMenuCtrl.BeginGroup = True
            For mlngID = gID.SysExportToExcel To gID.SysExportToWord
                .Add xtpControlButton, mlngID, ""
            Next
        End With
        
        Set cbsMenuCtrlTemp = .Add(xtpControlButtonPopup, gID.SysPrintMain, "打印")
        With cbsMenuCtrlTemp.CommandBar.Controls
            Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysPrintPageSet, "")
            cbsMenuCtrl.BeginGroup = True
            .Add xtpControlButton, gID.SysPrintPreview, ""
            .Add xtpControlButton, gID.SysPrint, ""
        End With
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysLoginAgain, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysLoginOut, ""
        
    End With
    
    '窗口主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.WndThemeSkinSet, "" '皮肤设置
        .Add xtpControlButton, gID.WndResetLayout, "" '重置布局
        .Add xtpControlButton, gID.PaneNavi, ""  '导航菜单显示/隐藏
        
        '特殊ID XTP_ID_CUSTOMIZE=35001自定义工具栏
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.WndToolBarCustomize, "")
        cbsMenuCtrl.BeginGroup = True
    
        '特殊ID XTP_ID_TOOLBARLIST=59392工具栏列表
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndToolBarList, "")
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
        
        'CommandBars工具栏主题子菜单
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeCommandBars, "")
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
                .Add xtpControlButton, mlngID, ""
            Next
        End With
        
        'TaskPanel任务面板主题
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeTaskPanel, "")
        cbsMenuCtrl.BeginGroup = True
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
                .Add xtpControlButton, mlngID, ""
            Next
        End With
        
        '子窗口控制
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndSon, "")
        cbsMenuCtrl.BeginGroup = True
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndSonCloseAll To gID.WndSonVbTileVertical
                .Add xtpControlButton, mlngID, ""
            Next
            .Find(, gID.WndSonVbAllBack).BeginGroup = True
        End With
                
        '子窗口列表
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndOpenListCaption, "")
        cbsMenuCtrl.BeginGroup = True
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, gID.WndOpenListID, ""
    End With
    
    '工具菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Tool, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.toolOptions, ""
    
    '帮助主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Help, "")
    With cbsMenuMain.CommandBar.Controls
        For mlngID = gID.HelpAbout To gID.HelpUpdate
            .Add xtpControlButton, mlngID, ""
        Next
    End With
        
    Set cbsMenuBar = Nothing
    Set cbsMenuMain = Nothing
    Set cbsMenuCtrl = Nothing
    Set cbsMenuCtrlTemp = Nothing
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
    
    '创建导航菜单面板上标题行上的弹出式菜单
    Set mcbsPopupNavi = cbsBars.Add(cbsBars.Actions(gID.PanePopupMenuNavi).Caption, xtpBarPopup)
    With mcbsPopupNavi.Controls
        .Add xtpControlButton, gID.PanePopupMenuNaviAutoFoldOther, ""
        .Add xtpControlButton, gID.PanePopupMenuNaviExpandALL, ""
        .Add xtpControlButton, gID.PanePopupMenuNaviFoldALL, ""
    End With
    
    '创建多标签上的右键菜单
    Set mcbsPopupTab = cbsBars.Add(cbsBars.Actions(gID.TabWorkspacePopupMenu).Caption, xtpBarPopup)
    With mcbsPopupTab.Controls
        For mlngID = gID.WndSonCloseAll To gID.WndSonVbTileVertical
            .Add xtpControlButton, mlngID, ""
        Next
        .Find(, gID.WndSonVbAllBack).BeginGroup = True
    End With
        
End Sub

Private Sub msAddTaskPanelItem(ByRef tskPanel As XtremeTaskPanel.TaskPanel)
    '创建导航菜单
    
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    Dim taskItem As XtremeTaskPanel.TaskPanelGroupItem
    Dim cbsActions As XtremeCommandBars.CommandBarActions
    Dim lngID As Long, lngLeftMargins As Long, L As Long, T As Long, R As Long, b As Long
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim imgIcon As MSComctlLib.ListImage
    
    Set cbsActions = Me.CommandBars1.Actions
    
    '创建系统菜单
    Set taskGroup = tskPanel.Groups.Add(gID.Sys, cbsActions(gID.Sys).Caption)
    With taskGroup.Items
        Set taskItem = .Add(gID.SysAuthChangePassword, cbsActions(gID.SysAuthChangePassword).Caption, xtpTaskItemTypeLink)
        taskItem.GetRect L, T, R, b '为了排列好看，每一级子菜单使用同样的缩进距离,主要是为了获取L值(左边距)
        lngLeftMargins = L
        .Add gID.SysFileManage, cbsActions(gID.SysFileManage).Caption, xtpTaskItemTypeLink
        .Add gID.SysAuthDepartment, cbsActions(gID.SysAuthDepartment).Caption, xtpTaskItemTypeLink
        .Add gID.SysAuthRole, cbsActions(gID.SysAuthRole).Caption, xtpTaskItemTypeLink
        .Add gID.SysAuthUser, cbsActions(gID.SysAuthUser).Caption, xtpTaskItemTypeLink
        .Add gID.SysAuthFunc, cbsActions(gID.SysAuthFunc).Caption, xtpTaskItemTypeLink
        .Add gID.SysAuthLog, cbsActions(gID.SysAuthLog).Caption, xtpTaskItemTypeLink
        
        Set taskItem = .Add(gID.SysExportMain, cbsActions(gID.SysExportMain).Caption, xtpTaskItemTypeText)
        taskItem.Bold = True
        taskItem.SetMargins lngLeftMargins, 0, 0, 0
        For lngID = gID.SysExportToCSV To gID.SysExportToWord
            Set taskItem = .Add(lngID, cbsActions(lngID).Caption, xtpTaskItemTypeLink)
            taskItem.SetMargins lngLeftMargins, 0, 0, 0
        Next
        
        Set taskItem = .Add(gID.SysPrintMain, cbsActions(gID.SysPrintMain).Caption, xtpTaskItemTypeText)
        taskItem.Bold = True
        taskItem.SetMargins lngLeftMargins, 0, 0, 0
        For lngID = gID.SysPrintPageSet To gID.SysPrint
            Set taskItem = .Add(lngID, cbsActions(lngID).Caption, xtpTaskItemTypeLink, lngID)
            taskItem.SetMargins lngLeftMargins, 0, 0, 0
        Next
        
        .Add gID.SysLoginAgain, cbsActions(gID.SysLoginAgain).Caption, xtpTaskItemTypeLink
        .Add gID.SysLoginOut, cbsActions(gID.SysLoginOut).Caption, xtpTaskItemTypeLink
    End With
    
    '创建窗口菜单
    Set taskGroup = tskPanel.Groups.Add(gID.Wnd, cbsActions(gID.Wnd).Caption)
    With taskGroup.Items
        .Add gID.WndThemeSkinSet, cbsActions(gID.WndThemeSkinSet).Caption, xtpTaskItemTypeLink
        .Add gID.WndResetLayout, cbsActions(gID.WndResetLayout).Caption, xtpTaskItemTypeLink
        
        .Add gID.WndToolBarCustomize, cbsActions(gID.WndToolBarCustomize).Caption, xtpTaskItemTypeLink
        
        Set taskItem = .Add(gID.WndThemeCommandBars, cbsActions(gID.WndThemeCommandBars).Caption, xtpTaskItemTypeText)
        taskItem.Bold = True
        taskItem.SetMargins lngLeftMargins, 0, 0, 0
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Add mlngID, cbsActions(mlngID).Caption, xtpTaskItemTypeLink
            .Find(mlngID).SetMargins lngLeftMargins, 0, 0, 0
        Next
        
        Set taskItem = .Add(gID.WndThemeTaskPanel, cbsActions(gID.WndThemeTaskPanel).Caption, xtpTaskItemTypeText)
        taskItem.Bold = True
        taskItem.SetMargins lngLeftMargins, 0, 0, 0
        For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
            .Add mlngID, cbsActions(mlngID).Caption, xtpTaskItemTypeLink
            .Find(mlngID).SetMargins lngLeftMargins, 0, 0, 0
        Next
        
        Set taskItem = .Add(gID.WndSon, cbsActions(gID.WndSon).Caption, xtpTaskItemTypeText)
        taskItem.Bold = True
        taskItem.SetMargins lngLeftMargins, 0, 0, 0
        For mlngID = gID.WndSonCloseAll To gID.WndSonVbTileVertical
            .Add mlngID, cbsActions(mlngID).Caption, xtpTaskItemTypeLink
            .Find(mlngID).SetMargins lngLeftMargins, 0, 0, 0
        Next
    End With
    
    '创建工具菜单
    Set taskGroup = tskPanel.Groups.Add(gID.Tool, cbsActions(gID.Tool).Caption)
    With taskGroup.Group.Items
        .Add gID.toolOptions, cbsActions(gID.toolOptions).Caption, xtpTaskItemTypeLink
    End With
    
    
    '创建帮助菜单
    Set taskGroup = tskPanel.Groups.Add(gID.Help, cbsActions(gID.Help).Caption)
    With taskGroup.Group.Items
        For mlngID = gID.HelpAbout To gID.HelpUpdate
            .Add mlngID, cbsActions(mlngID).Caption, xtpTaskItemTypeLink
        Next
    End With
    
    
    '添加GroupItem图标
    tskPanel.SetImageList Me.ImageList1 '绑定图标集合
    For Each taskGroup In tskPanel.Groups
        For Each taskItem In taskGroup.Items
            For Each imgIcon In Me.ImageList1.ListImages
                If Val(imgIcon.Tag) = taskItem.ID Then '先预告在ImageList1控件中设置每个图标的Tag值为GroupItem的ID值
                    taskItem.IconIndex = imgIcon.Index
                    Exit For
                End If
            Next
        Next
    Next
    
    '同步权限
    For Each cbsAction In cbsActions
        If Not tskPanel.Find(cbsAction.ID) Is Nothing Then '并不是每个Action对应一个GroupItem
            tskPanel.Find(cbsAction.ID).Enabled = cbsAction.Enabled
        End If
    Next
    
    '加载折叠菜单状态
    For Each taskGroup In tskPanel.Groups
        taskGroup.Expanded = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, "TP" & CStr(taskGroup.ID), 0))
    Next
    
    Set taskItem = Nothing
    Set taskGroup = Nothing
    Set cbsAction = Nothing
    Set cbsActions = Nothing
    Set imgIcon = Nothing
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
    
    '导航菜单主题
    Set cbsBar = cbsBars.Add(cbsActions(gID.WndThemeTaskPanel).Caption, xtpBarTop)
    With cbsBar.Controls
        For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '窗口检索工具栏
    Set cbsBar = cbsBars.Add(cbsActions(gID.SysSearch).Caption, xtpBarTop)
    With cbsBar.Controls
        .Add xtpControlLabel, gID.SysSearch1Label, ""
        Set cbsCtr = .Add(xtpControlEdit, gID.SysSearch2TextBox, "")
        cbsCtr.Width = 200
        cbsCtr.EditHint = "输入窗口标题关键字"
        .Add xtpControlButton, gID.SysSearch3Button, ""
        Set cbsCtr = .Add(xtpControlComboBox, gID.SysSearch4ListBoxCaption, "")
        cbsCtr.Width = 200
        cbsCtr.EditHint = "从列表中选择一个窗口标题"
        Set cbsCtr = .Add(xtpControlComboBox, gID.SysSearch4ListBoxFormID, "")
        cbsCtr.Visible = False
        .Add xtpControlButton, gID.SysSearch5Go, ""
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
        .SetPaneText gID.StatusBarPaneUserInfo, gVar.UserFullName
        .FindPane(gID.StatusBarPaneUserInfo).Width = 80
        
        .AddPane gID.StatusBarPaneIP
        .SetPaneText gID.StatusBarPaneIP, gVar.UserLoginIP ' Me.Winsock1.Item(1).LocalIP  'gVar.TCPSetIP
        .FindPane(gID.StatusBarPaneIP).Width = 90
        
        .AddPane gID.StatusBarPanePort
        .SetPaneText gID.StatusBarPanePort, gVar.TCPSetPort
        .FindPane(gID.StatusBarPanePort).Width = 60
        
        .AddPane gID.StatusBarPaneConnectState
        .SetPaneText gID.StatusBarPaneConnectState, gVar.ClientStateDisConnected
        .FindPane(gID.StatusBarPaneConnectState).Width = 60
                
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
            Case .WndThemeTaskPanelListView To .WndThemeTaskPanelVisualStudio2010
                Call msThemeTaskPanel(CID, cbsBars)
            Case .WndSonCloseAll To .WndSonVbTileVertical
                Call msWindowControl(CID)
            Case .PanePopupMenuNaviAutoFoldOther To .PanePopupMenuNaviFoldALL
                Call msPopupMenuNavi(CID, cbsBars)
            Case .SysSearch5Go, .SysSearch4ListBoxCaption
                If Len(cbsBars.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxCaption).Text) > 0 Then
                    Call msLeftClick(CLng(cbsBars.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxFormID).List(cbsBars.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxCaption).ListIndex)), cbsBars)
                End If
            Case .SysSearch3Button
                Call msSearchWindow(CID, cbsBars)
            Case .WndResetLayout
                Call msResetLayout(cbsBars)
            Case .PaneNavi
                Me.DockingPane1.FindPane(CID).Closed = Not Me.DockingPane1.FindPane(CID).Closed
                cbsActions.Action(CID).Checked = Not Me.DockingPane1.FindPane(CID).Closed
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
            
            Case .HelpUpdate
                If gfAppExist(gVar.EXENameOfUpdate) Then
                    MsgBox "更新程序在进程中已存在，请勿重复打开！", vbInformation, "已打开提示"
                Else
                    If MsgBox("更新程序打开期间需要占用一个用户数，确定现在进行更新检查吗？", vbQuestion + vbOKCancel, "打开更新程序询问") = vbOK Then
                        Dim strUP As String
                        strUP = gVar.AppPath & gVar.EXENameOfUpdate & " " & gVar.EXENameOfClient '显示打开更新程序命令行
                        Call msOpenUpdate(strUP) '用命令行打开更新程序
                    End If
                End If
            
            Case .SysExportToCSV To .SysExportToWord, .SysPrintPageSet To .SysPrint
                If Me.ActiveForm Is Nothing Then Exit Sub
                If Me.ActiveForm.ActiveControl Is Nothing Then Exit Sub
                If Not (TypeOf Me.ActiveForm.ActiveControl Is FlexCell.Grid) Then Exit Sub
                If Not cbsActions(CID).Enabled Then Exit Sub
                Select Case CID
                    Case .SysExportToCSV To .SysExportToXML
                        Call gsGridExportTo(Me.ActiveForm.ActiveControl, CID)
                    Case .SysExportToText
                        If MsgBox("是否将当前表格内容导出至txt文本文档？", vbQuestion + vbYesNo, "询问") = vbYes Then Call gsGridToText(Me.ActiveForm.ActiveControl)
                    Case .SysExportToWord
                        If MsgBox("是否将当前表格内容导出至Word文档？", vbQuestion + vbYesNo, "询问") = vbYes Then Call gsGridToWord(Me.ActiveForm.ActiveControl)
                        
                    Case .SysPrint
                        If MsgBox("确定打印当前表格内容吗？", vbQuestion + vbOKCancel, "打印询问") = vbOK Then Call gsGridPrint(Me.ActiveForm.ActiveControl)
                    Case .SysPrintPreview
                        Call gsGridPrintPreview(Me.ActiveForm.ActiveControl)
                    Case .SysPrintPageSet
                        Call gsGridPageSet(Me.ActiveForm.ActiveControl)
                End Select
                
            Case Else
                strKey = LCase(cbsActions.Action(CID).Key)
                If Left(strKey, 3) = "frm" Then
                    If cbsActions.Action(CID).Enabled Then
                        Select Case CID
                            Case .toolOptions, .SysAuthChangePassword, .WndThemeSkinSet
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
    
    Rem On Error Resume Next    '加/解密函数过程可能有异常
    With gVar
        .ParaBlnWindowCloseMin = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, 1))    '关闭时最小化
        .ParaBlnWindowMinHide = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, 1))  '最小化时隐藏
        .ParaBlnWindowStartMinC = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowStartMinC, 1)) '启动时最小化
        
        .TCPDefaultIP = Me.Winsock1.Item(1).LocalIP '本机IP地址
        .TCPSetIP = gfCheckIP(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPDefaultIP)) '要连接的服务端IP地址
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , .TCPDefaultPort, 10000, 65535) '要连接的服务器端口
        
        .ParaBlnAutoStartupAtBoot = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, 0))  '开机自动启动
        .ParaBlnUserAutoLogin = Val(GetSetting(.RegAppName, .RegSectionUser, .RegKeyParaUserAutoLogin, 0)) '自动登陆
        .ParaBlnRememberUserList = Val(GetSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserList, 0)) '记住用户名
        .ParaBlnRememberUserPassword = Val(GetSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserPassword, 0)) '记住密码
        
        .UserLoginIP = .TCPDefaultIP '本机IP赋值给另一个变量
        .UserComputerName = gfBackComputerInfo(ciComputerName) '获取计算机名
        
'''        '由服务端发过来给客户端
'''        .ConSource = gfCheckIP(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .TCPSetIP))   '服务器名称/IP
'''        .ConDatabase = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString("dbTest", .EncryptKey)), .EncryptKey)    '数据库名
'''        .ConUserID = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString("123", .EncryptKey)), .EncryptKey)  '登陆名
'''        .ConPassword = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString("888888", .EncryptKey)), .EncryptKey)    '登陆密码
        
        
    End With
End Sub

Private Sub msLoadUserAuthority(ByVal strUID As String)
    '权限控制
    
    Const strFRM As String = "frm"
    Dim cbsAction As CommandBarAction
    Dim strSQL As String, strKey As String, strSys As String
    
    strUID = Trim(strUID)
    If Len(strUID) = 0 Then Exit Sub
    
    strSys = LCase(gVar.UserLoginName)
    If strSys = LCase(gVar.AccountAdmin) Or strSys = LCase(gVar.AccountSystem) Then   '程序内定两个用户拥有所有权限
        For Each cbsAction In gWind.CommandBars1.Actions
            cbsAction.Enabled = True
            If Not Me.TaskPanel1.Find(cbsAction.ID) Is Nothing Then Me.TaskPanel1.Find(cbsAction.ID).Enabled = True
        Next
        Exit Sub
    End If
    
    strSQL = "SELECT DISTINCT t1.UserAutoID ,t1.UserLoginName ,t1.UserFullName " & _
             ",t5.FuncAutoID ,t5.FuncCaption ,t5.FuncName ,t5.FuncType " & _
             ",t6.FuncName AS [FuncFormName] FROM tb_FT_Sys_User AS [t1] " & _
             "INNER JOIN tb_FT_Sys_UserRole AS [t2] ON t1.UserAutoID =t2.UserAutoID " & _
             "INNER JOIN tb_FT_Sys_RoleFunc AS [t4] ON t2.RoleAutoID =t4.RoleAutoID " & _
             "INNER JOIN tb_FT_Sys_Func AS [t5] ON t4.FuncAutoID =t5.FuncAutoID " & _
             "INNER JOIN tb_FT_Sys_Func AS [t6] ON t5.FuncParentID =t6.FuncAutoID " & _
             "WHERE t1.UserAutoID =" & strUID
    Set gVar.rsURF = gfBackRecordset(strSQL)
    With gVar.rsURF
        If .State = adStateOpen Then
            If .RecordCount > 0 Then
                For Each cbsAction In Me.CommandBars1.Actions
                    strKey = LCase(cbsAction.Key)
                    If Len(strKey) > 0 Then
                        If Left(strKey, 3) = strFRM Then
                            .MoveFirst
                            Do While Not .EOF
                                If LCase(.Fields("FuncName")) = strKey Then
                                    cbsAction.Enabled = True
                                    Me.TaskPanel1.Find(cbsAction.ID).Enabled = True
                                End If
                                .MoveNext
                            Loop
                        End If
                    End If
                Next
            End If
        End If
    End With
    
    Set cbsAction = Nothing
    
End Sub

Public Sub msOpenUpdate(ByVal strCmd As String)
    '用命令行打开更新程序
    
    If Not gfShell(strCmd) Then
        Call gsAlarmAndLogEx("更新程序启动异常", "更新检查失败", True, vbCritical)
    End If
End Sub

Private Sub msPopupMenuNavi(ByVal PID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    '导航菜单上弹出菜单的响应
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    
    Select Case PID
        Case gID.PanePopupMenuNaviAutoFoldOther
            cbsBars.Actions(PID).Checked = Not cbsBars.Actions(PID).Checked
        Case gID.PanePopupMenuNaviExpandALL
            For Each taskGroup In Me.TaskPanel1.Groups
                taskGroup.Expanded = True
            Next
        Case gID.PanePopupMenuNaviFoldALL
            For Each taskGroup In Me.TaskPanel1.Groups
                taskGroup.Expanded = False
            Next
    End Select
    
    Set taskGroup = Nothing
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

Private Sub msSearchWindow(ByVal WID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    '窗口标题检索
    
    
    Dim strName As String
    Dim cbsAction As CommandBarAction
    Dim cbsCtrlCaption As CommandBarComboBox
    Dim cbsCtrlFormID As CommandBarComboBox
    Dim blnClear As Boolean
    
    strName = LCase(Trim(cbsBars.FindControl(xtpControlEdit, gID.SysSearch2TextBox).Text))
    If Len(strName) = 0 Then Exit Sub
    
    Set cbsCtrlCaption = cbsBars.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxCaption)
    Set cbsCtrlFormID = cbsBars.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxFormID)
    
    For Each cbsAction In cbsBars.Actions
        If cbsAction.ID < 20000 Then     '所有窗口的ID小于2000
            If LCase(Left(cbsAction.Key, 3)) = "frm" Then   '窗口的Name属性以frm开头
                If InStr(LCase(cbsAction.Caption), strName) > 0 Then
                    If Not blnClear Then
                        cbsCtrlCaption.Clear
                        cbsCtrlFormID.Clear
                        blnClear = True
                    End If
                    cbsCtrlCaption.AddItem cbsAction.Caption
                    cbsCtrlFormID.AddItem cbsAction.ID
                End If
            End If
        End If
    Next
    
    If blnClear Then
        If cbsCtrlCaption.ListCount > 0 Then cbsCtrlCaption.ListIndex = 1
    Else
        cbsCtrlCaption.ListIndex = 0
    End If
    
    Set cbsAction = Nothing
    Set cbsCtrlCaption = Nothing
    Set cbsCtrlFormID = Nothing
    
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

Private Sub msThemeTaskPanel(ByVal LID As Long, ByRef cbsSet As XtremeCommandBars.CommandBars)
    '任务面板主题设置
    Dim lngTheme As XtremeTaskPanel.XTPTaskPanelVisualTheme
    
    Select Case LID
        Case gID.WndThemeTaskPanelListView
            lngTheme = xtpTaskPanelThemeListView
        Case gID.WndThemeTaskPanelListViewOffice2003
            lngTheme = xtpTaskPanelThemeListViewOffice2003
        Case gID.WndThemeTaskPanelListViewOfficeXP
            lngTheme = xtpTaskPanelThemeListViewOfficeXP
        Case gID.WndThemeTaskPanelNativeWinXP
            lngTheme = xtpTaskPanelThemeNativeWinXP
        Case gID.WndThemeTaskPanelNativeWinXPPlain
            lngTheme = xtpTaskPanelThemeNativeWinXPPlain
        Case gID.WndThemeTaskPanelOffice2000
            lngTheme = xtpTaskPanelThemeOffice2000
        Case gID.WndThemeTaskPanelOffice2000Plain
            lngTheme = xtpTaskPanelThemeOffice2000Plain
        Case gID.WndThemeTaskPanelOffice2003
            lngTheme = xtpTaskPanelThemeOffice2003
        Case gID.WndThemeTaskPanelOffice2003Plain
            lngTheme = xtpTaskPanelThemeOffice2003Plain
        Case gID.WndThemeTaskPanelOfficeXPPlain
            lngTheme = xtpTaskPanelThemeOfficeXPPlain
        Case gID.WndThemeTaskPanelResource
            lngTheme = xtpTaskPanelThemeResource
        Case gID.WndThemeTaskPanelShortcutBarOffice2003
            lngTheme = xtpTaskPanelThemeShortcutBarOffice2003
        Case gID.WndThemeTaskPanelToolbox
            lngTheme = xtpTaskPanelThemeToolbox
        Case gID.WndThemeTaskPanelToolboxWhidbey
            lngTheme = xtpTaskPanelThemeToolboxWhidbey
        Case Else
            lngTheme = xtpTaskPanelThemeVisualStudio2010
            LID = gID.WndThemeTaskPanelVisualStudio2010
    End Select
    
    Me.TaskPanel1.VisualTheme = lngTheme
    For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        cbsSet.Actions(mlngID).Checked = False
    Next
    cbsSet.Actions(LID).Checked = True
End Sub

Private Sub msUnloadMe(Optional ByVal blnUnload As Boolean = True)
    '卸载窗体
    If Not blnUnload Then Exit Sub
    gVar.CloseWindow = True
    If Forms.Count > 1 Then '先关闭其它窗体，再关闭主窗体
        Dim frmUld As Form
        
        For Each frmUld In Forms
            If frmUld.Name <> Me.Name Then Unload frmUld
        Next
        Set frmUld = Nothing
    End If
    Unload Me
End Sub

Private Sub msWindowControl(ByVal WID As Long)
    '子窗口控制
    
    Dim frmTag As Form
    Dim C As Long
    Dim itemCur As XtremeCommandBars.TabControlItem
        
    With gID
        Select Case WID
            Case .WndSonCloseAll    '关闭所有窗口
                For Each frmTag In Forms
                    If frmTag.Name <> gWind.Name Then Unload frmTag
                Next
            Case .WndSonCloseCurrent    '关闭当前窗口
                If Not ActiveForm Is Nothing Then Unload ActiveForm
            Case .WndSonCloseLeft   '关闭左侧窗口
                If Forms.Count > 2 Then
                    Set itemCur = mTabWorkspace.Selected
                    itemCur.Tag = "c"   '标记当前窗口，因为Index值在窗口数量变化时会变化，不能作为唯一判断依据
                    For C = 0 To mTabWorkspace.ItemCount - 1
                        If mTabWorkspace.Item(0).Tag = itemCur.Tag Then
                            itemCur.Tag = ""    '记得清空。Tag属性默认值就是空字符串
                            Exit For
                        Else
                            mTabWorkspace.Item(0).Selected = True   '激活要删除的窗口
                            Unload ActiveForm
                        End If
                    Next
                End If
            Case .WndSonCloseOther  '关闭其它窗口
                If Forms.Count > 1 Then
                    For Each frmTag In Forms
                        If frmTag.Name <> gWind.Name Then
                            If Not (frmTag.Name = ActiveForm.Name And frmTag.Caption = ActiveForm.Caption) Then
                                Unload frmTag
                            End If
                        End If
                    Next
                End If
            Case .WndSonCloseRight  '关闭右侧窗口
                If Forms.Count > 2 Then
                    Set itemCur = mTabWorkspace.Selected
                    itemCur.Tag = "c"
                    For C = mTabWorkspace.ItemCount - 1 To 0 Step -1
                        If mTabWorkspace.Item(C).Tag = itemCur.Tag Then
                            itemCur.Tag = ""
                            Exit For
                        Else
                            mTabWorkspace.Item(C).Selected = True
                            Unload ActiveForm
                        End If
                    Next
                End If
            Case .WndSonVbAllBack
                For Each frmTag In Forms
                    If frmTag.Name <> gWind.Name Then frmTag.WindowState = vbNormal
                Next
            Case .WndSonVbAllMin
                For Each frmTag In Forms
                    If frmTag.Name <> gWind.Name Then frmTag.WindowState = vbMinimized
                Next
            Case .WndSonVbCascade
                Me.Arrange vbCascade
            Case .WndSonVbArrangeIcons
                Me.Arrange vbArrangeIcons
            Case .WndSonVbTileHorizontal
                Me.Arrange vbTileHorizontal
            Case .WndSonVbTileVertical
                Me.Arrange vbTileVertical
        End Select
    End With
    
    Set frmTag = Nothing
    Set itemCur = Nothing
End Sub


Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '命令单击事件
    Call msLeftClick(Control.ID, Me.CommandBars1)
End Sub


Private Sub CommandBars1_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'CommandBars控件的Action状态的切换
    
    Dim blnFC As Boolean    '判断是否为FC表格
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    Dim blnMainWindow As Boolean '判断主窗体是否已全部加载完成
    Dim tskItems As XtremeTaskPanel.TaskPanelGroupItems   '导航菜单集合
    
    Set cbsActions = Me.CommandBars1.Actions
    Set tskItems = Me.TaskPanel1.Groups.Find(gID.Sys).Items
    
    If Not Me.ActiveForm Is Nothing Then
        If Not Me.ActiveForm.ActiveControl Is Nothing Then
            blnFC = TypeOf Me.ActiveForm.ActiveControl Is FlexCell.Grid    '当前活动控件是FC表格
        End If
    End If
    
    blnMainWindow = gVar.ShowMainWindow
    
    With gID
        For mlngID = .SysExportToCSV To .SysExportToWord
            cbsActions(mlngID).Enabled = blnFC  '活动控件是FC表格则激活对应Action，否则使其不可用
            tskItems.Find(mlngID).Enabled = blnFC
        Next
        For mlngID = .SysPrintPageSet To .SysPrint
            cbsActions(mlngID).Enabled = blnFC
            tskItems.Find(mlngID).Enabled = blnFC
        Next
        For mlngID = .IconPopupMenuMaxWindow To .IconPopupMenuShowWindow
            cbsActions(mlngID).Enabled = blnMainWindow  '主窗体未加载完成之前，托盘图标菜单某些不可用，
        Next
    End With
    
    Set cbsActions = Nothing
    Set tskItems = Nothing
End Sub


Private Sub DockingPane1_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    '导航菜单任务面板显示与隐藏（关闭）
    If Action = PaneActionClosed Then
        If Pane.ID = gID.PaneNavi Then
            Me.CommandBars1.Actions(Pane.ID).Checked = False
        End If
    End If
End Sub

Private Sub DockingPane1_PanePopupMenu(ByVal Pane As XtremeDockingPane.IPane, ByVal X As Long, ByVal Y As Long, Handled As Boolean)
    '导航菜单任务面板上标题行上弹出式菜单
    If Pane.ID = gID.PaneNavi Then
        mcbsPopupNavi.ShowPopup , X * 15, Y * 15
    End If
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
    Call msAddDockingPane(cbsBars) '创建可拖曳的浮动面板
    Call msAddTaskPanelItem(Me.TaskPanel1)  '创建导航菜单
    
    cbsBars.AddImageList ImageList1         '使CommandBars控件匹配ImageList控件中图标
    cbsBars.EnableCustomization True        '允许CommandBars自定义，此属性最好放在所有CommandBars设定之后
    cbsBars.Options.UpdatePeriod = 250      '更改CommandBars的Update事件的执行周期，默认100ms
    
    Set mTabWorkspace = cbsBars.ShowTabWorkspace(True) '起用窗口多标签模式
    mTabWorkspace.Flags = xtpWorkspaceShowActiveFiles Or xtpWorkspaceShowCloseSelectedTab '显示活动窗口列表、当前窗口显示关闭按钮
    
    
    '加载工具栏主题
    Call gsThemeCommandBar(Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientCommandbarsTheme, gID.WndThemeCommandBarsRibbon)), cbsBars)
    
    '注册表信息加载-CommandBars设置
    Call cbsBars.LoadCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    
'    '加载面板样式。暂未知原因，加载后连添加的TaskPanel内容都抹掉了
'    Call Me.DockingPane1.LoadState(gVar.RegKeyDockingPane, gVar.RegAppName, gVar.RegKeyDockPaneClientSetting)
    
    '加载导航菜单的设置
    Call msThemeTaskPanel(gfGetRegNumericValue(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelTheme, _
        True, gID.WndThemeTaskPanelNativeWinXP, gID.WndThemeTaskPanelListView, gID.WndThemeTaskPanelVisualStudio2010), cbsBars)
    cbsBars.Actions(gID.PanePopupMenuNaviAutoFoldOther).Checked = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelAutoFold, 1))
    '菜单的折叠放在msAddTaskPanelItem方法中
    
    Call gsLoadSkin(Me, Me.SkinFramework1, sNone, True)  '加载窗口主题
    
    Set cbsBars = Nothing   '销毁使用完的对象
    
    Call gsFormSizeLoad(Me, False) '注册表信息加载-窗口位置大小
    
    '更新检查
    If gfAppExist(gVar.EXENameOfUpdate) Then
        Me.Timer1.Item(1).Enabled = False
        MsgBox "更新程序会占用一个用户数！" & vbCrLf & "请先结束已存在的更新程序进程后再打开软件。", vbExclamation, "更新已打开提醒"
        Call msUnloadMe(True)
        End
    Else
        strUpdate = gVar.AppPath & gVar.EXENameOfUpdate & " " & gVar.EXENameOfClient & _
                gVar.CmdLineSeparator & gVar.CmdLineParaOfHide      '生成隐式打开更新检测程序的命令行
        Call msOpenUpdate(strUpdate) '用命令行隐式打开更新程序
    End If
    
    '检查是否为试用版*******************************
    '==============================================
    
    
    '==============================================
    
    Call gfNotifyIconAdd(Me)    '添加托盘图标
    Me.Hide '未登陆，不显示主窗体
    
    If LCase(App.EXEName & ".exe") <> LCase(gVar.EXENameOfClient) Then
        Call gsAlarmAndLogEx("不可擅自修改可执行的应用程序文件名！", "严重警报", True, vbCritical)
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
    Dim gVarClear As gtypeCommonVariant
    Dim lngValue As Long
    Dim cbsActions As XtremeCommandBars.CommandBarActions
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    
    Set cbsActions = Me.CommandBars1.Actions
    
    '保存注册表信息-CommandBars设置
    Call Me.CommandBars1.SaveCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    lngValue = gID.WndThemeCommandBarsVS2008 '工具栏主题
    For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        If cbsActions(mlngID).Checked Then
            lngValue = mlngID
            Exit For
        End If
    Next
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientCommandbarsTheme, lngValue)
        
    '保存浮动面板样式
    Call Me.DockingPane1.SaveState(gVar.RegKeyDockingPane, gVar.RegAppName, gVar.RegKeyDockPaneClientSetting)
    
    '保存导航菜单样式
    lngValue = IIf(cbsActions(gID.PanePopupMenuNaviAutoFoldOther).Checked, 1, 0) '自动折叠
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelAutoFold, lngValue)
    lngValue = gID.WndThemeTaskPanelNativeWinXP 'TaskPanel主题
    For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        If cbsActions(mlngID).Checked Then
            lngValue = mlngID
            Exit For
        End If
    Next
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelTheme, lngValue)
    
    For Each taskGroup In Me.TaskPanel1.Groups '菜单折叠状态
        lngValue = IIf(taskGroup.Expanded, 1, 0)
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, "TP" & CStr(taskGroup.ID), lngValue)
    Next
    
    
    Call gsFormSizeSave(Me, False) '保存注册表信息-窗口位置大小
    Call gsSaveCommandbarsTheme(Me.CommandBars1, False)   '保存CommandBars的风格主题
    
    If gfAppExist(gVar.EXENameOfUpdate) Then '如果打开了一个更新程序则关闭
        If Not gfCloseApp(gVar.EXENameOfUpdate) Then
            Call gsAlarmAndLogEx("软件退出时无法同时关闭更新程序", "关闭更新程序异常", False)
        End If
    End If
    
    gArr(1) = gArr(0) '清空文件传输数组中的信息
    gVar = gVarClear '清除gVar公用变量
    
    Call SkinFramework1.LoadSkin("", "")    '清空皮肤
    Set mXtrStatusBar = Nothing  '清除状态栏
    Set mcbsPopupIcon = Nothing '清除Popup菜单
    Call gfNotifyIconDelete(Me) '删除托盘图标
    gNotifyIconData = resetNotifyIconData   '清空托盘气泡信息。否则重启程序时会自动弹出？而且只能放上句删除托盘图标语句的后面?
    
    Set taskGroup = Nothing
    Set cbsActions = Nothing
    Set gWind = Nothing '清除全局窗体引用
    
End Sub

Private Sub mTabWorkspace_RClick(ByVal Item As XtremeCommandBars.ITabControlItem)
    '右键多标签显示菜单
    If Not Item Is Nothing Then
        Item.Selected = True
        mTabWorkspace.Refresh
        mcbsPopupTab.ShowPopup
    End If
End Sub

Private Sub Picture1_Resize()
    '导航菜单面板中任务面板大小变化
    Me.TaskPanel1.Move 0, 0, Me.Picture1.ScaleWidth, Me.Picture1.ScaleHeight
End Sub

Private Sub TaskPanel1_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
    '导航菜单响应
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    
    Rem Debug.Print Me.ActiveForm.Name, Screen.ActiveForm.Name
    Rem Debug.Print Me.ActiveForm.ActiveControl.Name, Me.ActiveControl.Name
    If Me.CommandBars1.Actions(Item.ID).Enabled Then
        Call msLeftClick(Item.ID, Me.CommandBars1)
        If Me.CommandBars1.Actions(gID.PanePopupMenuNaviAutoFoldOther).Checked Then
            For Each taskGroup In Me.TaskPanel1.Groups
                If taskGroup.ID <> Item.Group.ID Then taskGroup.Expanded = False
            Next
        End If
    Else
        MsgBox "您目前无权打开此菜单！或者请联系管理员。", vbExclamation, "权限提醒"
    End If
    
    Set taskGroup = Nothing
End Sub

Private Sub Timer1_Timer(Index As Integer)
    Const conCon As Byte = 1    '连接状态检测间隔conConn秒
    Static byteCon As Byte
    Static byteChk As Byte
    
    '重启客户端程序
    '因在Winsock控件中的gArr的with语句中重启时无法清空gArr数组，权宜放此处
    If gVar.ClientReLoad Then
        Call msUnloadMe(True)
        Load Me
        Exit Sub
    End If
    
    '在登陆窗口中点击了关闭程序
    If gVar.UnloadFromLogin And Not gVar.ShowMainWindow Then '卸载登陆窗口+没有显示主窗体
        Call msUnloadMe(True)
        Exit Sub
    End If
      
    If gVar.ClientLoginCheckOver And (Not gVar.ShowMainWindow) And gVar.TCPStateConnected Then
        mXtrStatusBar.FindPane(gID.StatusBarPaneUserInfo).Text = gVar.UserFullName '主窗体状态中显示用户全名
        Me.Show '显示主窗体
        If gVar.ParaBlnWindowStartMinC Then
            Me.WindowState = vbMinimized '启动时最小化
        End If
        Call gfSendClientInfo(gVar.UserComputerName, gVar.UserLoginName, gVar.UserFullName, Me.Winsock1.Item(1)) '把用户登陆信息发送给服务端
        gVar.ShowMainWindow = True '显示主窗体标志。区别关闭程序时的主窗体状态
        Call msLoadUserAuthority(gVar.UserAutoID) '加载权限
        
        Dim frmUnload As Form  '卸载登陆窗口。不知为何，直接用Unload frmSysLogin不能卸载掉，没反应。
        For Each frmUnload In Forms
            If LCase(frmUnload.Name) = LCase("frmSysLogin") Then
                Unload frmUnload
                Exit For
            End If
        Next
    End If
    
    byteCon = byteCon + 1 '状态计时
    If Not gVar.TCPStateConnected And gVar.UpdateRunOver Then byteChk = byteChk + 1  '检查计时，只在无连接时
    
    If byteCon >= conCon Then
        If (Not gVar.UpdateRunOver) And (Not gfAppExist(gVar.EXENameOfUpdate)) Then '权且如此,仅判断进程是否存在是不全面的
            gVar.UpdateRunOver = True   '更新程序已运行完成标志
            Call gsOpenTheWindow("frmSysLogin", , vbNormal) ''显示登陆窗口
            gVar.ClientLoginShow = True '设置全局变量--已打开过登陆窗口
        End If
             
        With Me.Winsock1.Item(1)
            If .State = 7 Then      '已连接
                Call msSetClientState(vbGreen)  '设置连接状态
            ElseIf .State = 9 Then  '连接异常
                Call msSetClientState(vbRed)    '设置异常状态
            Else                    '未连接等
                Call msSetClientState(vbYellow) '设置未连接状态
            End If
        End With
        byteCon = 0 '清零静态累积变量
    End If
    
    If byteChk > (gVar.TCPWaitTime + 1) Then  '因为服务器端也是等待gVar.TCPWaitTime才断开连接，这里延迟一点
        byteChk = 0 '静态变量清零
        If Not gVar.TCPStateConnected Then  '未连接状态
            If gVar.ClientLoginCheckOver Then   '已校验过账号密码
                Call gsAlarmAndLogEx("与服务器建立连接失败，请确认服务端程序已启动", "连接警示")
                Call msUnloadMe(True)
            End If
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
            Call gfRestoreInfo(strGet, Me.Winsock1.Item(Index)) '解析文件信息
            
            If InStr(strGet, gVar.PTClientConfirm) > 0 Then '收到要回复服务端确认连接的信息
                Call gfSendInfo(gVar.PTClientIsTrue, Me.Winsock1.Item(Index))
                .Connected = True
                
            ElseIf InStr(strGet, gVar.PTConnectIsFull) Then '收到服务端发来的连接数已满
                Me.Timer1.Item(Index).Enabled = False
                Call msUnloadMe(True)
                MsgBox "客户端与服务端连接数受限，请其他用户退出后再试！", vbCritical, "连接数已满警告"
                Rem End '硬结束程序，以防万一？
            
            ElseIf InStr(strGet, gVar.PTDBDataSource) Then  '收到服务器发来的数据库连接信息
                Call gfRestoreDBInfo(strGet) '解析加密过的数据库连接信息
                gVar.RestoreDBInfoOver = True '数据库连接信息接收完成标志
                
            ElseIf InStr(strGet, gVar.PTConnectTimeOut) Then '连续连接时间已到
                Dim blnTimer As Boolean, tmrEn As VB.Timer
                For Each tmrEn In Me.Timer1 '重启客户端后老是发现timer1(1)不存在，权且用此检测
                    If tmrEn.Index = Index Then
                        blnTimer = True
                        Exit For
                    End If
                Next
                Set tmrEn = Nothing
                If blnTimer Then Me.Timer1.Item(Index).Enabled = False
                Me.Winsock1.Item(Index).Close
                Call msSetClientState(vbYellow) '设置未连接状态
                MsgBox "与服务器连续连接时间已到，请重新登陆！", vbExclamation, "连接时间限制提示"
                gVar.ClientReLoad = True
                If blnTimer Then Me.Timer1.Item(Index).Enabled = True
                
            ElseIf InStr(strGet, gVar.PTFileStart) > 0 Then '可以发送文件给服务端了的状态
                Call gsFileProgress(Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgress), _
                                    Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgressText), _
                                    ftZero, 0, .FileSizeTotal, 0) '初始化进度条
                gVar.FTIsOver = False   '设置传输结束标识为假
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index)) '发送文件给服务端
                Call gsFormEnable(Me, False)    '禁止客户端再操作
                
            ElseIf InStr(strGet, gVar.PTFileExist) > 0 Then '服务端发来客户端想要的文件存在的信号
                Dim strSize As String, lngInstrSize As Long
                
                lngInstrSize = InStr(strGet, gVar.PTFileSize) '获取客户端想要的存在于服务端的文件的大小
                If lngInstrSize > 0 Then
                    strSize = Mid(strGet, lngInstrSize + Len(gVar.PTFileSize))
                    If IsNumeric(strSize) Then
                        .FileSizeTotal = Val(strSize)
                        Call gsFileProgress(Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgress), _
                                            Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgressText), _
                                            ftZero, 0, .FileSizeTotal, 0) '初始化进度条
                        gVar.FTIsOver = False   '设置传输结束标识为假
                        Call gsFormEnable(Me, False)    '禁止客户端再操作
                        Debug.Print "Client: 开始接受服务端发来的文件," & Now
                        Rem 发送gVar.PTFileStart指令放在函数gfRestoreInfo中的【strType = gVar.PTFileReceive】判断中
                    End If
                End If
                
            ElseIf InStr(strGet, gVar.PTFileNoExist) > 0 Then   '
                MsgBox "需要的文件<" & .FileName & ">在服务端不存在！", vbExclamation, "文件警告"
                gArr(Index) = gArr(0)
                
            End If
            
            Debug.Print "Client: GetInfo--" & strGet, bytesTotal
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
            Call gsFileProgress(Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgress), _
                                Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgressText), _
                                ftRate, .FileSizeCompleted, .FileSizeTotal, 0)  '更新进度条
            
            If .FileSizeCompleted >= .FileSizeTotal Then    '传输完成后的一些处理
                Close #.FileNumber
                Call gsFormEnable(Me, True) '解除客户端的限制
                gArr(Index) = gArr(0)
                gVar.FTIsOver = True    '设置传输结束标识为真
                Call gfSendInfo(gVar.PTFileEnd, Me.Winsock1.Item(Index)) '发送结束标志
                Debug.Print "Client: File Received Over," & Now
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
            gVar.FTIsOver = False '设置传输结束标识为假
            gArr(Index).FileTransmitError = True    '异常结束
            Call gsFormEnable(Me, True)
        End If
        Call gsAlarmAndLogEx("与服务器连接发生异常！", "连接警报", True, vbCritical)
    End If
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    '发送完处理
    
    If Index = 0 Then Exit Sub
    With gArr(Index)
        If .FileTransmitState Then
            If .FileSizeCompleted < .FileSizeTotal Then '继续发送文件
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index))
                Call gsFileProgress(Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgress), _
                                    Me.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgressText), _
                                    ftRate, .FileSizeCompleted) '更新进度条的显示
            Else    '文件发送完成，恢复相关信息
                gArr(Index) = gArr(0)
                Call gsFormEnable(Me, True) '解锁窗口限制
                gVar.FTIsOver = True    '设置传输结束标识为真
                Debug.Print "Client: Send File Over ," & Now
            End If
        End If
    End With
End Sub
