VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
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
   StartUpPosition =   3  '����ȱʡ
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

Dim mlngID As Long  'ѭ������ID
Dim WithEvents mXtrStatusBar As XtremeCommandBars.StatusBar  '״̬���ؼ�
Attribute mXtrStatusBar.VB_VarHelpID = -1
Dim mcbsPopupIcon As XtremeCommandBars.CommandBar   '����ͼ��Pupup�˵�
Dim mcbsPopupNavi As XtremeCommandBars.CommandBar   '�����˵�������Popup�˵�
Dim mcbsPopupTab As XtremeCommandBars.CommandBar    '���ǩ�Ҽ�Popup�˵�
Dim WithEvents mTabWorkspace As XtremeCommandBars.TabWorkspace '���ǩ���ڿؼ�
Attribute mTabWorkspace.VB_VarHelpID = -1



Private Sub msAddAction(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '����CommandBars��Action
    
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.actions
    cbsBars.EnableActions   '����CommandBars��Actions����
    
'    cbsActions.Add "Id", "Caption", "TooltipText", "DescriptionText", "Category"   '����
    With cbsActions
        .Add gID.Sys, "ϵͳ", "", "", "ϵͳ"
        
        .Add gID.SysAuthChangePassword, "�����޸�", "", "", "frmSysAlterPWD"
        .Add gID.SysAuthDepartment, "���Ź���", "", "", "frmSysDepartment"
        .Add gID.SysAuthRole, "��ɫ����", "", "", "frmSysRole"
        .Add gID.SysAuthUser, "�û�����", "", "", "frmSysUser"
        .Add gID.SysAuthFunc, "Ȩ�޹���", "", "", "frmSysFunc"
        .Add gID.SysAuthLog, "��־����", "", "", "frmSysLog"
        
        .Add gID.SysLoginOut, "�˳�", "", "", ""
        .Add gID.SysLoginAgain, "����", "", "", ""
        
        .Add gID.SysExportMain, "����", "", "", ""
        .Add gID.SysExportToCSV, "������CSV", "", "", ""
        .Add gID.SysExportToExcel, "������Excel", "", "", ""
        .Add gID.SysExportToHTML, "������HTML", "", "", ""
        .Add gID.SysExportToPDF, "������PDF", "", "", ""
        .Add gID.SysExportToText, "������txt", "", "", ""
        .Add gID.SysExportToWord, "������Word", "", "", ""
        .Add gID.SysExportToXML, "������XML", "", "", ""
        
        .Add gID.SysPrintMain, "��ӡ", "", "", ""
        .Add gID.SysPrint, "��ӡ", "", "", ""
        .Add gID.SysPrintPageSet, "��ӡҳ������", "", "", ""
        .Add gID.SysPrintPreview, "��ӡԤ��", "", "", ""
        
        .Add gID.SysSearch, "���ڼ���", "", "", ""
        .Add gID.SysSearch1Label, "���봰�����ƹؼ���", "", "", ""
        .Add gID.SysSearch2TextBox, "�ؼ��������", "", "", ""
        .Add gID.SysSearch3Button, "��������", "", "", ""
        .Add gID.SysSearch4ListBoxCaption, "�������Ĵ��ڱ����б�", "", "", ""
        .Add gID.SysSearch4ListBoxFormID, "�������Ĵ��������б�", "", "", ""
        .Add gID.SysSearch5Go, "��ת��ѡ������", "", "", ""
        
        
        .Add gID.Wnd, "����", "", "", "����"
        
        .Add gID.WndThemeSkinSet, "������������...", "", "", ""
        .Add gID.WndResetLayout, "���ô��ڲ���", "", "", ""
        .Add gID.WndToolBarCustomize, "�Զ��幤������", "�Զ��幤����", "�Զ��幤����", ""
        .Add gID.WndToolBarList, "�������б�", "�������б�", "�������б�", ""
        .Add gID.WndOpenListCaption, "�Ѵ򿪴����б�", "", "", ""
        .Add gID.WndOpenListID, "", "", "", ""
        
        .Add gID.WndThemeCommandBars, "����������", "", "", ""
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
        
        .Add gID.WndThemeTaskPanel, "�����������", "", "", ""
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
        
        .Add gID.WndSon, "�Ӵ��ڿ���", "", "", ""
        .Add gID.WndSonCloseAll, "�ر������Ӵ���", "", "", ""
        .Add gID.WndSonCloseCurrent, "�رյ�ǰ�Ӵ���", "", "", ""
        .Add gID.WndSonCloseLeft, "�رյ�ǰ��ǩ����Ӵ���", "", "", ""
        .Add gID.WndSonCloseOther, "�ر������Ӵ���", "", "", ""
        .Add gID.WndSonCloseRight, "�رյ�ǰ��ǩ�Ҳ��Ӵ���", "", "", ""
        .Add gID.WndSonVbAllBack, "�ָ��Ӵ���", "", "", ""
        .Add gID.WndSonVbAllMin, "��С�������Ӵ���", "", "", ""
        .Add gID.WndSonVbArrangeIcons, "����������С��ͼ��", "", "", ""
        .Add gID.WndSonVbCascade, "�Ӵ��ڲ��", "", "", ""
        .Add gID.WndSonVbTileHorizontal, "�Ӵ���ˮƽƽ��", "", "", ""
        .Add gID.WndSonVbTileVertical, "�Ӵ��ڴ�ֱƽ��", "", "", ""
        
        .Add gID.WndThemeSkin, "��������", "", "", ""
        .Add gID.WndThemeSkinCodejock, "Codejock", "", "", ""
        .Add gID.WndThemeSkinOffice2007, "Office2007", "", "", ""
        .Add gID.WndThemeSkinOffice2010, "Office2010", "", "", ""
        .Add gID.WndThemeSkinVista, "Vista", "", "", ""
        .Add gID.WndThemeSkinWinXPLuna, "WinXPLuna", "", "", ""
        .Add gID.WndThemeSkinWinXPRoyale, "WinXPRoyale", "", "", ""
        .Add gID.WndThemeSkinZune, "Zune", "", "", ""
        
        
        .Add gID.Tool, "����", "", "", "����"
        .Add gID.toolOptions, "ѡ�", "ѡ��", "ѡ��", "frmOption"
        
        .Add gID.Help, "����", "", "", "����"
        .Add gID.HelpAbout, "���ڡ�", "", "", ""
        .Add gID.HelpDocument, "�����ĵ�", "", "", ""
        .Add gID.HelpUpdate, "���¼��", "", "", ""
                
        
        .Add gID.StatusBarPane, "״̬��", "", "", ""
        .Add gID.StatusBarPaneProgress, "������", "", "", ""
        .Add gID.StatusBarPaneUserInfo, "�û���Ϣ", "", "", ""
        .Add gID.StatusBarPaneTime, "����ʱ��", "", "", ""
        .Add gID.StatusBarPaneProgressText, "�������ٷֱ�ֵ", "", "", ""
        .Add gID.StatusBarPaneServerButton, "������/�Ͽ���ť", "", "", ""
        .Add gID.StatusBarPaneServerState, "����״̬", "", "", ""
        .Add gID.StatusBarPaneTime, "ϵͳʱ��", "", "", ""
        .Add gID.StatusBarPaneIP, "����IP��ַ", "", "", ""
        .Add gID.StatusBarPanePort, "���ӷ������˿�", "", "", ""
        .Add gID.StatusBarPaneConnectState, "���ӷ�����״̬", "", "", ""
        .Add gID.StatusBarPaneConnectButton, "��������������Ӱ�ť", "", "", ""
        .Add gID.StatusBarPaneReStartButton, "�����Զ�/�ֶ�����ģʽ�л���ť", "", "", ""
        
        .Add gID.IconPopupMenu, "����ͼ��˵�", "", "", ""
        .Add gID.IconPopupMenuMaxWindow, "��󻯴���", "", "", ""
        .Add gID.IconPopupMenuMinWindow, "��С������", "", "", ""
        .Add gID.IconPopupMenuShowWindow, "��ʾ����", "", "", ""

        .Add gID.Pane, "�������", "", "", ""
        .Add gID.PaneNavi, "�����˵�", "", "�����˵���ʾ/����", ""
        
        .Add gID.PanePopupMenuNavi, "�����˵�������Popup�˵�", "", "", ""
        .Add gID.PanePopupMenuNaviAutoFoldOther, "�Զ���£", "", "���ĳ�Ӳ˵�ʱ��£�����˵��������������˵�", ""
        .Add gID.PanePopupMenuNaviExpandALL, "ȫ��չ��", "", "չ�������˵����������˵�", ""
        .Add gID.PanePopupMenuNaviFoldALL, "ȫ����£", "", "��£�����˵����������˵�", ""
        
        .Add gID.TabWorkspacePopupMenu, "���ǩ�Ҽ��˵�", "", "", ""
        
'        .Add gID, "", "", "", ""
        
    End With
    
    '���cbsActions����������ToolTipText��DescriptionText��Key��Category
    For Each cbsAction In cbsActions
        With cbsAction
            If .ID < 20000 Then
                .ToolTipText = .Caption
                .DescriptionText = .ToolTipText
                .Key = .Category    'Ϊ�˵�ʱ�������ã�����Actionʱ������������Category��
                If LCase(Left(.Key, 3)) = "frm" Then
                    Select Case .ID
                        Case gID.toolOptions, gID.SysAuthChangePassword
                            'һЩ����Ҫ��Ȩ�޿��ƵĴ���
                        Case Else '�ܿ��ƴ���
                            cbsAction.Enabled = False '�Ƚ���ҪȨ�޿��ƵĴ��ڣ�����Ȩ��ʱ�ٽ���
                    End Select
                End If
                .Category = cbsActions((.ID \ 1000) * 1000).Category
            End If
        End With
    Next
    
    '���ϵ�е�cbsActions���������Ե���������
    With cbsActions
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Action(mlngID).DescriptionText = .Action(gID.WndThemeCommandBars).Caption & "����Ϊ��" & .Action(mlngID).DescriptionText
            .Action(mlngID).ToolTipText = .Action(mlngID).DescriptionText
        Next
    End With
    
    Set cbsAction = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddDesignerControls(ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars�Զ���Ի���������������
    
    Dim cbsControls As XtremeCommandBars.CommandBarControls
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.actions
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
    '�����������
    
    Dim paneNavigation As XtremeDockingPane.Pane
    Dim cbsActions As XtremeCommandBars.CommandBarActions
    
    Set cbsActions = cbsBars.actions
'    Me.Picture1.Appearance = 0
'    Me.Picture1.BackColor = Me.BackColor
    
    With Me.DockingPane1
        .SetCommandBars cbsBars '�������ֿ���ͬʱʹ�ñ�����ô���ã���CommandBars�ؼ���DockingPane�ؼ��Ķ���
        With .Options
            .AlphaDockingContext = True
            .ShowDockingContextStickers = True
            .StickerStyle = StickerStyleVisualStudio2008 '����ʹAlphaDockingContext��ShowDockingContextStickers��ΪTrue
        End With
        Set paneNavigation = .CreatePane(gID.PaneNavi, 260, 240, DockLeftOf)
        cbsBars.actions(gID.PaneNavi).Checked = True
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
    '������ݼ�
    
    With cbsBars.KeyBindings
        .AddShortcut gID.SysLoginOut, "F10"
    End With
    
End Sub

Private Sub msAddMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '�����˵���
    
    Dim cbsMenuBar As XtremeCommandBars.MenuBar
    Dim cbsMenuMain As XtremeCommandBars.CommandBarPopup
    Dim cbsMenuCtrl As XtremeCommandBars.CommandBarControl
    Dim cbsMenuCtrlTemp As XtremeCommandBars.CommandBarControl
    
    Set cbsMenuBar = cbsBars.ActiveMenuBar
    cbsMenuBar.ShowGripper = False  '����ʾ���϶����Ǹ������
    cbsMenuBar.EnableDocking xtpFlagStretched     '�˵�����ռһ���Ҳ��������϶�
    
    'ϵͳ���˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Sys, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.SysAuthChangePassword, ""
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysAuthDepartment, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysAuthRole, ""
        .Add xtpControlButton, gID.SysAuthUser, ""
        .Add xtpControlButton, gID.SysAuthFunc, ""
        .Add xtpControlButton, gID.SysAuthLog, ""
                
        Set cbsMenuCtrlTemp = .Add(xtpControlButtonPopup, gID.SysExportMain, "����")
        cbsMenuCtrlTemp.BeginGroup = True
        With cbsMenuCtrlTemp.CommandBar.Controls
            Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysExportToCSV, "")
            cbsMenuCtrl.BeginGroup = True
            For mlngID = gID.SysExportToExcel To gID.SysExportToWord
                .Add xtpControlButton, mlngID, ""
            Next
        End With
        
        Set cbsMenuCtrlTemp = .Add(xtpControlButtonPopup, gID.SysPrintMain, "��ӡ")
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
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.WndThemeSkinSet, "" 'Ƥ������
        .Add xtpControlButton, gID.WndResetLayout, "" '���ò���
        .Add xtpControlButton, gID.PaneNavi, ""  '�����˵���ʾ/����
        
        '����ID XTP_ID_CUSTOMIZE=35001�Զ��幤����
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.WndToolBarCustomize, "")
        cbsMenuCtrl.BeginGroup = True
    
        '����ID XTP_ID_TOOLBARLIST=59392�������б�
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndToolBarList, "")
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
        
        'CommandBars�����������Ӳ˵�
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeCommandBars, "")
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
                .Add xtpControlButton, mlngID, ""
            Next
        End With
        
        'TaskPanel�����������
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeTaskPanel, "")
        cbsMenuCtrl.BeginGroup = True
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
                .Add xtpControlButton, mlngID, ""
            Next
        End With
        
        '�Ӵ��ڿ���
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndSon, "")
        cbsMenuCtrl.BeginGroup = True
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndSonCloseAll To gID.WndSonVbTileVertical
                .Add xtpControlButton, mlngID, ""
            Next
            .Find(, gID.WndSonVbAllBack).BeginGroup = True
        End With
                
        '�Ӵ����б�
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndOpenListCaption, "")
        cbsMenuCtrl.BeginGroup = True
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, gID.WndOpenListID, ""
    End With
    
    '���߲˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Tool, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.toolOptions, ""
    
    '�������˵�
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
    '��������ͼ���Ҽ�����ʽ�˵�
    Set mcbsPopupIcon = cbsBars.Add(cbsBars.actions(gID.IconPopupMenu).Caption, xtpBarPopup)
    With mcbsPopupIcon.Controls
        .Add xtpControlButton, gID.IconPopupMenuMaxWindow, ""
        .Add xtpControlButton, gID.IconPopupMenuMinWindow, ""
        .Add xtpControlButton, gID.IconPopupMenuShowWindow, ""
        .Add xtpControlButton, gID.SysLoginAgain, ""
        .Add xtpControlButton, gID.SysLoginOut, ""
    End With
    
    '���������˵�����ϱ������ϵĵ���ʽ�˵�
    Set mcbsPopupNavi = cbsBars.Add(cbsBars.actions(gID.PanePopupMenuNavi).Caption, xtpBarPopup)
    With mcbsPopupNavi.Controls
        .Add xtpControlButton, gID.PanePopupMenuNaviAutoFoldOther, ""
        .Add xtpControlButton, gID.PanePopupMenuNaviExpandALL, ""
        .Add xtpControlButton, gID.PanePopupMenuNaviFoldALL, ""
    End With
    
    '�������ǩ�ϵ��Ҽ��˵�
    Set mcbsPopupTab = cbsBars.Add(cbsBars.actions(gID.TabWorkspacePopupMenu).Caption, xtpBarPopup)
    With mcbsPopupTab.Controls
        For mlngID = gID.WndSonCloseAll To gID.WndSonVbTileVertical
            .Add xtpControlButton, mlngID, ""
        Next
        .Find(, gID.WndSonVbAllBack).BeginGroup = True
    End With
        
End Sub

Private Sub msAddTaskPanelItem(ByRef tskPanel As XtremeTaskPanel.TaskPanel)
    '���������˵�
    
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    Dim taskItem As XtremeTaskPanel.TaskPanelGroupItem
    Dim cbsActions As XtremeCommandBars.CommandBarActions
    Dim lngID As Long, lngLeftMargins As Long, L As Long, T As Long, R As Long, b As Long
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim imgIcon As MSComctlLib.ListImage
    
    Set cbsActions = Me.CommandBars1.actions
    
    '����ϵͳ�˵�
    Set taskGroup = tskPanel.Groups.Add(gID.Sys, cbsActions(gID.Sys).Caption)
    With taskGroup.Items
        Set taskItem = .Add(gID.SysAuthChangePassword, cbsActions(gID.SysAuthChangePassword).Caption, xtpTaskItemTypeLink)
        taskItem.GetRect L, T, R, b 'Ϊ�����кÿ���ÿһ���Ӳ˵�ʹ��ͬ������������,��Ҫ��Ϊ�˻�ȡLֵ(��߾�)
        lngLeftMargins = L
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
    
    '�������ڲ˵�
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
    
    '�������߲˵�
    Set taskGroup = tskPanel.Groups.Add(gID.Tool, cbsActions(gID.Tool).Caption)
    With taskGroup.Group.Items
        .Add gID.toolOptions, cbsActions(gID.toolOptions).Caption, xtpTaskItemTypeLink
    End With
    
    
    '���������˵�
    Set taskGroup = tskPanel.Groups.Add(gID.Help, cbsActions(gID.Help).Caption)
    With taskGroup.Group.Items
        For mlngID = gID.HelpAbout To gID.HelpUpdate
            .Add mlngID, cbsActions(mlngID).Caption, xtpTaskItemTypeLink
        Next
    End With
    
    
    '���GroupItemͼ��
    tskPanel.SetImageList Me.ImageList1 '��ͼ�꼯��
    For Each taskGroup In tskPanel.Groups
        For Each taskItem In taskGroup.Items
            For Each imgIcon In Me.ImageList1.ListImages
                If Val(imgIcon.Tag) = taskItem.ID Then '��Ԥ����ImageList1�ؼ�������ÿ��ͼ���TagֵΪGroupItem��IDֵ
                    taskItem.IconIndex = imgIcon.Index
                    Exit For
                End If
            Next
        Next
    Next
    
    'ͬ��Ȩ��
    For Each cbsAction In cbsActions
        If Not tskPanel.Find(cbsAction.ID) Is Nothing Then '������ÿ��Action��Ӧһ��GroupItem
            tskPanel.Find(cbsAction.ID).Enabled = cbsAction.Enabled
        End If
    Next
    
    '�����۵��˵�״̬
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
    '����������
    
    Dim cbsBar As XtremeCommandBars.CommandBar
    Dim cbsCtr As XtremeCommandBars.CommandBarControl
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.actions
    
    'ϵͳ����������
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
    
    '����������
    Set cbsBar = cbsBars.Add(cbsActions(gID.WndThemeCommandBars).Caption, xtpBarTop)
    With cbsBar.Controls
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '�����˵�����
    Set cbsBar = cbsBars.Add(cbsActions(gID.WndThemeTaskPanel).Caption, xtpBarTop)
    With cbsBar.Controls
        For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '���ڼ���������
    Set cbsBar = cbsBars.Add(cbsActions(gID.SysSearch).Caption, xtpBarTop)
    With cbsBar.Controls
        .Add xtpControlLabel, gID.SysSearch1Label, ""
        Set cbsCtr = .Add(xtpControlEdit, gID.SysSearch2TextBox, "")
        cbsCtr.Width = 200
        cbsCtr.EditHint = "���봰�ڱ���ؼ���"
        .Add xtpControlButton, gID.SysSearch3Button, ""
        Set cbsCtr = .Add(xtpControlComboBox, gID.SysSearch4ListBoxCaption, "")
        cbsCtr.Width = 200
        cbsCtr.EditHint = "���б���ѡ��һ�����ڱ���"
        Set cbsCtr = .Add(xtpControlComboBox, gID.SysSearch4ListBoxFormID, "")
        cbsCtr.Visible = False
        .Add xtpControlButton, gID.SysSearch5Go, ""
    End With
    
    Set cbsBar = Nothing
    Set cbsCtr = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddXtrStatusBar(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '����״̬��
    
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    Dim BarPane As XtremeCommandBars.StatusBarPane
    
    Set cbsActions = cbsBars.actions
    Set mXtrStatusBar = cbsBars.StatusBar
    With mXtrStatusBar
        .AddPane 0      'ϵͳPane����ʾCommandBarActions��Description
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
        
        .AddPane 59137  'CapsLock����״̬
        .AddPane 59138  'NumLK����״̬
        .AddPane 59139  'ScrLK����״̬
        .FindPane(0).Caption = "Idle Text"
        .FindPane(59137).Caption = "Caps Lock��״̬"
        .FindPane(59138).Caption = "Num LocK��״̬"
        .FindPane(59139).Caption = "Scroll LocK��״̬"
        
        .Visible = True
        .EnableCustomization True
    End With
    
    For Each BarPane In mXtrStatusBar     '����Caption��ToolTip��Alignment����
        If Len(BarPane.Caption) = 0 Then BarPane.Caption = cbsActions(BarPane.ID).Caption
        BarPane.ToolTip = BarPane.Caption
        If BarPane.ID <> 0 Then BarPane.Alignment = xtpAlignmentCenter
    Next
    
    Set cbsActions = Nothing
    Set BarPane = Nothing
End Sub

Private Sub msConnectToServer(ByRef sckCon As MSWinsockLib.Winsock, Optional ByVal blnConnect As Boolean = False)
    '�����������������
    
    If Not blnConnect Then Exit Sub
    With sckCon
        If .State <> 0 Then .Close
        .RemoteHost = gVar.TCPSetIP
        .RemotePort = gVar.TCPSetPort
        .Connect
    End With
End Sub

Private Sub msLeftClick(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars����������Ӧ��������
    
    Dim strKey As String
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.actions
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
                If MsgBox("ȷ�����������ͻ��˳�����", vbQuestion + vbOKCancel, "����������ѯ��") = vbOK Then
                    Call msUnloadMe(True)
                    Load Me
                End If
            Case .SysLoginOut
                If MsgBox("ȷ���˳��ͻ��˳�����", vbQuestion + vbOKCancel, "�ر�������ѯ��") = vbOK Then
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
                strAbout = "���ƣ�" & App.Title & vbCrLf & _
                           "�汾��" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                           "��Ȩ���У�XMH"
                MsgBox strAbout, vbInformation, "����" & App.Title
            
            Case .HelpUpdate
                If gfAppExist(gVar.EXENameOfUpdate) Then
                    MsgBox "���³����ڽ������Ѵ��ڣ������ظ��򿪣�", vbInformation, "�Ѵ���ʾ"
                Else
                    If MsgBox("���³�����ڼ���Ҫռ��һ���û�����ȷ�����ڽ��и��¼����", vbQuestion + vbOKCancel, "�򿪸��³���ѯ��") = vbOK Then
                        Dim strUP As String
                        strUP = gVar.AppPath & gVar.EXENameOfUpdate & " " & gVar.EXENameOfClient '��ʾ�򿪸��³���������
                        Call msOpenUpdate(strUP) '�������д򿪸��³���
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
                        If MsgBox("�Ƿ񽫵�ǰ������ݵ�����txt�ı��ĵ���", vbQuestion + vbYesNo, "ѯ��") = vbYes Then Call gsGridToText(Me.ActiveForm.ActiveControl)
                    Case .SysExportToWord
                        If MsgBox("�Ƿ񽫵�ǰ������ݵ�����Word�ĵ���", vbQuestion + vbYesNo, "ѯ��") = vbYes Then Call gsGridToWord(Me.ActiveForm.ActiveControl)
                        
                    Case .SysPrint
                        If MsgBox("ȷ����ӡ��ǰ���������", vbQuestion + vbOKCancel, "��ӡѯ��") = vbOK Then Call gsGridPrint(Me.ActiveForm.ActiveControl)
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
                            Case .toolOptions, .SysAuthChangePassword
                                Call gsOpenTheWindow(strKey, vbModal, vbNormal)
                            Case Else
                                Call gsOpenTheWindow(strKey)
                        End Select
                    End If
                Else
                    MsgBox "��" & cbsActions(CID).Caption & "������δ���壡", vbExclamation, "�����"
                End If
        End Select
    End With
    
    Set cbsActions = Nothing
End Sub

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    '��ע����м��ز���ֵ�����ñ�����
    Dim tempVal
    
    If Not blnLoad Then Exit Sub
    
    Rem On Error Resume Next    '��/���ܺ������̿������쳣
    With gVar
        .ParaBlnWindowCloseMin = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, 1))    '�ر�ʱ��С��
        .ParaBlnWindowMinHide = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, 1))  '��С��ʱ����
        
        .TCPDefaultIP = Me.Winsock1.Item(1).LocalIP '����IP��ַ
        .TCPSetIP = gfCheckIP(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPDefaultIP)) 'Ҫ���ӵķ����IP��ַ
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , .TCPDefaultPort, 10000, 65535) 'Ҫ���ӵķ������˿�
        
        .ParaBlnAutoStartupAtBoot = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, 0))  '�����Զ�����
        .ParaBlnUserAutoLogin = Val(GetSetting(.RegAppName, .RegSectionUser, .RegKeyParaUserAutoLogin, 0)) '�Զ���½
        .ParaBlnRememberUserList = Val(GetSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserList, 0)) '��ס�û���
        .ParaBlnRememberUserPassword = Val(GetSetting(.RegAppName, .RegSectionUser, .RegKeyParaRememberUserPassword, 0)) '��ס����
        
        .UserLoginIP = .TCPDefaultIP '����IP��ֵ����һ������
        .UserComputerName = gfBackComputerInfo(ciComputerName) '��ȡ�������
        
'''        '�ɷ���˷��������ͻ���
'''        .ConSource = gfCheckIP(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .TCPSetIP))   '����������/IP
'''        .ConDatabase = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString("dbTest", .EncryptKey)), .EncryptKey)    '���ݿ���
'''        .ConUserID = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString("123", .EncryptKey)), .EncryptKey)  '��½��
'''        .ConPassword = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString("888888", .EncryptKey)), .EncryptKey)    '��½����
        
        
    End With
End Sub

Private Sub msLoadUserAuthority(ByVal strUID As String)
    'Ȩ�޿���
    
    Const strFRM As String = "frm"
    Dim cbsAction As CommandBarAction
    Dim strSQL As String, strKey As String, strSys As String
    
    strUID = Trim(strUID)
    If Len(strUID) = 0 Then Exit Sub
    
    strSys = LCase(gVar.UserLoginName)
    If strSys = LCase(gVar.AccountAdmin) Or strSys = LCase(gVar.AccountSystem) Then   '�����ڶ������û�ӵ������Ȩ��
        For Each cbsAction In gWind.CommandBars1.actions
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
                For Each cbsAction In Me.CommandBars1.actions
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
    '�������д򿪸��³���
    
    If Not gfShell(strCmd) Then
        Call gsAlarmAndLogEx("���³��������쳣", "���¼��ʧ��", True, vbCritical)
    End If
End Sub

Private Sub msPopupMenuNavi(ByVal PID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    '�����˵��ϵ����˵�����Ӧ
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    
    Select Case PID
        Case gID.PanePopupMenuNaviAutoFoldOther
            cbsBars.actions(PID).Checked = Not cbsBars.actions(PID).Checked
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
    '���ô��ڲ��֣�CommandBars��Dockingpane�ؼ�����
    
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
    '���ڱ������
    
    
    Dim strName As String
    Dim cbsAction As CommandBarAction
    Dim cbsCtrlCaption As CommandBarComboBox
    Dim cbsCtrlFormID As CommandBarComboBox
    Dim blnClear As Boolean
    
    strName = LCase(Trim(cbsBars.FindControl(xtpControlEdit, gID.SysSearch2TextBox).Text))
    If Len(strName) = 0 Then Exit Sub
    
    Set cbsCtrlCaption = cbsBars.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxCaption)
    Set cbsCtrlFormID = cbsBars.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxFormID)
    
    For Each cbsAction In cbsBars.actions
        If cbsAction.ID < 20000 Then     '���д��ڵ�IDС��2000
            If LCase(Left(cbsAction.Key, 3)) = "frm" Then   '���ڵ�Name������frm��ͷ
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
    '����״̬���пͻ������ӷ���˵�״̬
    
    Dim paneState As XtremeCommandBars.StatusBarPane
    
    Set paneState = mXtrStatusBar.FindPane(gID.StatusBarPaneConnectState)
    With paneState
        If ColorSet = vbGreen Then  '������
            .Text = gVar.ClientStateConnected
            .BackgroundColor = ColorSet
            gVar.TCPStateConnected = True   '�ñ�����¼״̬
        Else    '����
            gVar.TCPStateConnected = False
            If ColorSet = vbRed Then    '�����쳣
                .Text = gVar.ClientStateConnectError
                .BackgroundColor = ColorSet
            Else    'δ���ӵ�
                .Text = gVar.ClientStateDisConnected
                .BackgroundColor = vbYellow
            End If
        End If
    End With
    Set paneState = Nothing
    
End Sub

Private Sub msThemeTaskPanel(ByVal LID As Long, ByRef cbsSet As XtremeCommandBars.CommandBars)
    '���������������
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
        cbsSet.actions(mlngID).Checked = False
    Next
    cbsSet.actions(LID).Checked = True
End Sub

Private Sub msUnloadMe(Optional ByVal blnUnload As Boolean = True)
    'ж�ش���
    If Not blnUnload Then Exit Sub
    gVar.CloseWindow = True
    If Forms.Count > 1 Then '�ȹر��������壬�ٹر�������
        Dim frmUld As Form
        
        For Each frmUld In Forms
            If frmUld.Name <> Me.Name Then Unload frmUld
        Next
        Set frmUld = Nothing
    End If
    Unload Me
End Sub

Private Sub msWindowControl(ByVal WID As Long)
    '�Ӵ��ڿ���
    
    Dim frmTag As Form
    Dim C As Long
    Dim itemCur As XtremeCommandBars.TabControlItem
        
    With gID
        Select Case WID
            Case .WndSonCloseAll    '�ر����д���
                For Each frmTag In Forms
                    If frmTag.Name <> gWind.Name Then Unload frmTag
                Next
            Case .WndSonCloseCurrent    '�رյ�ǰ����
                If Not ActiveForm Is Nothing Then Unload ActiveForm
            Case .WndSonCloseLeft   '�ر���ര��
                If Forms.Count > 2 Then
                    Set itemCur = mTabWorkspace.Selected
                    itemCur.Tag = "c"   '��ǵ�ǰ���ڣ���ΪIndexֵ�ڴ��������仯ʱ��仯��������ΪΨһ�ж�����
                    For C = 0 To mTabWorkspace.ItemCount - 1
                        If mTabWorkspace.Item(0).Tag = itemCur.Tag Then
                            itemCur.Tag = ""    '�ǵ���ա�Tag����Ĭ��ֵ���ǿ��ַ���
                            Exit For
                        Else
                            mTabWorkspace.Item(0).Selected = True   '����Ҫɾ���Ĵ���
                            Unload ActiveForm
                        End If
                    Next
                End If
            Case .WndSonCloseOther  '�ر���������
                If Forms.Count > 1 Then
                    For Each frmTag In Forms
                        If frmTag.Name <> gWind.Name Then
                            If Not (frmTag.Name = ActiveForm.Name And frmTag.Caption = ActiveForm.Caption) Then
                                Unload frmTag
                            End If
                        End If
                    Next
                End If
            Case .WndSonCloseRight  '�ر��Ҳര��
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
    '������¼�
    Call msLeftClick(Control.ID, Me.CommandBars1)
End Sub


Private Sub CommandBars1_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'CommandBars�ؼ���Action״̬���л�
    
    Dim blnFC As Boolean    '�ж��Ƿ�ΪFC���
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    Dim blnMainWindow As Boolean '�ж��������Ƿ���ȫ���������
    Dim tskItems As XtremeTaskPanel.TaskPanelGroupItems   '�����˵�����
    
    Set cbsActions = Me.CommandBars1.actions
    Set tskItems = Me.TaskPanel1.Groups.Find(gID.Sys).Items
    
    If Not Me.ActiveForm Is Nothing Then
        If Not Me.ActiveForm.ActiveControl Is Nothing Then
            blnFC = TypeOf Me.ActiveForm.ActiveControl Is FlexCell.Grid    '��ǰ��ؼ���FC���
        End If
    End If
    
    blnMainWindow = gVar.ShowMainWindow
    
    With gID
        For mlngID = .SysExportToCSV To .SysExportToWord
            cbsActions(mlngID).Enabled = blnFC  '��ؼ���FC����򼤻��ӦAction������ʹ�䲻����
            tskItems.Find(mlngID).Enabled = blnFC
        Next
        For mlngID = .SysPrintPageSet To .SysPrint
            cbsActions(mlngID).Enabled = blnFC
            tskItems.Find(mlngID).Enabled = blnFC
        Next
        For mlngID = .IconPopupMenuMaxWindow To .IconPopupMenuShowWindow
            cbsActions(mlngID).Enabled = blnMainWindow  '������δ�������֮ǰ������ͼ��˵�ĳЩ�����ã�
        Next
    End With
    
    Set cbsActions = Nothing
    Set tskItems = Nothing
End Sub


Private Sub DockingPane1_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    '�����˵����������ʾ�����أ��رգ�
    If Action = PaneActionClosed Then
        If Pane.ID = gID.PaneNavi Then
            Me.CommandBars1.actions(Pane.ID).Checked = False
        End If
    End If
End Sub

Private Sub DockingPane1_PanePopupMenu(ByVal Pane As XtremeDockingPane.IPane, ByVal X As Long, ByVal Y As Long, Handled As Boolean)
    '�����˵���������ϱ������ϵ���ʽ�˵�
    If Pane.ID = gID.PaneNavi Then
        mcbsPopupNavi.ShowPopup , X * 15, Y * 15
    End If
End Sub

Private Sub MDIForm_Load()
    '�������
    
    Dim cbsBars As XtremeCommandBars.CommandBars
    Dim strUpdate As String
    
    ReDim gArr(1)   '��ʼ�����顣�ͻ���ͳһʹ���±�1������Timer1�ؼ���Winsocket�ؼ�
    Timer1.Item(1).Interval = 1000  '��ʱ��ѭ��ʱ��
    Set gWind = Me  'ָ���������ȫ�����ö���
    
    XtremeCommandBars.CommandBarsGlobalSettings.App = App 'һ��Ĭ������
    Set cbsBars = Me.CommandBars1
    
    Call Main   '��ʼ��ȫ�ֹ��ñ���
    Call msLoadParameter(True)  '�������ò���
    Call msAddAction(cbsBars)   '����Actions����
    Call msAddMenu(cbsBars)     '�����˵���
    Call msAddToolBar(cbsBars)  '����������
    Call msAddPopupMenu(cbsBars)    '��������ͼ��Ĳ˵�
    Call msAddXtrStatusBar(cbsBars) '����״̬��
    Call msAddKeyBindings(cbsBars)  '��ӿ�ݼ�,�ŵ�LoadCommandBars�������������Ч������
    Call msAddDesignerControls(cbsBars) 'CommandBars�Զ���Ի�����ʹ�õ�
    Call msAddDockingPane(cbsBars) '��������ҷ�ĸ������
    Call msAddTaskPanelItem(Me.TaskPanel1)  '���������˵�
    
    cbsBars.AddImageList ImageList1         'ʹCommandBars�ؼ�ƥ��ImageList�ؼ���ͼ��
    cbsBars.EnableCustomization True        '����CommandBars�Զ��壬��������÷�������CommandBars�趨֮��
    cbsBars.Options.UpdatePeriod = 250      '����CommandBars��Update�¼���ִ�����ڣ�Ĭ��100ms
    
    Set mTabWorkspace = cbsBars.ShowTabWorkspace(True) '���ô��ڶ��ǩģʽ
    mTabWorkspace.Flags = xtpWorkspaceShowActiveFiles Or xtpWorkspaceShowCloseSelectedTab '��ʾ������б���ǰ������ʾ�رհ�ť
    
    
    Call gsLoadSkin(Me, Me.SkinFramework1, sMSO7, True)  '���ش�������
    
    '���ع���������
    Call gsThemeCommandBar(Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientCommandbarsTheme, gID.WndThemeCommandBarsRibbon)), cbsBars)
    
    'ע�����Ϣ����-CommandBars����
    Call cbsBars.LoadCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    
'    '���������ʽ����δ֪ԭ�򣬼��غ�����ӵ�TaskPanel���ݶ�Ĩ����
'    Call Me.DockingPane1.LoadState(gVar.RegKeyDockingPane, gVar.RegAppName, gVar.RegKeyDockPaneClientSetting)
    
    '���ص����˵�������
    Call msThemeTaskPanel(gfGetRegNumericValue(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelTheme, _
        True, gID.WndThemeTaskPanelNativeWinXP, gID.WndThemeTaskPanelListView, gID.WndThemeTaskPanelVisualStudio2010), cbsBars)
    cbsBars.actions(gID.PanePopupMenuNaviAutoFoldOther).Checked = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelAutoFold, 1))
    '�˵����۵�����msAddTaskPanelItem������
    
    
    Set cbsBars = Nothing   '����ʹ����Ķ���
    
    Call gsFormSizeLoad(Me, False) 'ע�����Ϣ����-����λ�ô�С
    
    '���¼��
    If gfAppExist(gVar.EXENameOfUpdate) Then
        Me.Timer1.Item(1).Enabled = False
        MsgBox "���³����ռ��һ���û�����" & vbCrLf & "���Ƚ����Ѵ��ڵĸ��³�����̺��ٴ������", vbExclamation, "�����Ѵ�����"
        Call msUnloadMe(True)
        End
    Else
        strUpdate = gVar.AppPath & gVar.EXENameOfUpdate & " " & gVar.EXENameOfClient & _
                gVar.CmdLineSeparator & gVar.CmdLineParaOfHide      '������ʽ�򿪸��¼������������
        Call msOpenUpdate(strUpdate) '����������ʽ�򿪸��³���
    End If
    
    '����Ƿ�Ϊ���ð�*******************************
    '==============================================
    
    
    '==============================================
    
    Call gfNotifyIconAdd(Me)    '�������ͼ��
    Me.Hide 'δ��½������ʾ������
    
    If LCase(App.EXEName & ".exe") <> LCase(gVar.EXENameOfClient) Then
        Call gsAlarmAndLogEx("���������޸Ŀ�ִ�е�Ӧ�ó����ļ�����", "���ؾ���", True, vbCritical)
        Call msUnloadMe(True)    '��ֹexe�ļ�������
    End If
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '��Ӧ����ͼ��������Ҽ����������̲˵�
    Dim sngMsg As Single
    
    If Y <> 0 Then Exit Sub    '�ƺ��˾������ס���һ����������ͼ���ϣ������ڴ�����
    sngMsg = X / Screen.TwipsPerPixelX
    Select Case sngMsg
        Case WM_RBUTTONUP
            mcbsPopupIcon.ShowPopup  '�Ҽ�����Popup�˵�
        Case WM_LBUTTONDBLCLK   '���˫������ͼ��ʱ ��������ʾ/��С�� �л�
            Rem If Button <> vbLeftButton Then Exit Sub '���ڲ���ԭ���ż���Զ���С���ˣ��ƺ��˾������ס
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
    '�ж��Ƿ�����Ҫ�رմ���
    
    If gVar.ParaBlnWindowCloseMin Then
        If Not gVar.CloseWindow Then
            Cancel = True
            Me.WindowState = vbMinimized
        End If
        gVar.CloseWindow = False
    Else
        If Not gVar.CloseWindow Then
            If MsgBox("�Ƿ���С�����ڣ�", vbQuestion + vbYesNo, "�رջ���С��") = vbYes Then
                Cancel = True
                Me.WindowState = vbMinimized
            End If
        End If
    End If
End Sub

Private Sub MDIForm_Resize()
    '������С����ʾ
    If Me.Visible And Me.WindowState = vbMinimized Then
        If gVar.ParaBlnWindowMinHide Then
            Me.Hide
            Call gfNotifyIconBalloon(Me, "��С����ϵͳ����ͼ����", "��С����ʾ")
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'ж�ش���ʱ������Ϣ
    Dim resetNotifyIconData As gtypeNOTIFYICONDATA
    Dim gVarClear As gtypeCommonVariant
    Dim lngValue As Long
    Dim cbsActions As XtremeCommandBars.CommandBarActions
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    
    Set cbsActions = Me.CommandBars1.actions
    
    '����ע�����Ϣ-CommandBars����
    Call Me.CommandBars1.SaveCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    lngValue = gID.WndThemeCommandBarsVS2008 '����������
    For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        If cbsActions(mlngID).Checked Then
            lngValue = mlngID
            Exit For
        End If
    Next
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientCommandbarsTheme, lngValue)
        
    '���渡�������ʽ
    Call Me.DockingPane1.SaveState(gVar.RegKeyDockingPane, gVar.RegAppName, gVar.RegKeyDockPaneClientSetting)
    
    '���浼���˵���ʽ
    lngValue = IIf(cbsActions(gID.PanePopupMenuNaviAutoFoldOther).Checked, 1, 0) '�Զ��۵�
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelAutoFold, lngValue)
    lngValue = gID.WndThemeTaskPanelNativeWinXP 'TaskPanel����
    For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        If cbsActions(mlngID).Checked Then
            lngValue = mlngID
            Exit For
        End If
    Next
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientTaskPanelTheme, lngValue)
    
    For Each taskGroup In Me.TaskPanel1.Groups '�˵��۵�״̬
        lngValue = IIf(taskGroup.Expanded, 1, 0)
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, "TP" & CStr(taskGroup.ID), lngValue)
    Next
    
    
    Call gsFormSizeSave(Me, False) '����ע�����Ϣ-����λ�ô�С
    Call gsSaveCommandbarsTheme(Me.CommandBars1, False)   '����CommandBars�ķ������
    
    If gfAppExist(gVar.EXENameOfUpdate) Then '�������һ�����³�����ر�
        If Not gfCloseApp(gVar.EXENameOfUpdate) Then
            Call gsAlarmAndLogEx("����˳�ʱ�޷�ͬʱ�رո��³���", "�رո��³����쳣", False)
        End If
    End If
    
    gArr(1) = gArr(0) '����ļ����������е���Ϣ
    gVar = gVarClear '���gVar���ñ���
    
    Call SkinFramework1.LoadSkin("", "")    '���Ƥ��
    Set mXtrStatusBar = Nothing  '���״̬��
    Set mcbsPopupIcon = Nothing '���Popup�˵�
    Call gfNotifyIconDelete(Me) 'ɾ������ͼ��
    gNotifyIconData = resetNotifyIconData   '�������������Ϣ��������������ʱ���Զ�����������ֻ�ܷ��Ͼ�ɾ������ͼ�����ĺ���?
    
    Set taskGroup = Nothing
    Set cbsActions = Nothing
    Set gWind = Nothing '���ȫ�ִ�������
    
End Sub

Private Sub mTabWorkspace_RClick(ByVal Item As XtremeCommandBars.ITabControlItem)
    '�Ҽ����ǩ��ʾ�˵�
    If Not Item Is Nothing Then
        Item.Selected = True
        mTabWorkspace.Refresh
        mcbsPopupTab.ShowPopup
    End If
End Sub

Private Sub Picture1_Resize()
    '�����˵��������������С�仯
    Me.TaskPanel1.Move 0, 0, Me.Picture1.ScaleWidth, Me.Picture1.ScaleHeight
End Sub

Private Sub TaskPanel1_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
    '�����˵���Ӧ
    Dim taskGroup As XtremeTaskPanel.TaskPanelGroup
    
    Rem Debug.Print Me.ActiveForm.Name, Screen.ActiveForm.Name
    Rem Debug.Print Me.ActiveForm.ActiveControl.Name, Me.ActiveControl.Name
    If Me.CommandBars1.actions(Item.ID).Enabled Then
        Call msLeftClick(Item.ID, Me.CommandBars1)
        If Me.CommandBars1.actions(gID.PanePopupMenuNaviAutoFoldOther).Checked Then
            For Each taskGroup In Me.TaskPanel1.Groups
                If taskGroup.ID <> Item.Group.ID Then taskGroup.Expanded = False
            Next
        End If
    Else
        MsgBox "��Ŀǰ��Ȩ�򿪴˲˵�����������ϵ����Ա��", vbExclamation, "Ȩ������"
    End If
    
    Set taskGroup = Nothing
End Sub

Private Sub Timer1_Timer(Index As Integer)
    Const conCon As Byte = 1    '����״̬�����conConn��
    Static byteCon As Byte
    Static byteChk As Byte
    
    '�����ͻ��˳���
    '����Winsock�ؼ��е�gArr��with���������ʱ�޷����gArr���飬Ȩ�˷Ŵ˴�
    If gVar.ClientReLoad Then
        Call msUnloadMe(True)
        Load Me
        Exit Sub
    End If
    
    '�ڵ�½�����е���˹رճ���
    If gVar.UnloadFromLogin And Not gVar.ShowMainWindow Then 'ж�ص�½����+û����ʾ������
        Call msUnloadMe(True)
        Exit Sub
    End If
      
    If gVar.ClientLoginCheckOver And (Not gVar.ShowMainWindow) And gVar.TCPStateConnected Then
        mXtrStatusBar.FindPane(gID.StatusBarPaneUserInfo).Text = gVar.UserFullName '������״̬����ʾ�û�ȫ��
        Me.Show '��ʾ������
        Call gfSendClientInfo(gVar.UserComputerName, gVar.UserLoginName, gVar.UserFullName, Me.Winsock1.Item(1)) '���û���½��Ϣ���͸������
        gVar.ShowMainWindow = True '��ʾ�������־������رճ���ʱ��������״̬
        Call msLoadUserAuthority(gVar.UserAutoID) '����Ȩ��
        
        Dim frmUnload As Form  'ж�ص�½���ڡ���֪Ϊ�Σ�ֱ����Unload frmSysLogin����ж�ص���û��Ӧ��
        For Each frmUnload In Forms
            If LCase(frmUnload.Name) = LCase("frmSysLogin") Then
                Unload frmUnload
                Exit For
            End If
        Next
    End If
    
    byteCon = byteCon + 1 '״̬��ʱ
    If Not gVar.TCPStateConnected And gVar.UpdateRunOver Then byteChk = byteChk + 1  '����ʱ��ֻ��������ʱ
    
    If byteCon >= conCon Then
        If (Not gVar.UpdateRunOver) And (Not gfAppExist(gVar.EXENameOfUpdate)) Then 'Ȩ�����,���жϽ����Ƿ�����ǲ�ȫ���
            gVar.UpdateRunOver = True   '���³�����������ɱ�־
            Call gsOpenTheWindow("frmSysLogin", , vbNormal) ''��ʾ��½����
            gVar.ClientLoginShow = True '����ȫ�ֱ���--�Ѵ򿪹���½����
        End If
             
        With Me.Winsock1.Item(1)
            If .State = 7 Then      '������
                Call msSetClientState(vbGreen)  '��������״̬
            ElseIf .State = 9 Then  '�����쳣
                Call msSetClientState(vbRed)    '�����쳣״̬
            Else                    'δ���ӵ�
                Call msSetClientState(vbYellow) '����δ����״̬
            End If
        End With
        byteCon = 0 '���㾲̬�ۻ�����
    End If
    
    If byteChk > (gVar.TCPWaitTime + 1) Then  '��Ϊ��������Ҳ�ǵȴ�gVar.TCPWaitTime�ŶϿ����ӣ������ӳ�һ��
        byteChk = 0 '��̬��������
        If Not gVar.TCPStateConnected Then  'δ����״̬
            If gVar.ClientLoginCheckOver Then   '��У����˺�����
                Call gsAlarmAndLogEx("���������������ʧ�ܣ���ȷ�Ϸ���˳���������", "���Ӿ�ʾ")
                Call msUnloadMe(True)
            End If
        End If
    End If
    
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '���ӹر�ʱ��մ�����Ϣ
    If UBound(gArr) = 1 Then gArr(1) = gArr(0)
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strGet As String
    Dim byteGet() As Byte
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '�ַ���Ϣ����״̬��
            
            Me.Winsock1.Item(Index).GetData strGet  '�����ַ�
            
            If InStr(strGet, gVar.PTClientConfirm) > 0 Then '�յ�Ҫ�ظ������ȷ�����ӵ���Ϣ
                Call gfSendInfo(gVar.PTClientIsTrue, Me.Winsock1.Item(Index))
                .Connected = True
                
            ElseIf InStr(strGet, gVar.PTConnectIsFull) Then '�յ�����˷���������������
                Me.Timer1.Item(Index).Enabled = False
                Call msUnloadMe(True)
                MsgBox "�ͻ������������������ޣ��������û��˳������ԣ�", vbCritical, "��������������"
                Rem End 'Ӳ���������Է���һ��
            
            ElseIf InStr(strGet, gVar.PTDBDataSource) Then  '�յ����������������ݿ�������Ϣ
                Call gfRestoreDBInfo(strGet) '�������ܹ������ݿ�������Ϣ
                gVar.RestoreDBInfoOver = True '���ݿ�������Ϣ������ɱ�־
                
            ElseIf InStr(strGet, gVar.PTConnectTimeOut) Then '��������ʱ���ѵ�
                Dim blnTimer As Boolean, tmrEn As VB.Timer
                For Each tmrEn In Me.Timer1 '�����ͻ��˺����Ƿ���timer1(1)�����ڣ�Ȩ���ô˼��
                    If tmrEn.Index = Index Then
                        blnTimer = True
                        Exit For
                    End If
                Next
                Set tmrEn = Nothing
                If blnTimer Then Me.Timer1.Item(Index).Enabled = False
                Me.Winsock1.Item(Index).Close
                Call msSetClientState(vbYellow) '����δ����״̬
                MsgBox "���������������ʱ���ѵ��������µ�½��", vbExclamation, "����ʱ��������ʾ"
                gVar.ClientReLoad = True
                If blnTimer Then Me.Timer1.Item(Index).Enabled = True
                
            ElseIf InStr(strGet, gVar.PTFileStart) > 0 Then '���Է����ļ���������˵�״̬
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index)) '�����ļ��������
                Call gsFormEnable(Me, False)    '��ֹ�ͻ����ٲ���
                
            ElseIf InStr(strGet, gVar.PTFileExist) > 0 Then '����˷����ͻ�����Ҫ���ļ�����
                Dim strSize As String, lngInstrSize As Long
                
                lngInstrSize = InStr(strGet, gVar.PTFileSize) '��ȡ�ͻ�����Ҫ�Ĵ����ڷ���˵��ļ��Ĵ�С
                If lngInstrSize > 0 Then
                    strSize = Mid(strGet, lngInstrSize + Len(gVar.PTFileSize))
                    If IsNumeric(strSize) Then
                        .FileSizeTotal = Val(strSize)
                        Call gfSendInfo(gVar.PTFileSend, Me.Winsock1.Item(Index)) '֪ͨ����˿��Է��͹�����
                        '��ʼ��������
                        
                        .FileTransmitState = True
                        Call gsFormEnable(Me, False)    '��ֹ�ͻ����ٲ���
                    End If
                End If
                
            ElseIf InStr(strGet, gVar.PTFileNoExist) > 0 Then   '
                MsgBox "��Ҫ�ļ�<" & .FileName & ">�ڷ���˲����ڣ�", vbExclamation, "�ļ�����"
                gArr(Index) = gArr(0)
                
            End If
            
            Debug.Print "Client GetInfo:" & strGet, bytesTotal
            '�ַ���Ϣ����״̬��
        Else
            '�ļ�����״̬��
            
            If .FileNumber = 0 Then '�����ļ���
                .FileNumber = FreeFile
                Open .FilePath For Binary As #.FileNumber
            End If
            
            ReDim byteGet(bytesTotal - 1)   '�ض��������С
            Me.Winsock1.Item(Index).GetData byteGet, vbArray + vbByte   '�����ļ���Ϣ����������
            Put #.FileNumber, , byteGet '������ļ���
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal    '��¼�Ѵ����С
            '���½�����
            
            If .FileSizeCompleted >= .FileSizeTotal Then    '������ɺ��һЩ����
                Close #.FileNumber
                Call gsFormEnable(Me, True) '����ͻ��˵�����
                gArr(Index) = gArr(0)
                Call gfSendInfo(gVar.PTFileEnd, Me.Winsock1.Item(Index)) '���ͽ�����־
                Debug.Print "Client Received Over"
            End If
            
            '�ļ�����״̬��
        End If
    End With
    
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '�����쳣����
    
    If Index <> 0 Then
        If gArr(Index).FileTransmitState Then   '�쳣ʱ����ļ�������Ϣ
            Debug.Print "ClientWinsockError:" & Index & "--" & Err.Number & "  " & Err.Description
            Close
            gArr(Index) = gArr(0)
            Call gsFormEnable(Me, True)
        End If
        Call gsAlarmAndLogEx("����������ӷ����쳣��", "���Ӿ���", True, vbCritical)
    End If
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    '�����괦��
    
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
