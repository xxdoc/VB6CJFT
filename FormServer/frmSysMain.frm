VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmSysMain 
   Caption         =   "Main�����"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9315
   Icon            =   "frmSysMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9315
   StartUpPosition =   2  '��Ļ����
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   720
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   1440
      Top             =   3840
   End
   Begin FlexCell.Grid Grid1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5106
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   3840
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
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2040
      Top             =   3840
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   2640
      Top             =   3840
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
Dim mcbsPopupIcon As XtremeCommandBars.CommandBar    '����ͼ��Pupup�˵�
Dim CheckConnectTime As Long
Dim ConfirmTime() As Long
Dim ConfirmOK() As Boolean
Dim CountTime() As Long


Private Sub msAddAction(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '����CommandBars��Action
    
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.Actions
    cbsBars.EnableActions   '����CommandBars��Actions����
    
'    cbsActions.Add "Id", "Caption", "TooltipText", "DescriptionText", "Category"   '����
    With cbsActions
        .Add gID.Sys, "ϵͳ", "", "", "ϵͳ"
        
        .Add gID.SysLoginOut, "�˳�", "", "", ""
        .Add gID.SysLoginAgain, "����", "", "", ""
        
        .Add gID.SysExportToCSV, "������CSV", "", "", ""
        .Add gID.SysExportToExcel, "������Excel", "", "", ""
        .Add gID.SysExportToHTML, "������HTML", "", "", ""
        .Add gID.SysExportToPDF, "������PDF", "", "", ""
        .Add gID.SysExportToText, "������txt", "", "", ""
        .Add gID.SysExportToWord, "������Word", "", "", ""
        .Add gID.SysExportToXML, "������XML", "", "", ""
        
        .Add gID.SysPrint, "��ӡ", "", "", ""
        .Add gID.SysPrintPageSet, "��ӡҳ������", "", "", ""
        .Add gID.SysPrintPreview, "��ӡԤ��", "", "", ""
        
        .Add gID.Wnd, "����", "", "", "����"
        
        .Add gID.WndResetLayout, "���ô��ڲ���", "", "", ""
        
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
        
        .Add gID.Help, "����", "", "", "����"
        .Add gID.HelpAbout, "���ڡ�", "", "", ""
        
        .Add gID.StatusBarPane, "״̬��", "", "", ""
        .Add gID.StatusBarPaneProgress, "������", "", "", ""
        .Add gID.StatusBarPaneProgressText, "�������ٷֱ�ֵ", "", "", ""
        .Add gID.StatusBarPaneServerButton, "������/�Ͽ���ť", "", "", ""
        .Add gID.StatusBarPaneServerState, "����״̬", "", "", ""
        .Add gID.StatusBarPaneTime, "ϵͳʱ��", "", "", ""
        .Add gID.StatusBarPaneIP, "����IP��ַ", "", "", ""
        .Add gID.StatusBarPanePort, "�����˿�", "", "", ""
        .Add gID.StatusBarPaneReStartButton, "�����Զ�/�ֶ�����ģʽ�л���ť", "", "", ""
        
        .Add gID.IconPopupMenu, "����ͼ��˵�", "", "", ""
        .Add gID.IconPopupMenuMaxWindow, "��󻯴���", "", "", ""
        .Add gID.IconPopupMenuMinWindow, "��С������", "", "", ""
        .Add gID.IconPopupMenuShowWindow, "��ʾ����", "", "", ""
        
        .Add gID.Tool, "����", "", "", "����"
        .Add gID.toolOptions, "ѡ��", "", "", "frmOption"
        
        
'        .Add gID, "", "", "", ""
        
    End With
    
    '���cbsActions����������ToolTipText��DescriptionText��Key��Category
    For Each cbsAction In cbsActions
        With cbsAction
            If .ID < 20000 Then
                .ToolTipText = .Caption
                .DescriptionText = .ToolTipText
                .Key = .Category    'Ϊ�˵�ʱ�������ã�����Actionʱ������������Category��
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

Private Sub msAddConnectToGrid(ByRef sckAdd As MSWinsockLib.Winsock)
    '���������Ϣ�������
    
    Dim K As Long, C As Long
    
    With Me.Grid1
        .AutoRedraw = False
        C = .Rows - 1
        For K = 1 To C  'Ѱ�ҿ��ÿ���
            If Len(.Cell(K, 1).Text) = 0 Then Exit For
        Next
        If K = C + 1 Then   '����������������
            .Rows = .Rows + 1
        End If
        .Cell(K, 1).Text = sckAdd.RemoteHostIP
        .Cell(K, 2).Text = sckAdd.RemoteHost
        .Cell(K, 5).Text = Format(Now, gVar.Formaty_M_dH_m_s)
        .Cell(K, 6).Text = sckAdd.Index
        .Cell(K, 7).Text = sckAdd.Tag
        .AutoRedraw = True
        .Refresh
    End With
    
End Sub


Private Sub msAddDesignerControls(ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars�Զ���Ի���������������
    
    Dim cbsControls As XtremeCommandBars.CommandBarControls
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
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
    
    Set cbsMenuBar = cbsBars.ActiveMenuBar
    cbsMenuBar.ShowGripper = False  '����ʾ���϶����Ǹ������
    cbsMenuBar.EnableDocking xtpFlagStretched     '�˵�����ռһ���Ҳ��������϶�
    
    'ϵͳ���˵�
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
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    With cbsMenuMain.CommandBar.Controls
        '���ò���
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.WndResetLayout, "")
        cbsMenuCtrl.BeginGroup = True
        
        '����ID35001�Զ��幤����
        Set cbsMenuCtrl = .Add(xtpControlButton, XTP_ID_CUSTOMIZE, "�Զ��幤����...")
        cbsMenuCtrl.BeginGroup = True
    
        '����ID59392�������б�
        Set cbsMenuCtrl = .Add(xtpControlPopup, 0, "�������б�")
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
        
        'CommandBars�����������Ӳ˵�
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeCommandBars, "")
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
                .Add xtpControlButton, mlngID, ""
            Next
        End With
    End With
    
    '���߲˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Tool, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.toolOptions, ""
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Help, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.HelpAbout, ""
    
    Set cbsMenuBar = Nothing
    Set cbsMenuMain = Nothing
    Set cbsMenuCtrl = Nothing
End Sub

Private Sub msAddXtrStatusBar(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '����״̬��
    
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    Dim BarPane As XtremeCommandBars.StatusBarPane
    
    Set cbsActions = cbsBars.Actions
    Set mXtrStatusBar = cbsBars.StatusBar
    With mXtrStatusBar
        .AddPane 0      'ϵͳPane����ʾCommandBarActions��Description
        .SetPaneStyle 0, SBPS_STRETCH
        .SetPaneText 0, "Hello"
        .IdleText = "Hello"
        
        .AddPane gID.StatusBarPaneIP
        .SetPaneText gID.StatusBarPaneIP, Me.Winsock1.Item(0).LocalIP  'gVar.TCPSetIP
        .FindPane(gID.StatusBarPaneIP).Width = 90
        
        .AddPane gID.StatusBarPanePort
        .SetPaneText gID.StatusBarPanePort, gVar.TCPSetPort
        .FindPane(gID.StatusBarPanePort).Width = 60
        
        .AddPane gID.StatusBarPaneReStartButton
        .SetPaneText gID.StatusBarPaneReStartButton, IIf(gVar.ParaBlnAutoReStartServer, "��", "��") & "����������ģʽ"
        .FindPane(gID.StatusBarPaneReStartButton).Width = 120
        .FindPane(gID.StatusBarPaneReStartButton).BackgroundColor = vbCyan
        .FindPane(gID.StatusBarPaneReStartButton).Button = True
        
        .AddPane gID.StatusBarPaneServerState
        .FindPane(gID.StatusBarPaneServerState).Text = gVar.ServerStateNotStarted
        .FindPane(gID.StatusBarPaneServerState).Width = 60
        
        .AddPane gID.StatusBarPaneServerButton
        .FindPane(gID.StatusBarPaneServerButton).Text = gVar.ServerButtonStart
        .FindPane(gID.StatusBarPaneServerButton).Width = 60
        .FindPane(gID.StatusBarPaneServerButton).Button = True
        
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

Private Sub msAddPopupMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '��������ͼ���Ҽ�����ʽ�˵�
        
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
    '����������
    
    Dim cbsBar As XtremeCommandBars.CommandBar
    Dim cbsCtr As XtremeCommandBars.CommandBarControl
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.Actions
    
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
    
    Set cbsBar = Nothing
    Set cbsCtr = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msCloseAllConnect(Optional ByVal blnClose As Boolean = True, _
                              Optional ByVal blnCloseListen As Boolean = True)
    '�ر����пͻ�������
    Dim sckDel As MSWinsockLib.Winsock
    
    If Not blnClose Then Exit Sub   '��ִ�йر�
    
    For Each sckDel In Me.Winsock1
        If sckDel.Index = 0 Then
            If blnCloseListen Then
                sckDel.Close     '�ر�����
            End If
        Else
            If sckDel.State <> 0 Then sckDel.Close  '�ȹر�����
            gArr(sckDel.Index) = gArr(0)    '��ն�Ӧ������Ϣ
            Unload sckDel   'ж�ض�Ӧ�ؼ�
        End If
    Next
            
    With Me.Grid1
        .AutoRedraw = False
        .ReadOnly = False   '��ȡ����������������˱�������ݡ�����д�룿
        .Range(1, 1, .Rows - 1, .Cols - 1).ClearText
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    Debug.Print "msCloseAllConnect(" & blnClose & "," & blnCloseListen & ")"
    Set sckDel = Nothing
End Sub

Private Sub msGetClientInfo(ByVal strInfo As String, ByVal Index As Long)
    '���ͻ��˷�������Ϣ������������û���½ϵͳ�����û�������������
    'ע�⣺�ͻ��˷�����Ϣ˳��̶�Ϊ�������--�û���½ϵͳ��--�û�����
    Dim strPC As String, strLogin As String, strFull As String
    Dim LocPC As Long, LocLogin As Long, LocFull As Long, lngTemp As Long, C As Long
    
    With gVar
        LocPC = InStr(strInfo, .PTClientUserComputerName)
        LocLogin = InStr(strInfo, .PTClientUserLoginName)
        LocFull = InStr(strInfo, .PTClientUserFullName)
        
        If LocPC > 0 And LocLogin > LocPC Then
            lngTemp = LocPC + Len(.PTClientUserComputerName)
            strPC = Mid(strInfo, lngTemp, LocLogin - lngTemp)
        End If
        
        If LocLogin > LocPC And LocFull > LocLogin Then
            lngTemp = LocLogin + Len(.PTClientUserLoginName)
            strLogin = Mid(strInfo, lngTemp, LocFull - lngTemp)
        End If
        
        If LocFull > LocLogin Then
            lngTemp = LocFull + Len(.PTClientUserFullName)
            strFull = Mid(strInfo, lngTemp)
        End If
    End With
    
    With Me.Grid1
        .AutoRedraw = False
        .ReadOnly = False
        C = .Rows - 1
        For lngTemp = 1 To C
            If Val(.Cell(lngTemp, 6).Text) = Index Then
                .Cell(lngTemp, 2).Text = strPC
                .Cell(lngTemp, 3).Text = strLogin
                .Cell(lngTemp, 4).Text = strFull
                If LCase(strLogin) <> LCase(gVar.UpdateAccount) Then
                    Call msWriteLoginInfoLog(.Cell(lngTemp, 1).Text, .Cell(lngTemp, 2).Text, _
                        .Cell(lngTemp, 3).Text, .Cell(lngTemp, 4).Text, .Cell(lngTemp, 5).Text, _
                        .Cell(lngTemp, 6).Text, .Cell(lngTemp, 7).Text)
                End If
                Exit For
            End If
        Next
        .AutoRedraw = True
        .ReadOnly = True
        .Refresh
    End With
End Sub

Private Sub msGridSet(ByRef gridSet As FlexCell.Grid)
    With gridSet
        .AutoRedraw = False
        .Appearance = Flat
        .BackColorBkg = Me.BackColor
        .DisplayRowIndex = True
        .ExtendLastCol = True
        .ReadOnly = True    '��ֹ���༭
        
        .Cols = 9
        .Rows = 2
        .Cell(0, 0).Text = "���"
        .Cell(0, 1).Text = "�����û�IP��ַ"
        .Cell(0, 2).Text = "�����û����������"
        .Cell(0, 3).Text = "�����û���½�˺�"
        .Cell(0, 4).Text = "�����û�����"
        .Cell(0, 5).Text = "���ӽ���ʱ��"
        .Cell(0, 6).Text = "������" '"Index"
        .Cell(0, 7).Text = "�����" '"RequestID"
        .Cell(0, 8).Text = "����ʱ��"
        .Column(1).Width = 120
        .Column(2).Width = 130
        .Column(3).Width = 130
        .Column(4).Width = 120
        .Column(5).Width = 120
        .RowHeight(0) = 40
        .Range(0, 0, 0, .Cols - 1).WrapText = True

        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub msLeftClick(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars����������Ӧ��������
    
    Dim strKey As String
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.Actions
    With gID
        Select Case CID
            Case .WndThemeCommandBarsOffice2000 To .WndThemeCommandBarsWinXP
                Call gsThemeCommandBar(CID, cbsBars)
            Case .WndResetLayout
                Call msResetLayout(cbsBars)
                
            Case .SysLoginAgain
                If MsgBox("ȷ��������������˳�����", vbQuestion + vbOKCancel, "����������ѯ��") = vbOK Then
                    gVar.CloseWindow = True
                    Unload Me
                    Me.Show
                End If
            Case .SysLoginOut
                If MsgBox("ȷ���˳�����˳�����", vbQuestion + vbOKCancel, "�ر�������ѯ��") = vbOK Then
                    gVar.CloseWindow = True
                    Unload Me
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
                
            Case .SysExportToCSV To .SysExportToWord, .SysPrintPageSet To .SysPrint
                If Me.ActiveControl Is Nothing Then Exit Sub
                If Not (TypeOf Me.ActiveControl Is FlexCell.Grid) Then Exit Sub
                If Not cbsActions(CID).Enabled Then Exit Sub
                Select Case CID
                    Case .SysExportToCSV To .SysExportToXML
                        Call gsGridExportTo(Me.ActiveControl, CID)
                    Case .SysExportToText
                        If MsgBox("�Ƿ񽫵�ǰ������ݵ�����txt�ı��ĵ���", vbQuestion + vbYesNo, "ѯ��") = vbYes Then Call gsGridToText(Me.ActiveControl)
                    Case .SysExportToWord
                        If MsgBox("�Ƿ񽫵�ǰ������ݵ�����Word�ĵ���", vbQuestion + vbYesNo, "ѯ��") = vbYes Then Call gsGridToWord(Me.ActiveControl)
                        
                    Case .SysPrint
                        If MsgBox("ȷ����ӡ��ǰ���������", vbQuestion + vbOKCancel, "��ӡѯ��") = vbOK Then Call gsGridPrint(Me.ActiveControl)
                    Case .SysPrintPreview
                        Call gsGridPrintPreview(Me.ActiveControl)
                    Case .SysPrintPageSet
                        Call gsGridPageSet(Me.ActiveControl)
                End Select
                
            Case Else
                strKey = LCase(cbsActions.Action(CID).Key)
                If Left(strKey, 3) = "frm" Then
                    If cbsActions.Action(CID).Enabled Then
'''                        Select Case strKey
'''                            Case LCase(cbsActions(gID.toolOptions).Key)
'''                                Call gsOpenTheWindow(strKey, vbModal, vbNormal)
'''                            Case Else
'''                                Call gsOpenTheWindow(strKey)
'''                        End Select
                        Select Case CID
                            Case .toolOptions
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
    
    On Error Resume Next    '��/���ܺ������̿������쳣
    With gVar
        .ParaBlnWindowCloseMin = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, 1))    '�ر�ʱ��С��
        .ParaBlnWindowMinHide = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, 1))  '��С��ʱ����
        
        .TCPDefaultIP = Me.Winsock1.Item(0).LocalIP '����IP��ַ
        .TCPSetIP = gVar.TCPDefaultIP   '�����ʹ�ñ���IP��ַ
        .ParaBlnAutoReStartServer = Val(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyParaAutoReStartServer, 1))   '�ֶ�/�Զ���������ģʽ
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , gVar.TCPDefaultPort, 10000, 65535) '�����˿�
        .ParaBlnAutoStartupAtBoot = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, 0))  '�����Զ�����
        
        .ConSource = gfCheckIP(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .TCPSetIP))   '����������/IP
        .ConDatabase = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString("dbTest", .EncryptKey)), .EncryptKey)    '���ݿ���
        .ConUserID = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString("123", .EncryptKey)), .EncryptKey)  '��½��
        .ConPassword = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString("888888", .EncryptKey)), .EncryptKey)    '��½����
        
        .ParaBlnLimitClientConnect = Val(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnect, 0)) '���ƿͻ�������
        .ParaLimitClientConnectTime = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectTime, True, 30, 1, 60) '���ƿͻ�������ʱ��
        .TCPConnectMax = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectNumber, True, 2, 1) '���ƿͻ���������
        
    End With
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

Private Sub msSetServerState(ByVal colorSet As Long)
    '����״̬���з���˵�״̬
    
    Dim paneState As XtremeCommandBars.StatusBarPane
    Dim paneButton As XtremeCommandBars.StatusBarPane
    
    Set paneState = mXtrStatusBar.FindPane(gID.StatusBarPaneServerState)
    Set paneButton = mXtrStatusBar.FindPane(gID.StatusBarPaneServerButton)
    If colorSet = vbGreen Then
        paneState.BackgroundColor = vbGreen
        paneState.Text = gVar.ServerStateStarted
        paneButton.Text = gVar.ServerButtonClose
        paneButton.TextColor = vbMagenta 'vbRed
    ElseIf colorSet = vbRed Then
        paneState.BackgroundColor = vbRed
        paneState.Text = gVar.ServerStateError
        paneButton.Text = gVar.ServerButtonStart
        paneButton.TextColor = vbBlue ' vbGreen
    Else
        paneState.BackgroundColor = vbYellow
        paneState.Text = gVar.ServerStateNotStarted
        paneButton.Text = gVar.ServerButtonStart
        paneButton.TextColor = vbBlue ' vbGreen
    End If
    
    Set paneState = Nothing
    Set paneButton = Nothing
End Sub

Private Sub msStartConfirm(ByVal Index As Integer)
    '���ǿͻ��������ӷ�������������Ӧ��ʱ�����з������
    Dim tmrConfirm As VB.Timer
    Dim blnExist As Boolean
    
    If Index = 0 Then
        MsgBox "��ʱ��Indexֵ�Ƿ����룡", vbCritical, "�����⽨�����Ӿ���"
        Exit Sub
    End If
    
    For Each tmrConfirm In Me.Timer1    '�����⣬����Ƿ��Ѵ��ڸü�ʱ��
        If tmrConfirm.Index = Index Then
            blnExist = True
            Exit For
        End If
    Next
    
    With Me.Timer1
        If Not blnExist Then Load .Item(Index) '������ָ��Index�Ŀؼ�ʱ�ż���
        .Item(Index).Interval = 1000    '��ʱ�����
        .Item(Index).Enabled = True '�����ʱ
    End With
    
    Set tmrConfirm = Nothing
End Sub

Private Sub msStartServer(ByRef sckCon As MSWinsockLib.Winsock)
    '��������
    
    With sckCon
        If .State <> 0 Then .Close  '�ȹر�
        .LocalPort = gVar.TCPSetPort
        .Listen
    End With
End Sub

Private Sub msVersionCS(ByVal strVer As String, ByRef sckVer As MSWinsockLib.Winsock)
    '����ͻ��˷����İ汾��Ϣ
    Dim strVC As String, strVS As String, strCompare As String
    Dim strNetFile As String, strSetupFile As String
    
    strNetFile = gVar.AppPath & gVar.EXENameOfClient    '�����ڷ���˵ĶԱ��õĿͻ������exe�ļ�
    strVS = gfBackVersion(strNetFile)   '��ȡ����˶Ա��õĿͻ��˰汾��
    strVC = Mid(strVer, Len(gVar.PTVersionOfClient) + 1)    '��ȡ�ͻ��˷����İ汾��
    
    strCompare = gfVersionCompare(strVC, strVS) '�Ƚ� �ͻ��˷����� �����˶Ա��õ� �汾��
    If strCompare = "0" Then    'û���°汾�����ø���
        Call gfSendInfo(gVar.PTVersionNotUpdate, sckVer)
    ElseIf strCompare = "1" Then    '���°棬Ҫ����
        Call gfSendInfo(gVar.PTVersionNeedUpdate & strVS, sckVer)   '֪ͨ�ͻ�����Ҫ����
        
        '��֯�����͸����ļ���װ�������Ϣ���ͻ���
        gArr(sckVer.Index) = gArr(0)
        strSetupFile = gVar.AppPath & gVar.EXENameOfSetup
        If Not gfFileExist(strSetupFile) Then
            Call gsAlarmAndLog("Ҫ���͵ĸ��³��򲻴���", False)
            Exit Sub
        End If
        With gArr(sckVer.Index)
            .FileFolder = "" ' gVar.FolderTemp
            .FileName = gVar.EXENameOfSetup
            .FilePath = strSetupFile
            .FileSizeTotal = FileLen(.FilePath)
        End With
        If sckVer.State = 7 Then
            If gfSendInfo(gfFileInfoJoin(sckVer.Index, ftSend), sckVer) Then '�ȷ����ļ�����Ϣ
                Debug.Print "Server:�ѷ��͸��³�����ļ���Ϣ"
            End If
        End If
        
    Else    '�汾����쳣����
        Call gfSendInfo(gVar.PTVersionNotUpdate & strCompare, sckVer)
        Call gsAlarmAndLogEx("�ͻ��˰汾��" & strVC & ",����˰汾��" & strVS, "Server�˰汾����쳣", False)
    End If
End Sub

Private Sub msWriteLoginInfoLog(ByVal strIP As String, ByVal strPC As String, ByVal strAccount As String, _
    ByVal strUserName As String, ByVal strTime As String, ByVal strIndex As String, ByVal strApplyID As String)
    '��¼�û��ĵ�½��־���ļ���������gVar.FileNameLoginLog��
    
    Const conSize As Long = 1000000 '�̶���־�ļ���С�����������ڴ洢
    Dim strNewFile As String
    Dim intNum As Integer
    
    If Not gfFileRepair(gVar.FolderData, True) Then Exit Sub    '�ж���־Ŀ¼�Ƿ����
    If Not gfFileRepair(gVar.FileNameLoginLog) Then Exit Sub    '�ж���־�ļ��Ƿ����
    
    If FileLen(gVar.FileNameLoginLog) > conSize Then    '��־�ļ�̫��ʱ�浵
        strNewFile = Left(gVar.FileNameLoginLog, InStrRev(gVar.FileNameLoginLog, ".") - 1) & _
            Format(Now, gVar.Formatymdhms) & Mid(gVar.FileNameLoginLog, InStrRev(gVar.FileNameLoginLog, "."))
        Debug.Print strNewFile  '���ɰ����ڱ�����ļ���
        Close
        If Not gfFileRename(gVar.FileNameLoginLog, strNewFile) Then Exit Sub    '�����洢
    End If
    
    intNum = FreeFile
    On Error Resume Next
    
    Open gVar.FileNameLoginLog For Append As intNum
    Print #intNum, strIP & vbTab & strPC & vbTab & strAccount & _
        vbTab & strUserName & vbTab & strTime & vbTab & strIndex & vbTab & strApplyID
    Close
    
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '������¼�
    Call msLeftClick(Control.ID, Me.CommandBars1)
End Sub

Private Sub CommandBars1_Resize()
    '�������ڲ���
    
    Dim L As Long, T As Long, R As Long, b As Long
    
    On Error Resume Next
    Me.CommandBars1.GetClientRect L, T, R, b
    Grid1.Move L, T, R - L, b - T
    
End Sub

Private Sub CommandBars1_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'CommandBars�ؼ���Action״̬���л�
    
    Dim blnFC As Boolean
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = Me.CommandBars1.Actions
    If Screen.ActiveControl Is Nothing Then
        blnFC = False
    Else
        blnFC = TypeOf Screen.ActiveControl Is FlexCell.Grid
    End If
    With gID
        For mlngID = .SysExportToCSV To .SysExportToWord
            cbsActions(mlngID).Enabled = blnFC
        Next
        For mlngID = .SysPrintPageSet To .SysPrint
            cbsActions(mlngID).Enabled = blnFC
        Next
    End With
End Sub

Private Sub Form_Load()
    '�������
    
    Dim cbsBars As XtremeCommandBars.CommandBars
    
    '�򿪶��Ӧ�ó�����
    If App.PrevInstance Then
        MsgBox "����ͬʱ�򿪶��Ӧ�ó���", vbCritical, "����"
        End
    End If
    
    Timer1.Item(0).Interval = 1000  '��ʱ��ѭ��ʱ��
    Call Main   '��ʼ��ȫ�ֹ��ñ���
    Set gWind = Me  'ָ���������ȫ�����ö���
    XtremeCommandBars.CommandBarsGlobalSettings.App = App 'һ��Ĭ������
    Set cbsBars = Me.CommandBars1
    
    Call msLoadParameter(True)  '�������ò���
    Call msAddAction(cbsBars)   '����Actions����
    Call msAddMenu(cbsBars)     '�����˵���
    Call msAddToolBar(cbsBars)  '����������
    Call msAddPopupMenu(cbsBars)    '��������ͼ��Ĳ˵�
    Call msAddXtrStatusBar(cbsBars) '����״̬��
    Call msAddKeyBindings(cbsBars)  '��ӿ�ݼ�,�ŵ�LoadCommandBars�������������Ч������
    Call msAddDesignerControls(cbsBars) 'CommandBars�Զ���Ի�����ʹ�õ�
    
    cbsBars.AddImageList ImageList1         'ʹCommandBars�ؼ�ƥ��ImageList�ؼ���ͼ��
    cbsBars.EnableCustomization True        '����CommandBars�Զ��壬��������÷�������CommandBars�趨֮��
    cbsBars.Options.UpdatePeriod = 250      '����CommandBars��Update�¼���ִ�����ڣ�Ĭ��100ms
    
    Call gsLoadSkin(Me, Me.SkinFramework1, sMSO7, True)  '���ش�������
    
    '���ع���������
    Call gsThemeCommandBar(Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerCommandbarsTheme, gID.WndThemeCommandBarsRibbon)), cbsBars)
    
    'ע�����Ϣ����-CommandBars����
    Call cbsBars.LoadCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSServerSetting)

    Call gsFormSizeLoad(Me) 'ע�����Ϣ����-����λ�ô�С
    
    
    
    
    '����Ƿ�Ϊ���ð�******************************
    
    
    Call msGridSet(Grid1)  '�������
    Call gfNotifyIconAdd(Me)    '�������ͼ��
    
    Set cbsBars = Nothing   '����ʹ����Ķ���
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '��Ӧ����ͼ��Ĳ˵�
    Dim sngMsg As Single
    
    If Y <> 0 Then Exit Sub    '�ƺ��˾������ס���һ����������ͼ���ϣ������ڴ�����
    sngMsg = X / Screen.TwipsPerPixelX
    Select Case sngMsg
        Case WM_RBUTTONUP
            mcbsPopupIcon.ShowPopup  '�Ҽ�����Popup�˵�

        Case WM_LBUTTONDBLCLK   '���˫������ͼ��ʱ ��������ʾ/��С�� �л�
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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

Private Sub Form_Resize()
    '������С����ʾ
    If Me.Visible And Me.WindowState = vbMinimized Then
        If gVar.ParaBlnWindowMinHide Then
            Me.Hide
            Call gfNotifyIconBalloon(Me, "��С����ϵͳ����ͼ����", "��С����ʾ")
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж�ش���ʱ������Ϣ
    Dim resetNotifyIconData As gtypeNOTIFYICONDATA
    
    '����ע�����Ϣ-CommandBars����
    Call Me.CommandBars1.SaveCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSServerSetting)
    
    Call gsFormSizeSave(Me) '����ע�����Ϣ-����λ�ô�С
    Call gsSaveCommandbarsTheme(Me.CommandBars1)   '����CommandBars�ķ������
    
    gVar.CloseWindow = False    '����رմ���״̬
    Call SkinFramework1.LoadSkin("", "")    '���Ƥ��
    Set mXtrStatusBar = Nothing  '���״̬��
    Set mcbsPopupIcon = Nothing '���Popup�˵�
    Call gfNotifyIconDelete(Me) 'ɾ������ͼ��
    gNotifyIconData = resetNotifyIconData   '�������������Ϣ��������������ʱ���Զ�����������ֻ�ܷ��Ͼ�ɾ������ͼ�����ĺ���?
    ReDim gArr(0)
    Set gWind = Nothing '���ȫ�ִ�������
    
End Sub

Private Sub mXtrStatusBar_PaneClick(ByVal Pane As XtremeCommandBars.StatusBarPane)
    '״̬����ť�¼�
    Dim strMsg As String
    
    If Pane.ID = gID.StatusBarPaneServerButton Then '�Ͽ�/��������
        If Pane.Text = gVar.ServerButtonClose Then strMsg = "�رպ��Ͽ������û������ӡ�"
        If MsgBox("�Ƿ�" & Pane.Text & "��" & strMsg, vbQuestion + vbYesNo, "����/�Ͽ�����ѯ��") = vbNo Then Exit Sub
        If Pane.Text = gVar.ServerButtonClose Then     '�رշ���
            Pane.Text = gVar.ServerButtonStart
            Call msCloseAllConnect(True, True)
        ElseIf Pane.Text = gVar.ServerButtonStart Then     '��������
            Pane.Text = gVar.ServerButtonClose
            Call msStartServer(Me.Winsock1.Item(0))
        End If
        
    ElseIf Pane.ID = gID.StatusBarPaneReStartButton Then    '�ֶ�/�Զ���������ģʽ
        strMsg = "�Ƿ��л���" & IIf(gVar.ParaBlnAutoReStartServer, "��", "��") & "����������ģʽ��"
        If MsgBox(strMsg, vbQuestion + vbYesNo, "ģʽ�л�ѯ��") = vbYes Then
            gVar.ParaBlnAutoReStartServer = Not gVar.ParaBlnAutoReStartServer
            mXtrStatusBar.FindPane(gID.StatusBarPaneReStartButton).Text = IIf(gVar.ParaBlnAutoReStartServer, "��", "��") & "����������ģʽ"
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionTCP, gVar.RegKeyParaAutoReStartServer, IIf(gVar.ParaBlnAutoReStartServer, 1, 0))
        End If
        
    End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
    'Index=0�ļ�ʱ�����1�롣Timer1��Indexֵ �� Winsock1��Index��Ӧ
    
    Dim sckClose As MSWinsockLib.Winsock, sckCheck As MSWinsockLib.Winsock, tmrUld As VB.Timer
    Dim timeOut As Long, lngRows As Long
'    Static CheckConnectTime As Long
'    Static ConfirmTime() As Long
'    Static ConfirmOK() As Boolean
'    Static CountTime() As Long
        
'''    On Error Resume Next
    If Index = 0 Then
        If Me.Winsock1.Item(Index).State = 2 Then  '��������״̬
            Call msSetServerState(vbGreen)
        Else    '����״̬
            If Me.Winsock1.Item(Index).State = 9 Then  '�쳣״̬
                Call msSetServerState(vbRed)
            Else    '�رյ�
                Call msSetServerState(vbYellow)
            End If
            If gVar.ParaBlnAutoReStartServer Then   '����ѡ���Զ�������������������
                Call msStartServer(Me.Winsock1.Item(0))
            End If
        End If
        
        '���ͻ��˷������ر�ʱ�����Ӳ����Զ��Ͽ����˴�ÿ��һ��ʱ����һ���������ӵ�״̬��������7��رյ�����
        CheckConnectTime = CheckConnectTime + 1
        If CheckConnectTime > 5 Then    'ÿ��N����һ��
            For Each sckCheck In Me.Winsock1
                If sckCheck.Index <> 0 Then
                    If sckCheck.State <> 7 Then
                        For Each tmrUld In Me.Timer1
                            If tmrUld.Index = sckCheck.Index Then
                                Unload tmrUld 'Me.Timer1.Item(sckCheck.Index)
                                Exit For
                            End If
                        Next
                        Debug.Print "CheckConnect:" & sckCheck.Index
                        Call Winsock1_Close(sckCheck.Index)
                    End If
                End If
            Next
            CheckConnectTime = 0
        End If
        '''index=0��ʱ��Ϊ�����������
        
        'ˢ��ÿ�����ӵ�ʱ��
        With Me.Grid1
            lngRows = .Rows - 1
            If lngRows > 0 Then
                For mlngID = 1 To lngRows
                    If Len(.Cell(mlngID, 1).Text) = 0 Or Not IsDate(.Cell(mlngID, 5).Text) Then
                        Exit For
                    Else
                        .Cell(mlngID, 8).Text = Format(Now - CDate(.Cell(mlngID, 5).Text), "HH:mm:ss")
                    End If
                Next
            End If
        End With
    Else
        '''index>0Ϊ�����ͻ���������
        
        timeOut = gVar.ParaLimitClientConnectTime * 60  '������������ʱ����������
'''        ReDim Preserve ConfirmTime(Me.Timer1.UBound)    '��Ҫÿ�ζ�������
'''        ReDim Preserve ConfirmOK(Me.Timer1.UBound)
'''        ReDim Preserve CountTime(Me.Timer1.UBound)
        
        If Not ConfirmOK(Index) Then ConfirmTime(Index) = ConfirmTime(Index) + 1
        If gVar.ParaBlnLimitClientConnect Then CountTime(Index) = CountTime(Index) + 1
        
        If ConfirmTime(Index) > gVar.TCPWaitTime Then   'ȷ���ǿͻ��˷���������
            If Not gArr(Index).Connected Then   '���ǿͻ�����ر�
                For Each sckClose In Me.Winsock1
                    If sckClose.Index = Index Then
                        Call Winsock1_Close(Index) '�ر�����
                        Exit For
                    End If
                Next
            End If
            ConfirmTime(Index) = 0  'ȷ�ϼ�ʱ������
            ConfirmOK(Index) = True 'ȷ�ϱ�־
            If Not gVar.ParaBlnLimitClientConnect Then  '��û��ѡ�����ƿͻ������ӹ���
                ConfirmOK(Index) = False
                CountTime(Index) = 0    '�������Ӽ�ʱ������
                For Each tmrUld In Me.Timer1
                    If tmrUld.Index = Index Then
                        Unload tmrUld   'Unload Me.Timer1.Item(Index) 'ȷ�����ж�ص���Ӧ��ʱ���ؼ�
                        Exit For
                    End If
                Next
            End If
        End If
            
        If CountTime(Index) > timeOut Then   '��ʱ������رն�Ӧ�ͻ�������
            '�����ļ�����״̬����ȴ�����2���Ӻ�ֱ�ӹر�
            If (Not gArr(Index).FileTransmitState) Or (CountTime(Index) - timeOut > 120) Then
                CountTime(Index) = 0    '��ռ�ʱ
                If gArr(Index).FileTransmitState Then   '����ļ�������Ϣ
                    Close
                    gArr(Index) = gArr(0)
                End If
                Unload Me.Timer1.Item(Index)
                For Each sckCheck In Me.Winsock1
                    If sckCheck.Index = Index Then 'ȷ�Ͽؼ��Ƿ����
                        If sckCheck.State = 7 Then 'ȷ�Ͽؼ�״̬�Ƿ��
                            Call gfSendInfo(gVar.PTConnectTimeOut, Me.Winsock1.Item(Index)) '��������ʱ���ѵ���Ϣ���ͻ���
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        
    End If
    
    Set sckClose = Nothing
    Set sckCheck = Nothing
    Set tmrUld = Nothing
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '�ر�����
    Dim K As Long, C As Long
    Dim strIP As String, strRequestID As String
    Dim tmDel As VB.Timer
    
'''    On Error Resume Next
    If Index = 0 Then
        Call msCloseAllConnect(True, False)  '�ر������ؼ���ر���������
    Else
        With Me.Grid1
            .AutoRedraw = False
            .ReadOnly = False   '��ȡ������
            C = .Rows - 1
            For K = 1 To C
                strIP = Trim(.Cell(K, 1).Text)
                strRequestID = Trim(.Cell(K, 7).Text)
                If Len(strIP) > 0 Then
                    If (strIP = Me.Winsock1.Item(Index).RemoteHostIP) And _
                            (strRequestID = Me.Winsock1.Item(Index).Tag) Then   '�п���ͬһIP��½����ͻ��ˣ�����RequestID����
                        .RemoveItem K   '�Ƴ���Ӧ����Ϣ
                        .AddItem ""     'ĩβ���һ��ά�ֱ����������
                        Debug.Print K & ",Winsock_Close:" & Index, Me.Winsock1.Item(Index).Tag, Me.Winsock1.Item(Index).RemoteHostIP
                        For Each tmDel In Me.Timer1 'ж�ض�Ӧ��ʱ��
                            If tmDel.Index <> 0 Then
                                If tmDel.Index = Index Then 'ȷ�Ͽؼ��Ƿ����
                                    Unload tmDel
                                    If Index <= UBound(CountTime) Then CountTime(Index) = 0
                                    Exit For
                                End If
                            End If
                        Next
                        Unload Me.Winsock1.Item(Index)  'ж�ضϿ��Ŀͻ��˵����ӿؼ�
                        gArr(Index) = gArr(0)   '�������
                        Close   '�ر����д򿪵��ļ�
                        Exit For
                    End If
                End If
            Next
            .AutoRedraw = True
            .ReadOnly = True
            .Refresh
        End With
    End If
    
    Set tmDel = Nothing
    
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    '��������
    
    Dim sckNew As MSWinsockLib.Winsock
    Dim K As Long
    Dim blnFull As Boolean
    
    If Index <> 0 Then Exit Sub '��0Ԫ�صĿؼ��������ܽ�����������
    
    '����Winsock1�п��õ���С��Indexֵ
    For Each sckNew In Winsock1
        If sckNew.Index = K Then
            K = K + 1
        Else    '˵��Index�ϺŲ�������
            Exit For
        End If
    Next
    Set sckNew = Nothing
    
    If Me.Winsock1.Count > gVar.TCPConnectMax Then
        Call gsAlarmAndLog("�ͻ��˷��͵������������ѳ������ֵ" & CStr(gVar.TCPConnectMax), False)
        blnFull = True
    End If
    
    With Me.Winsock1
        If K = .Count Then ReDim Preserve gArr(K)   '���������С
        gArr(K) = gArr(0)   '��ʼ������Ԫ��K�����ܱ�֮ǰ������������ʹ�ù�
        
        Load .Item(K)   '����Winsock1(K)�ؼ�
        .Item(K).Accept requestID   'ָ���ͻ�����������Ӹ������ɵ�Winsock1(K)�ؼ�
        .Item(K).Tag = requestID    '�洢������ţ���������
        Debug.Print "Load winsock1(" & K & ")"
        Call msAddConnectToGrid(.Item(K)) '��������ӽ������
        
        ReDim Preserve ConfirmTime(.UBound) '����Ԫ��ֵ�ı仯��Timer1�ؼ��У���ʼ��Ȩ�ҷŴ�
        ReDim Preserve ConfirmOK(.UBound)
        ReDim Preserve CountTime(.UBound)
    End With
    
    If blnFull Then '������������������֪ͨ�ͻ��ˡ��ÿͻ����Զ��ر����ӣ��ͻ���ͬʱ�����ر��¼�.
        Call gfSendInfo(gVar.PTConnectIsFull, Me.Winsock1.Item(K))
        Exit Sub
    End If
    
    Call gfSendInfo(gVar.PTClientConfirm, Me.Winsock1.Item(K)) '���Ϳͻ���ȷ����Ϣ�����涨ʱ���ڷ���ȷ����Ϣ����������������Ͽ����ӡ�
    Call msStartConfirm(K)  '���� ����ȷ����Ϣ ��ʱ��
    
    
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    '������Ϣ
    Dim strGet As String
    Dim byteGet() As Byte
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '�ַ���Ϣ����״̬
            
            Me.Winsock1.Item(Index).GetData strGet  '�ȿؼ�������Ϣ
            
            Call gfRestoreInfo(strGet, Me.Winsock1.Item(Index))  '��ȡ�������ǲ����ļ���Ϣ
            
            If InStr(strGet, gVar.PTClientIsTrue) Then '�ͻ��˷��ص�����ȷ����Ϣ
                .Connected = True   '����״̬�Թ���ʱ���ж�
                Call gfSendInfo(gfDatabaseInfoJoin(True), Me.Winsock1.Item(Index))
                
            ElseIf InStr(strGet, gVar.PTVersionOfClient) > 0 Then '���յ��ͻ��˰汾��Ϣ
                Call msVersionCS(strGet, Me.Winsock1.Item(Index))
            
            ElseIf InStr(strGet, gVar.PTClientUserComputerName) > 0 Then '�ͻ��˷����ļ���������û�������Ϣ
                Call msGetClientInfo(strGet, Index)
                
            ElseIf InStr(strGet, gVar.PTFileStart) > 0 Then 'Ҫ��ʼ�����ļ����ͻ���
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index))
                
            End If
            Debug.Print "Server GetInfo:" & strGet, bytesTotal
            
        '�ַ���Ϣ����״̬
        Else
        '�ļ�����״̬
            
            If .FileNumber = 0 Then
                .FileNumber = FreeFile '���ɲ����ļ���
                Open .FilePath For Binary As #.FileNumber '�Զ����Ʒ�ʽ���ļ�
            End If
            
            ReDim byteGet(bytesTotal - 1)   'ȷ���ֽ������С
            Me.Winsock1.Item(Index).GetData byteGet, vbArray + vbByte    '�����ļ�
            Put #.FileNumber, , byteGet 'д���ļ�
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal    'ͳ���Ѵ����ļ���С������
            
            If .FileSizeCompleted >= .FileSizeTotal Then    '���������
                Close #.FileNumber  '�رղ����ļ���
                Call gfSendInfo(gVar.PTFileEnd, Me.Winsock1.Item(Index)) '���ͽ�������ź�
                gArr(Index) = gArr(0)   '����ļ���Ϣ
                Debug.Print "Server Received Over"
            End If
            
        End If
    End With
    
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '�쳣����
    
    If Index = 0 Then
        Call msCloseAllConnect(True, True)  '�����ؼ��쳣��ر��������ӡ�����ûʲô��
    Else
        If gArr(Index).FileTransmitState Then   '�쳣ʱ����ļ�������Ϣ
            Close
            gArr(Index) = gArr(0)
        End If
    End If
    Debug.Print "ServerWinsockError:" & Index & "--" & Err.Number & "  " & Err.Description
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    '�����괦��
    
    If Index = 0 Then Exit Sub  '0Ԫ��ֻ��������������
    With gArr(Index)
        If .FileTransmitState Then
            If .FileSizeCompleted < .FileSizeTotal Then 'δ���������������
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index))
            Else    '����������մ�����Ϣ
                gArr(Index) = gArr(0)
                Debug.Print "Server Send File Over"
            End If
        End If
    End With
End Sub
