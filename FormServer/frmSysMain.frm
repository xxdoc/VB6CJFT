VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmSysMain 
   Caption         =   "Main"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8910
   StartUpPosition =   2  '屏幕中心
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   720
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   3840
   End
   Begin FlexCell.Grid Grid1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4895
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   960
      Top             =   240
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   360
      Top             =   240
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

Private Sub Form_Load()
    '窗体加载
    Grid1.BackColorBkg = Me.BackColor
    Call Main   '初始化变量
        
End Sub

