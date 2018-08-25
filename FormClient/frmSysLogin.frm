VERSION 5.00
Begin VB.Form frmSysLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   Icon            =   "frmSysLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5295
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "进入"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   750
      TabIndex        =   3
      Top             =   1380
      Width           =   800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   750
      TabIndex        =   2
      Top             =   800
      Width           =   800
   End
End
Attribute VB_Name = "frmSysLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Load gWind
    gWind.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
    Unload gWind
End Sub
