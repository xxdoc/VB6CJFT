VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSysLog 
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   13530
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar Hsb 
      Height          =   255
      Left            =   11760
      TabIndex        =   25
      Top             =   5760
      Width           =   1455
   End
   Begin VB.VScrollBar Vsb 
      Height          =   1935
      Left            =   12840
      TabIndex        =   24
      Top             =   3720
      Width           =   255
   End
   Begin VB.Frame ctlMove 
      Caption         =   "Frame3"
      Height          =   6495
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin FlexCell.Grid Grid1 
         Height          =   2055
         Left            =   1320
         TabIndex        =   26
         Top             =   2400
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3625
         Cols            =   5
         GridColor       =   12632256
         Rows            =   30
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   615
         Left            =   1320
         TabIndex        =   15
         Top             =   5880
         Width           =   8175
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   0
            Left            =   7320
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   2
            Left            =   1440
            TabIndex        =   19
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   3
            Left            =   2520
            TabIndex        =   18
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   4
            Left            =   3480
            TabIndex        =   17
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   270
            Left            =   6240
            TabIndex        =   16
            Text            =   "Text2"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   4560
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   5520
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9855
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   720
            TabIndex        =   8
            Text            =   "Combo1"
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   0
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   1
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   7680
            TabIndex        =   4
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   375
            Left            =   8640
            TabIndex        =   3
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   4560
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   840
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   5
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   89128961
            CurrentDate     =   42628
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Index           =   1
            Left            =   6000
            TabIndex        =   10
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   89128961
            CurrentDate     =   42628
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   11
            Top             =   840
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmSysLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim lngPageSize As Long
Dim lngPageCount As Long
Dim lngPageCur As Long
Dim lngAddWith As Long
Dim rsLog As New ADODB.Recordset

Private Type typeInitialSize
    frmWidth As Long
    frmHeight As Long
    vsWidth As Long
    vsHeight As Long
    frameLeft As Long
    frameTop As Long
    rowHeight As Long
    pageSize As Long
End Type
Dim lngSize As typeInitialSize

Dim strLastTxt As String    '���浥Ԫ��༭֮ǰֵ

Private Sub Check1_Click()
    'ʱ��
    Check1.ForeColor = IIf(Check1.Value, vbBlue, vbRed)
    DTPicker1.Item(0).Enabled = IIf(Check1.Value, True, False)
    DTPicker1.Item(1).Enabled = DTPicker1.Item(0).Enabled
    
End Sub


Private Sub Combo2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '������
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        Combo2.Item(Index).ListIndex = -1
    End If
    
End Sub

Private Sub Command1_Click()
    '''��ѯ

    Dim strSQL As String
    Dim strMen As String
    Dim strClass As String
    Dim strDateA As String
    Dim strDateB As String
    Dim strInfo As String
    Dim strCK As String
    
    strMen = Trim(Combo1.Text)
    strCK = gfStringCheck(strMen)
    If Len(strCK) > 0 Then
        MsgBox Label1(0).Caption & "�в��ܰ����ַ���" & strCK & "����", vbExclamation, "�����ַ�����"
        Combo1.SetFocus
        Exit Sub
    End If
    
    strClass = Trim(Combo2.Item(0).Text)

    strInfo = Trim(Text1.Text)
    strCK = gfStringCheck(strInfo)
    If Len(strCK) > 0 Then
        MsgBox Label1(3).Caption & "�в��ܰ����ַ���" & strCK & "����", vbExclamation, "�����ַ�����"
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
    
    If Check1.Value Then
        strDateA = Format(DTPicker1.Item(0).Value, "yyyy-MM-dd hh:mm:ss")
        strDateB = Format(DTPicker1.Item(1).Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSQL = "EXEC sp_FT_Sys_LogQuery '" & strClass & "','" & strInfo & "','" & _
            strDateA & "','" & strDateB & "','','" & strMen & "'"

    Set rsLog = gfBackRecordset(strSQL)
    If rsLog.State = adStateOpen Then
        If rsLog.RecordCount = 0 Then MsgBox "û�з������������ݣ�", vbExclamation, "��ֵ����"
    End If
    lngPageCur = 1
    
    Call msShowValue

End Sub
    
Private Sub Command2_Click()
    '''�˳�
    Unload Me
    
End Sub

Private Sub Command3_Click(Index As Integer)
    '''��ҳ
    Dim C As Long
    
    Select Case Index
        Case 1
            lngPageCur = 1
        Case 2
            lngPageCur = lngPageCur - 1
        Case 3
            lngPageCur = lngPageCur + 1
        Case 4
            lngPageCur = lngPageCount
        Case 0
            C = Val(Text2.Text)
            If C < 1 Then C = 1
            If C > lngPageCount Then C = lngPageCount
            lngPageCur = C
'            Text2.Text = CStr(c)
        Case Else
            Exit Sub
    End Select
    
    Call msShowValue
    
End Sub

Private Sub Form_Load()
    '���������
    Dim intAli As Integer
    Dim lngColor As Long

    With lngSize
        .frmWidth = 15800   '''��ʼ�������ߴ�
        .frmHeight = 8900
        .vsWidth = 13800
        .vsHeight = 6000
        .rowHeight = 18
        .pageSize = 20
    End With
    
    intAli = 1
    lngPageSize = lngSize.pageSize
    lngColor = vbBlue
    
    Me.Icon = gWind.ImageList1.ListImages("SysLog").Picture
    Me.Caption = gWind.CommandBars1.Actions(gID.SysAuthLog).Caption

    With Frame1 '��ѯ����
        .Move 120, 120, lngSize.vsWidth, 1200
        .Caption = "ѡ���������������"
        .ForeColor = vbMagenta
        
        Label1.Item(0).Move 120, 300, 900, 255
        
        Grid1.Move .Left, (.Top + .Height + 120), .Width, lngSize.vsHeight
        
    End With
    
    With Label1.Item(0)
        .Caption = "�����û�"
        .Alignment = intAli
        .ForeColor = lngColor
        
        Combo1.Move (.Left + .Width + 50), .Top - 30, .Width * 1.5
        
        Check1.Move (Combo1.Left + Combo1.Width + 500), .Top, .Width, .Height
        Check1.Caption = "ʱ���"
        Check1.Value = 1
        
        DTPicker1.Item(0).Move (Check1.Left + Check1.Width), .Top, 1300, .Height

        Label1.Item(2).Move .Left, (.Top + .Height + 200), .Width, .Height

    End With
    
    With DTPicker1.Item(0)
        .CustomFormat = "yyyy-MM-dd"
        .Format = dtpCustom
        .Value = Date
        
        Label1.Item(1).Caption = "--"
        Label1.Item(1).Move (.Left + .Width), .Top, 200, .Height
        
        DTPicker1.Item(1).Move (Label1(1).Left + Label1(1).Width), .Top, .Width, .Height
        DTPicker1.Item(1).CustomFormat = .CustomFormat
        DTPicker1.Item(1).Format = .Format
        DTPicker1.Item(1).Value = Date + 1
        
    End With
    
    With Label1.Item(2)
        .Caption = "��������"
        .Alignment = intAli
        .ForeColor = lngColor
        
        Combo2.Item(0).Move (Combo1.Left), .Top - 30, Combo1.Width
        Combo2.Item(1).Visible = False
        
        Label1.Item(3).Caption = "��������"
        .Alignment = intAli
        Label1.Item(3).Move Check1.Left, .Top, .Width, .Height
        Label1.Item(3).ForeColor = lngColor
        
        Text1.Text = ""
        Text1.Move DTPicker1(0).Left, .Top - 30, (DTPicker1(1).Left + DTPicker1(1).Width - DTPicker1(0).Left), .Height
        
    End With
    
    With Command1
        .Caption = "��ѯ"
        .Height = 400
        .Move (Text1.Left + Text1.Width + 1000), (Text1.Top + Text1.Height - DTPicker1(1).Top - .Height) / 2 + DTPicker1(1).Top, 1000
        
        Command2.Caption = "�˳�"
        Command2.Move (.Left + .Width + 3000), .Top, .Width, .Height
    End With
    
    With Grid1
        .Appearance = 0
        .FixedCols = 1
        .FixedRows = 1
        .Cols = 9
        Rem .FormatString = "^���|^�����û�|^����ʱ��|^��������|��������|��������|^����IP|��������|ϵͳ����"
        .BackColorBkg = Me.BackColor
        .BackColorFixed = RGB(121, 151, 219)
        .BackColor2 = RGB(250, 235, 215)
        .AllowUserResizing = True
        .BackColorFixedSel = vbYellow

        Frame2.Top = (.Top + .Height)
        
    End With
        
    With Command3.Item(1)
        .Caption = "��һҳ"
        .Move 120, 120, 800, 375
        
        Command3.Item(2).Caption = "��һҳ"
        Command3.Item(2).Move (.Left + .Width), .Top, .Width, .Height
        
        Command3.Item(3).Caption = "��һҳ"
        Command3.Item(3).Move (.Left + .Width * 2), .Top, .Width, .Height
        
        Command3.Item(4).Caption = "���ҳ"
        Command3.Item(4).Move (.Left + .Width * 3), .Top, .Width, .Height
        
        Label1.Item(21).Caption = "��    ҳ"
        Label1.Item(21).Move (.Left + .Width * 4), .Top + 100, .Width * 1.5, .Height

        
        Label1.Item(22).Caption = "������         ҳ"
        Label1.Item(22).Move (.Left + .Width * 5.5), Label1(21).Top, .Width * 2, .Height
        Label1.Item(22).ForeColor = vbMagenta
        
        Text2.Move (.Left + .Width * 6.17), Label1(21).Top - 30, .Width, 255
        Text2.Text = ""
        Text2.Alignment = 2
        
        Command3.Item(0).Caption = "��ת"
        Command3.Item(0).Move (.Left + .Width * 7.5), .Top, .Width, .Height
        
    End With

    With Frame2     '��ҳ�����
        .Caption = ""
        .BorderStyle = 0
        .Width = Command3.Item(0).Left + Command3.Item(3).Width + 120
        .Height = Command3.Item(0).Top + Command3.Item(0).Height + 120
        .Left = Grid1.Left + (Grid1.Width - .Width) / 2
        lngSize.frameLeft = .Left
        lngSize.frameTop = .Top
    End With

    Me.Move 0, 0, lngSize.frmWidth, lngSize.frmHeight
    ctlMove.BorderStyle = 0
    ctlMove.Move 120, 120, 25000, 20000
    
    Call msSetTable
    Call msLoadMen
    

    For lngColor = udSelect To udUpdateBatch
        Combo2.Item(0).AddItem gfBackLogType(lngColor)
    Next
    
    Call gfLoadAuthority(Me, Command1)
    
End Sub

Private Sub Form_Resize()

    Dim lngW As Long
    Dim lngH As Long
    Dim lngVar As Long
    
    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 14400, 9000)
    
    If gWind.ActiveForm Is Nothing Then Exit Sub
    If gWind.ActiveForm.Name <> Me.Name Then Exit Sub
    If gWind.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState = vbMinimized Then Exit Sub
    
    lngW = Me.Width
    lngH = Me.Height
    
    If lngW > lngSize.frmWidth Then     '����ȱ仯
        lngVar = lngW - lngSize.frmWidth
    Else
        lngVar = 0
    End If
    Frame1.Width = lngSize.vsWidth + lngVar
    Grid1.Width = Frame1.Width

    Frame2.Left = lngSize.frameLeft + lngVar / 2
    lngAddWith = lngVar

    If lngH > lngSize.frmHeight Then    '���߶ȱ仯
        lngVar = lngH - lngSize.frmHeight
    Else
        lngVar = 0
    End If
    Grid1.Height = lngSize.vsHeight + lngVar
    Frame2.Top = lngSize.frameTop + lngVar
    lngPageSize = lngSize.pageSize + Int(lngVar / lngSize.rowHeight / 15)
    Grid1.Rows = Grid1.FixedRows + lngPageSize
    
    If Len(Grid1.Cell(Grid1.FixedRows, 0).Text) > 0 Then Call msShowValue  '������¸�ֵ
    
End Sub


Private Sub Hsb_Change()
    ctlMove.Left = -Hsb.Value
End Sub

Private Sub Hsb_Scroll()
    Call Hsb_Change    'Ҳ�ɲ���Ӵ�Scroll�¼�������ͬ��
End Sub

Private Sub Vsb_Change()
    ctlMove.Top = -Vsb.Value
End Sub

Private Sub Vsb_Scroll()
    Call Vsb_Change
End Sub

Private Sub msSetTable()
    '''���ñ���ʽ
    
    With Grid1
        .AutoRedraw = False
        .Cell(0, 0).Text = "���"
        .Cell(0, 1).Text = "�����û�"
        .Cell(0, 2).Text = "��¼ʱ��"
        .Cell(0, 3).Text = "��������"
        .Cell(0, 4).Text = "��־��ϸ����"
        .Cell(0, 5).Text = "��������"
        .Cell(0, 6).Text = "����ϵͳ����"
        .Cell(0, 7).Text = "�����ߵ���IP"
        .Cell(0, 8).Text = "�����ߵ�����"
        .ExtendLastCol = True '���һ�п�ȶ�����ĩ��
        .AllowUserSort = True '����˫����������
        
        .Rows = lngPageSize + 1
        .rowHeight(0) = 30
        .Column(0).Width = 65
        .Column(1).Width = 110
        .Column(2).Width = 190
        .Column(3).Width = 80
        .Column(4).Width = 300
        .Column(5).Width = 80
        .Column(6).Width = 120
        .Column(7).Width = 100
        .Column(8).Width = 80
        If Not (LCase(gVar.UserLoginName) = LCase(gVar.AccountAdmin) _
          Or LCase(gVar.UserLoginName) = LCase(gVar.AccountSystem)) Then
            If Not gfLoadAuthority(Me, Me.Grid1) Then
                .Column(6).Width = 0 'ϵͳ������ֻ��ʾ��ϵͳ������
            End If
        End If
        .Enabled = True '��û��Ȩ��ʱ�ᱻdisable
        
        .AutoRedraw = True
    End With
    
End Sub

Private Sub msLoadMen()
    '''���ز������б�
    
    Dim strSQL As String
    Dim rsL As ADODB.Recordset
    
    strSQL = "SELECT DISTINCT RIGHT(LogUserFullName,LEN(LogUserFullName)-" & _
            "CHARINDEX(',',LogUserFullName)) AS [LogUserFullName] FROM tb_FT_Sys_OperationLog"
    Set rsL = gfBackRecordset(strSQL)
    
    If rsL.State = adStateClosed Then Exit Sub
    If Not (rsL.BOF And rsL.EOF) Then
        With Combo1
            .Clear
            While Not rsL.EOF
                .AddItem rsL.Fields("LogUserFullName")
                rsL.MoveNext
            Wend
        End With
    End If
    Set rsL = Nothing
    
End Sub

Private Sub msShowValue()
    
    Dim I As Long
    Dim K As Long
    Dim n As Long
    Dim W As Long
    
    If rsLog.State = adStateClosed Then Exit Sub
    
    rsLog.pageSize = lngPageSize
    lngPageCount = rsLog.PageCount
        
    If rsLog.RecordCount = 0 Then
        lngPageCur = 0
    Else
        If lngPageCur > lngPageCount Then lngPageCur = lngPageCount '''�淶��ǰҳ��
        If lngPageCur < 1 Then lngPageCur = 1
        rsLog.AbsolutePage = lngPageCur
        
        n = lngPageSize * (lngPageCur - 1) + 1  '''��һ����¼�����
        
        With Grid1
            .AutoRedraw = False
            For I = 1 To lngPageSize    '''��ָ��ҳ�����ݸ�ֵ�������
                If rsLog.EOF Then Exit For
                K = .FixedRows - 1 + I
                .Cell(K, 0).Text = CStr(n)
                .Cell(K, 1).Text = rsLog.Fields("LogUserFullName")
                .Cell(K, 2).Text = rsLog.Fields("LogTime")
                .Cell(K, 3).Text = rsLog.Fields("LogType")
                .Cell(K, 4).Text = rsLog.Fields("LogContent")
                W = InStr(rsLog.Fields("LogFormName"), ",")
                If W < 1 Then W = Len(rsLog.Fields("LogFormName"))
                .Cell(K, 5).Text = Right(rsLog.Fields("LogFormName"), Len(rsLog.Fields("LogFormName")) - W)
                .Cell(K, 6).Text = rsLog.Fields("LogTable")
                .Cell(K, 7).Text = rsLog.Fields("LogPCIP")
                .Cell(K, 8).Text = rsLog.Fields("LogPCName")
                n = n + 1
                rsLog.MoveNext
            Next
            
            If I < lngPageSize + 1 Then '''�����ֵ���ݲ���һҳ������ձ�������һ����¼���������
                For I = I To lngPageSize
                    K = .FixedRows - 1 + I
                    If Len(.Cell(K, 0).Text) = 0 Then Exit For
                    For n = 0 To .Cols - 1
                        .Cell(K, n).Text = ""
                    Next
                Next
            End If
            .Refresh
            .AutoRedraw = True
        End With
    End If
    
    Command3.Item(1).Enabled = IIf(lngPageCur < 2, False, True)     '''����4����ҳ��ť�Ŀ���״̬
    Command3.Item(2).Enabled = Command3.Item(1).Enabled
    Command3.Item(3).Enabled = IIf(lngPageCur = lngPageCount, False, True)
    Command3.Item(4).Enabled = Command3.Item(3).Enabled
    
    Label1.Item(21).Caption = "�� " & CStr(lngPageCount) & " ҳ"
    Text2.Text = CStr(lngPageCur)
    
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Grid1.Cell(Row, Col).Text = strLastTxt
End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    strLastTxt = Grid1.Cell(Row, Col).Text
End Sub


