VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_SettingOffMPP 
   Caption         =   "Setting Off Day"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   8280
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   7995
      TabIndex        =   2
      Top             =   120
      Width           =   8055
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbYear 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "HKW : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   6000
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   6600
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_SettingOffMPP.frx":0000
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGridDATE 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483631
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form_SettingOffMPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type CtrlProportion
    Heightproportion As Single
    WidthProportion As Single
    TopProportion As Single
    LeftProportion As Single
End Type
Dim proportionArray() As CtrlProportion
Const MERAH As Integer = 255
Const HIJAU As Integer = 93
Const BIRU As Integer = 93
Dim i As Integer
Dim stDate As Boolean
Dim rNo As Integer
Dim valDate As Variant
Dim arrDate() As Variant

Sub ResizeControls()
    On Error Resume Next
    
    For i = 0 To Controls.Count - 1
        With proportionArray(i)
            ' move and resize controls
            Controls(i).Move .LeftProportion * ScaleWidth, _
            .TopProportion * ScaleHeight, _
            .WidthProportion * ScaleWidth, _
            .Heightproportion * ScaleHeight
        End With
    Next
End Sub

Private Sub initDataSaved(dYear As Integer, dMonth As Integer)
    Set RsGet = Con.Execute("select work_date, work_status from plansys_setoffday where extract(year from work_date) = '" & dYear & "' and extract(month from work_date) = '" & dMonth & "' order by work_date")
    If RsGet.RecordCount > 0 Then
        ReDim arrDate(1 To RsGet.RecordCount, 1 To 2)
        cmdSave.Caption = "UPDATE"
        i = 0
        Do Until RsGet.EOF
            i = i + 1
            arrDate(i, 1) = RsGet!work_date
            arrDate(i, 2) = RsGet!work_status
        
            RsGet.MoveNext
        Loop
        RsGet.Fields("work_status").Properties("Optimize") = True
        RsGet.Filter = adFilterNone
        RsGet.Filter = "work_status=1"
        Label5.Caption = "HKW : " & RsGet.RecordCount
    Else
        cmdSave.Caption = "SAVE"
    End If
End Sub

Private Sub getListDayPmonth(dYear As Integer, dMonth As Integer)
    settingHeaderGrid
    initDataSaved dYear, dMonth
    With MSFlexGridDATE
        valDate = DateSerial(dYear, dMonth, 1)
        stDate = True
        i = Weekday(valDate)
        If i = 7 Then i = 0
        rNo = 1
        Do Until stDate = False
            .TextMatrix(rNo, i) = Day(valDate)
            .Col = i
            .Row = rNo
            If cmdSave.Caption = "UPDATE" Then
                If arrDate(Day(valDate), 2) = 0 Then
                    .CellBackColor = RGB(MERAH, HIJAU, BIRU)
                    'MsgBox .CellForeColor
                    If .CellForeColor = vbRed Then
                        .CellForeColor = vbWhite
                    End If
                Else
                    .CellBackColor = RGB(255, 255, 255)
                End If
            Else
                .CellBackColor = RGB(255, 255, 255)
                If .CellForeColor = vbWhite Then
                    .CellForeColor = vbRed
                End If
            End If
            If i = 1 Then  'i = 0 Or
                If .CellBackColor = RGB(MERAH, HIJAU, BIRU) Then
                    .CellForeColor = vbWhite
                Else
                    .CellForeColor = vbRed
                End If
            End If
            valDate = valDate + 1
            If Month(valDate) > dMonth Or Year(valDate) > dYear Then
                stDate = False
            End If
            If i < 6 Then
                i = i + 1
            Else
                If stDate = True Then
                    i = 0
                    rNo = rNo + 1
                    .rows = .rows + 1
                    '.Height = .rows '* 615
                End If
            End If
        Loop
    End With
    Call Form_Resize
    cmdSave.Refresh
End Sub

Private Sub settingHeaderGrid()
    With MSFlexGridDATE
        .Clear
        .BackColor = RGB(255, 255, 255)
        .RowHeightMin = 600
        .Cols = 7
        .rows = 2
        .FixedCols = 0
        .FixedRows = 1
        For i = 0 To 6
            .ColAlignment(i) = flexAlignCenterCenter
            .TextMatrix(0, i) = UCase(Format(i, "DDDD"))
            If i = 1 Then  'i = 0 Or
                .Col = i
                .Row = 0
                .CellForeColor = vbRed
            End If
        Next
    End With
End Sub

Private Sub getListMonth()
    Dim iMonth As Integer
    cmbMonth.Clear
    For iMonth = 1 To 12
        cmbMonth.AddItem Format(DateSerial(Year(Now), iMonth, 1), "MMMM")
    Next
    cmbMonth.ListIndex = Month(Now) - 1
End Sub

Private Sub getListYear()
    Dim iYear As Integer
    cmbYear.Clear
    For iYear = 0 To 2
        cmbYear.AddItem Year(Now) - 1 + iYear
    Next
    cmbYear.ListIndex = 1
End Sub

Private Sub cmbMonth_Click()
    getListDayPmonth Val(cmbYear.Text), cmbMonth.ListIndex + 1
    
End Sub

Private Sub cmbYear_Click()
    getListDayPmonth Val(cmbYear.Text), cmbMonth.ListIndex + 1
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSave
    Dim sNo As Integer
    Dim sSt As Boolean
    Dim wDate As Date
    With MSFlexGridDATE
        For sNo = 1 To .rows - 1
            For i = 0 To 6
                If Val(.TextMatrix(sNo, i)) <> 0 Then
                    .Row = sNo
                    .Col = i

                    If .CellBackColor = RGB(255, 255, 255) Then  'RGB(255, 255, 255)
                        sSt = True
                    Else
                        sSt = False
                    End If
                    wDate = DateSerial(cmbYear, cmbMonth.ListIndex + 1, Val(.TextMatrix(sNo, i)))
                    If LCase(cmdSave.Caption) = "save" Then
                        Con.Execute "insert into plansys_setoffday (work_date, work_status, time_update, user_update) values " _
                            & "('" & Format(wDate, "YYYY-MM-DD") & "', " & sSt & ", now(), '" & pUserName & "')"
                    Else
                        Con.Execute "update plansys_setoffday set work_status = " & sSt & ", time_update = now(), user_update = '" & pUserName & "' where work_date = '" & Format(wDate, "YYYY-MM-DD") & "'"
                    End If
                End If
            Next
        Next
    End With
    MsgBox "Setting Updated...", vbInformation, "Information"
    Call cmbMonth_Click
Exit Sub
errSave:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Save: " & Err.Number
    End If
'    Form_SettingOffDay.Show
'    Form_SettingOffDay.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.WindowState = vbNormal
    Dim i As Integer
    On Error Resume Next
    
    ReDim proportionArray(0 To Controls.Count - 1)
    
    For i = 0 To Controls.Count - 1
         With proportionArray(i)
            .Heightproportion = Controls(i).Height / ScaleHeight
            .WidthProportion = Controls(i).Width / ScaleWidth
            .TopProportion = Controls(i).Top / ScaleHeight
            .LeftProportion = Controls(i).Left / ScaleWidth
         End With
    Next
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingHeaderGrid
    Call getListMonth
    Call getListYear
    Me.Height = 6540
    Me.Width = 8520
  
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    With MSFlexGridDATE
        For i = 0 To .Cols - 1
            .ColWidth(i) = .Width / 7 '(.ColWidth(0) + .ColWidth(1))
        Next
        For i = 0 To .rows - 1
            .RowHeight(i) = .Height / .rows - 1
        Next
    End With
    cmbYear.Top = Label2.Top
    cmbYear.Left = Label4.Left
    cmbMonth.Top = Label1.Top
    cmbMonth.Left = Label3.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub MSFlexGridDATE_Click()
On Error GoTo errClick
    With MSFlexGridDATE
        If .Row > 0 Then
            If .TextMatrix(.Row, .Col) <> "" Then
                If .CellBackColor = RGB(255, 255, 255) Then
                    .CellBackColor = RGB(MERAH, HIJAU, BIRU)
                    If .CellForeColor = vbRed Then
                        .CellForeColor = vbWhite
                    End If
                Else
                    .CellBackColor = RGB(255, 255, 255)
                    If .CellForeColor = vbWhite Then
                        .CellForeColor = vbRed
                    End If
                End If
            End If
        End If
    End With
Exit Sub
errClick:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Click: " & Err.Number
    End If
End Sub
