VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_SettingOffDay 
   Caption         =   "Setting Off Day"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_SettingOffDay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   6480
      TabIndex        =   6
      Top             =   5280
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.Skin skn 
      Left            =   -240
      OleObjectBlob   =   "Form_SettingOffDay.frx":000C
      Top             =   960
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGridDATE 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
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
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   7995
      TabIndex        =   1
      Top             =   120
      Width           =   8055
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
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
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
         TabIndex        =   2
         Top             =   240
         Width           =   1815
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
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   735
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
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form_SettingOffDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Private Sub cmbMonth_Click()
On Error GoTo errClick
    dateGrid Val(cmbYear.Text), cmbMonth.ListIndex + 1
Exit Sub
errClick:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Click: " & Err.Number
    End If
End Sub

Private Sub cmbYear_Click()
On Error GoTo errClick
    dateGrid Val(cmbYear.Text), cmbMonth.ListIndex + 1
Exit Sub
errClick:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Click: " & Err.Number
    End If
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
                    If .CellBackColor = RGB(255, 255, 255) Then
                        sSt = True
                    Else
                        sSt = False
                    End If
                    wDate = DateSerial(cmbYear, cmbMonth.ListIndex + 1, Val(.TextMatrix(sNo, i)))
                    If cmdSave.Caption = "SAVE" Then
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
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call activeTheme(skn, Me)
    Call BukaKoneksi
    getListMonth
    getListYear
    'dateGrid Year(Now), Month(Now)
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub dateGrid(dYear As Integer, dMonth As Integer)
    Dim rNo As Integer
    Dim stDate As Boolean
    Dim valDate As Variant
    Dim arrDate() As Variant
    stDate = True
    valDate = DateSerial(dYear, dMonth, 1)
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
    Else
        cmdSave.Caption = "SAVE"
    End If
    RsGet.Close
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
            If i = 0 Or i = 1 Then  '
                .Col = i
                .Row = 0
                .CellForeColor = vbRed
            End If
        Next
        i = Weekday(valDate)
        If i = 7 Then i = 0
        rNo = 1
        Do Until stDate = False
            .TextMatrix(rNo, i) = Day(valDate)
            .Col = i
            .Row = rNo
            If cmdSave.Caption = "UPDATE" Then
                If arrDate(Day(valDate), 2) = 0 Then
                    .CellBackColor = RGB(220, 150, 150)
                Else
                    .CellBackColor = RGB(255, 255, 255)
                End If
            Else
                .CellBackColor = RGB(255, 255, 255)
            End If
            If i = 0 Or i = 1 Then  '
                .CellForeColor = vbRed
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
                    .Height = .rows * 615
                End If
            End If
        Loop
    End With
End Sub

Private Sub MSFlexGridDATE_Click()
On Error GoTo errClick
    With MSFlexGridDATE
        If .Row > 0 Then
            If .TextMatrix(.Row, .Col) <> "" Then
                If .CellBackColor = RGB(255, 255, 255) Then
                    .CellBackColor = RGB(220, 150, 150)
                Else
                    .CellBackColor = RGB(255, 255, 255)
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
