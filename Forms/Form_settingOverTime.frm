VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_settingOverTime 
   Caption         =   "Setting Overtime"
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
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
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   7995
      TabIndex        =   0
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
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
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
         Left            =   4320
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
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
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   975
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
         Left            =   4320
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   975
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
         Left            =   840
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_settingOverTime.frx":0000
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGridDATE 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      BackColor       =   16777215
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483631
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form_settingOverTime"
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
Const MERAH As Integer = 93
Const HIJAU As Integer = 255
Const BIRU As Integer = 93
Dim i As Integer
Dim K As Integer
Dim x As Integer
Dim i2 As Integer
Dim iCol As Integer
Dim nBulan() As String
Dim nDataa() As String
Dim qry As String

Dim boCA As Integer
Dim boCB As Integer
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

Private Sub getListMonth()
    If cmbYear.ListCount > 0 And cmbYear <> "" Then
        Set RsGet = Con.Execute("select distinct on (extract(month from work_date)) extract(month from work_date) bulan " _
            & " from plansys_setoffday where " _
            & " extract(year from work_date)=" & cmbYear)
        cmbMonth.Clear
        If RsGet.RecordCount > 0 Then
            ReDim nBulan(1 To RsGet.RecordCount, 1 To 2) As String
            i = 1
            While Not RsGet.EOF
                cmbMonth.AddItem Format(DateSerial(cmbYear, RsGet(0), 1), "MMMM")
                nBulan(i, 1) = IIf(Len(RsGet(0)) > 1, RsGet(0), "0" & RsGet(0))
                nBulan(i, 2) = Format(DateSerial(cmbYear, RsGet(0), 1), "MMMM")
                i = i + 1
                RsGet.MoveNext
            Wend
        End If
    End If
End Sub

Private Sub getListYear()
    Set RsGet = Con.Execute("select distinct on (extract(year from work_date )) extract(year from work_date ) tahun  from plansys_setoffday")
    cmbYear.Clear
    While Not RsGet.EOF
        cmbYear.AddItem RsGet(0)
        RsGet.MoveNext
    Wend
End Sub

Private Sub cmbMonth_Click()
    qry = "select extract(day from work_date) hari,to_char(work_date,'day') nm from plansys_setoffday where work_status=FALSE and " _
    & " extract(year from work_date)=" & cmbYear & " and extract(month from work_date)=" & getnBulan(cmbMonth) * 1 _
    & " order by 1 asc"
'    MsgBox qry
'    Clipboard.Clear
'    Clipboard.SetText qry
    Set RsGet = Con.Execute(qry)
    With MSFlexGridDATE
        If RsGet.RecordCount > 0 Then
            .Cols = RsGet.RecordCount + 1
            i = 1
            While Not RsGet.EOF
                .TextMatrix(0, i) = RsGet("hari") & vbNewLine & " (" & Left(RsGet("nm"), 3) & ")"
                .Col = i
                .Row = 0
                .CellAlignment = flexAlignCenterCenter
                .CellBackColor = RGB(254, 193, 93)
                For K = 1 To .rows - 1
                    .Col = i
                    .Row = K
                    .CellBackColor = RGB(255, 255, 255)
                Next
                i = i + 1
                RsGet.MoveNext
            Wend
        End If
    End With
    qry = "select extract(day from wrk_date) hari,no_mach from mpp_setovrtime where ovr_status=TRUE " _
        & " AND extract(month from wrk_date)=" & getnBulan(cmbMonth) * 1 & " " _
        & " AND extract(year from wrk_date)=" & cmbYear
    Set RsGet = Con.Execute(qry)
    With MSFlexGridDATE
        If RsGet.RecordCount > 0 Then
            ReDim nDataa(1 To RsGet.RecordCount, 1 To 2) As String
            i = 1
            While Not RsGet.EOF
                nDataa(i, 1) = RsGet("hari")
                nDataa(i, 2) = RsGet("no_mach")
                i = i + 1
                RsGet.MoveNext
            Wend
            For i = 1 To UBound(nDataa)
                For x = 1 To .rows - 1
                    If nDataa(i, 2) = .TextMatrix(x, 0) Then
                        iCol = getKolDtExist(i)
                        If getKolDtExist(i) > 0 Then
                            .Row = x
                            .Col = iCol
                            .CellBackColor = RGB(MERAH, HIJAU, BIRU)
                        End If
                    End If
                Next
            Next
        End If
        .SetFocus
    End With
    
End Sub

Private Function getKolDtExist(p1 As Integer) As Integer
    With MSFlexGridDATE
        For K = 1 To .Cols - 1
            If nDataa(p1, 1) = Left(.TextMatrix(0, K), 2) * 1 Then
                getKolDtExist = K
                Exit For
            Else
                getKolDtExist = 0
            End If
        Next
    End With
End Function

Private Sub cmbMonth_DropDown()
    Call getListMonth
End Sub

Private Sub cmbYear_DropDown()
    Call getListYear
End Sub

Private Function getnBulan(pnmBuln As String) As String
    For x = 1 To UBound(nBulan)
        If nBulan(x, 2) = pnmBuln Then
            getnBulan = nBulan(x, 1)
            Exit For
        Else
            getnBulan = "._."
        End If
    Next
End Function

Private Sub cmdSave_Click()
On Error GoTo errSave
    Dim sNo As Integer
    Dim sSt As Boolean
    Dim wDate As Date
    Dim mcOvr As String
    Dim sTgl As String
    Con.Execute "delete from mpp_setovrtime where " _
     & " extract(year from wrk_date)=" & cmbYear & " and extract(month from wrk_date)=" & getnBulan(cmbMonth)
    With MSFlexGridDATE
        For sNo = 1 To .rows - 1
            For i = 1 To .Cols - 1
                .Row = sNo
                .Col = i
                If .CellBackColor = RGB(MERAH, HIJAU, BIRU) Then
                    sSt = True
                Else
                    sSt = False
                End If
                sTgl = cmbYear & "-" & getnBulan(cmbMonth) & "-" & .TextMatrix(0, i) ' & " [" & .TextMatrix(sNo, 0) & "]"
'                MsgBox sTgl, vbInformation, sSt
                If LCase(cmdSave.Caption) = "save" Then
                    
                    If sSt Then
                        Con.Execute "insert into mpp_setovrtime (id_mppovr, wrk_date, no_mach, time_update, user_update,ovr_status) values " _
                            & "(DEFAULT,'" & sTgl & "','" & .TextMatrix(sNo, 0) & "', Now(), '" & pUserName & "', " & sSt & ")"
                    End If
                Else
                    Con.Execute "update mpp_setovrtime set ovr_status = " & sSt & ", time_update = now(), user_update = '" & pUserName & "' where work_date = '" & Format(wDate, "YYYY-MM-DD") & "'"
                End If
            Next
        Next
    End With
    MsgBox "Setting Updated...", vbInformation, "Information"

Exit Sub
errSave:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Save: " & Err.Number
    End If
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

Private Sub settingHeaderGrid()
    With MSFlexGridDATE
        .Clear
        .BackColor = RGB(255, 255, 255)
'        .RowHeightMin = 600
        .Cols = 2
        .rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .TextMatrix(0, 0) = "MC\DATE"
        .RowHeight(0) = 600
    End With
End Sub

Private Sub loadData()
    Set RsGet = Con.Execute("select idmst_mach,no_mach from loadcap_mst_mach order by no_mach asc")
    MSFlexGridDATE.rows = 1
    If RsGet.RecordCount > 0 Then
        MSFlexGridDATE.rows = RsGet.RecordCount + 1
        i = 1
        While Not RsGet.EOF
            With MSFlexGridDATE
                .TextMatrix(i, 0) = RsGet("no_mach")
            End With
            i = i + 1
            RsGet.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingHeaderGrid
    Me.Height = 6540
    Me.Width = 8520
    loadData
    Call WheelHook(Me.hWnd)
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbYear.Top = Label2.Top
    cmbYear.Left = Label4.Left
    cmbMonth.Top = Label1.Top
    cmbMonth.Left = Label3.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hWnd)
End Sub

Private Sub MSFlexGridDATE_Click()
On Error GoTo errClick
    With MSFlexGridDATE
        If .Row > 0 Then
            If .TextMatrix(0, .Col) <> "" Then 'jika belum termuat
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

Private Sub MSFlexGridDATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 1 Then
        DoEvents
        For i = 1 To MSFlexGridDATE.rows - 1
            For i2 = 1 To MSFlexGridDATE.Cols - 1
                MSFlexGridDATE.Row = i
                MSFlexGridDATE.Col = i2
                If MSFlexGridDATE.CellBackColor = RGB(MERAH, HIJAU, BIRU) Then
                    MSFlexGridDATE.CellBackColor = RGB(255, 255, 255)
                Else
                    MSFlexGridDATE.CellBackColor = RGB(MERAH, HIJAU, BIRU)
                End If
            Next
        Next
    End If
End Sub

Private Sub MSFlexGridDATE_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim cont As Integer
    Dim found As Boolean
    
'    MsgBox y, vbInformation, Shift
   'Only do it if the button is the left one, and there
   'are no mask keys pressed.
   'Of course, you can vary this to fit your needs.
    If (Button = vbLeftButton) And (Shift = 0) Then
        With MSFlexGridDATE
            'Do not proceed if the row clicked is
            'not the header row.
            If (.RowHeight(0) < Y) Then
                If .ColWidth(0) < x Then
                    Exit Sub
                End If
                'Initialize variables
                cont = 0
                found = False
                'Find the column clicked using mouse coords
                Do While (cont < .rows) And Not (found)
                    If (.rowPos(cont) + .RowHeight(cont) < Y) Then
                        cont = cont + 1
                    Else
                        found = True
                        boCA = cont
                    End If
                Loop
                DoEvents
                'If column found, proceed to run the appropriate code
                If found Then
                    If cont > 0 Then
                        For i = 1 To .Cols - 1
                            .Row = cont
                            .Col = i
                            If .CellBackColor <> RGB(MERAH, HIJAU, BIRU) Then
                                .CellBackColor = RGB(MERAH, HIJAU, BIRU) 'biru asin
                            Else
                                .CellBackColor = RGB(255, 255, 255)
                            End If
                        Next
                    End If
                End If
            Else
                'Initialize variables
                cont = 0
                found = False
                'Find the column clicked using mouse coords
                Do While (cont < .Cols) And Not (found)
                    If (.ColPos(cont) + .ColWidth(cont) < x) Then
                        cont = cont + 1
                    Else
                        found = True
                        boCB = cont
                    End If
                Loop
                DoEvents
                'If column found, proceed to run the appropriate code
                If found Then
                    If cont > 0 Then
                        For i = 2 To .rows - 1
                            .Row = i
                            .Col = cont
                            If .CellBackColor <> RGB(MERAH, HIJAU, BIRU) Then
                                .CellBackColor = RGB(MERAH, HIJAU, BIRU) 'biru asin
                            Else
                                .CellBackColor = RGB(255, 255, 255)
                            End If
                        Next
                    End If
                End If
            End If
            
        End With
    End If
End Sub

Private Sub MSFlexGridDATE_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If boCA > 0 Then
        With MSFlexGridDATE
            .Row = boCA
            .Col = .Cols - 1
            If .CellBackColor <> RGB(MERAH, HIJAU, BIRU) Then
                .CellBackColor = RGB(MERAH, HIJAU, BIRU) 'biru asin
            Else
                .CellBackColor = RGB(255, 255, 255)
            End If
        End With
        boCA = 0
    End If
    If boCB > 0 Then
        With MSFlexGridDATE
            .Row = .rows - 1
            .Col = boCB
            If .CellBackColor <> RGB(MERAH, HIJAU, BIRU) Then
                .CellBackColor = RGB(MERAH, HIJAU, BIRU) 'biru asin
            Else
                .CellBackColor = RGB(255, 255, 255)
            End If
        End With
        boCB = 0
    End If
End Sub


Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xPos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hWnd, xPos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, xPos, Ypos
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
End Sub

