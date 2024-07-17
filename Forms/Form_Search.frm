VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_Search 
   Caption         =   "Master Product List"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   11160
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid afg 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.ComboBox cmbFiletype 
         Height          =   390
         ItemData        =   "Form_Search.frx":0000
         Left            =   8400
         List            =   "Form_Search.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   735
         Left            =   10440
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1296
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "Form_Search.frx":0023
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cmbCols 
         Height          =   390
         ItemData        =   "Form_Search.frx":008B
         Left            =   2280
         List            =   "Form_Search.frx":0095
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtFind 
         Height          =   390
         Left            =   4200
         TabIndex        =   1
         ToolTipText     =   "enter the keyword"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   8400
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_Search.frx":00AF
      Top             =   0
   End
End
Attribute VB_Name = "Form_Search"
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
Private i As Long
Private qry As String
Private oExcel      As Object
Private oBook       As Object
Private oSheet      As Object
Private ttl         As Double

Private Sub cmbCols_Click()
    txtfind_KeyPress 13
End Sub

Private Sub cmdExport_Click()
    Dim spreasheet As String
    If cmbFiletype.ListIndex = 0 Then
        spreasheet = "Excel.Application"
    Else
        spreasheet = "Ket.Application"
    End If
    If afg.rows < 2 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        Dim k As Byte
        Screen.MousePointer = 11
        prog1.Visible = True
        With afg
            ttl = .rows - 1
            For i = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If i = 0 Then
                        oSheet.Cells(i + 2, k + 1).Font.Bold = True
                        oSheet.Cells(i + 2, k + 1) = .TextMatrix(i, k)
                    Else
                        oSheet.Cells(i + 2, k + 1) = .TextMatrix(i, k)
                    End If
                    If k = 5 Then
                        If Right$(.TextMatrix(i, k), 2) = vbCrLf Or Right$(.TextMatrix(i, k), 2) = vbNewLine Then
                            oSheet.Cells(i + 2, k + 1) = RTrim(Left$(.TextMatrix(i, k), Len(.TextMatrix(i, k)) - 2))
                        Else
                            oSheet.Cells(i + 2, k + 1) = RTrim(.TextMatrix(i, k))
                        End If
                    End If
                Next
                prog1.Value = (i * 100) / ttl
            Next
        End With
        With oSheet
            .Cells(1, 1) = "Export Date : " & Now
            .Columns("B:E").AutoFit
            .Range("A2:K" & ttl + 2).Borders.LineStyle = xlContinuous
        End With
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
        prog1.Visible = False
        Screen.MousePointer = 0
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub Form_Initialize()
    Me.WindowState = vbNormal
    
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

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub settingFG()
    With afg
        .rows = 2
        .Cols = 12
        .TextMatrix(0, 0) = "MC ID"
        .ColWidth(0) = 850
        .TextMatrix(0, 1) = "Man Power"
        .TextMatrix(0, 2) = "Part No"
        .ColWidth(2) = 3000
        .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Part Name"
        .ColWidth(3) = 3000
        .TextMatrix(0, 4) = "Category"
        .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "Mold No"
        .ColWidth(5) = 3000
        .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(0, 6) = "Subcont"
        .ColWidth(6) = 900
        .TextMatrix(0, 7) = "Cavity"
        .ColWidth(7) = 900
        .TextMatrix(0, 8) = "Cavity STD"
        .TextMatrix(0, 9) = "CT"
        .ColWidth(9) = 800
        .TextMatrix(0, 10) = "CT 2nd"
        .ColWidth(10) = 800
        .TextMatrix(0, 11) = "Priority"
        .ColWidth(11) = 900
    End With
End Sub

Private Sub Form_Load()
    AddTab Me
    settingFG
    BukaKoneksi
    activeTheme skinFD, Me
    Me.Height = 6225
    Me.Width = 11400
    cmbCols.ListIndex = 0
    Call WheelHook(Me.hwnd)
End Sub

Private Sub LoadDB()
    qry = "select a.partno,partname,prod_nomach,mold_no,cavity,cavity_std,ct,ct_2,priorit,subcont,manpower " _
        & ",catgory from loadcap_mst_product_r a inner join loadcap_proc b on a.partno=b.partno"
    Set RsGet = Con.Execute(qry)
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    ' Is the mouse over the control
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hwnd, xpos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      ' If so, respond accordingly
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, xpos, Ypos
          
        Case TypeOf ctl Is PictureBox
          PictureBoxZoom ctl, MouseKeys, Rotation, xpos, Ypos
          
        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then ctl.SetFocus
          
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
  
End Sub

Private Sub getList()
    With afg
'        If RsGet.RecordCount > 0 Then
            .rows = RsGet.RecordCount + 1
            For i = 1 To RsGet.RecordCount
                RsGet.AbsolutePosition = i
                .TextMatrix(i, 0) = IIf(IsNull(RsGet("prod_nomach")), "-", RsGet("prod_nomach"))
                .TextMatrix(i, 1) = RsGet("manpower")
                .TextMatrix(i, 2) = RsGet("partno")
                .TextMatrix(i, 3) = RsGet("partname")
                .TextMatrix(i, 4) = RsGet("catgory")
                .TextMatrix(i, 5) = RsGet("mold_no")
                .TextMatrix(i, 6) = RsGet("subcont")
                .TextMatrix(i, 7) = RsGet("cavity")
                .TextMatrix(i, 8) = RsGet("cavity_std")
                .TextMatrix(i, 9) = RsGet("ct")
                .TextMatrix(i, 10) = RsGet("ct_2")
                .TextMatrix(i, 11) = RsGet("priorit")
            Next
'        End If
    End With
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbCols.Left = Label1.Left
    cmbCols.Top = txtfind.Top
    cmbCols.Width = Label1.Width
    cmbFiletype.Top = cmbCols.Top
    cmbFiletype.Left = Label2.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Call WheelUnHook(Me.hwnd)
        DelTab Me
    End If
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LoadDB
        If Len(Trim(txtfind)) > 1 Then
            txtfind = FilterIn(txtfind)
            RsGet.Filter = adFilterNone
            RsGet.Filter = "prod_nomach LIKE '*" & txtfind & "*'"
            If RsGet.RecordCount = 0 Then
                RsGet.Filter = adFilterNone
                RsGet.Filter = "partno LIKE '*" & txtfind & "*'"
                If RsGet.RecordCount = 0 Then
                    RsGet.Filter = adFilterNone
                    RsGet.Filter = "partname LIKE '*" & txtfind & "*'"
                End If
            End If
        Else
            RsGet.Filter = adFilterNone
        End If
        If cmbCols.ListIndex = 0 Then
            RsGet.Sort = "prod_nomach asc, priorit asc"
        ElseIf cmbCols.ListIndex = 1 Then
            RsGet.Sort = "partno asc, priorit asc"
        End If
        getList
    End If
End Sub
