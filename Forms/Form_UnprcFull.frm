VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_UnprcFull 
   Caption         =   "Report of Unprocessed Item"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   8220
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8070
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.OptionButton optKing 
         Caption         =   "KS"
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         ToolTipText     =   "Kingsoft"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optMic 
         Caption         =   "MS"
         Height          =   255
         Left            =   7080
         TabIndex        =   8
         ToolTipText     =   "Microsoft"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   735
         Left            =   6000
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdLu 
         Caption         =   "..."
         Height          =   345
         Left            =   3770
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox cmbRev 
         Height          =   345
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtdoc 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Form_UnprcFull.frx":0000
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Form_UnprcFull.frx":006E
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3360
      OleObjectBlob   =   "Form_UnprcFull.frx":00DE
      Top             =   120
   End
End
Attribute VB_Name = "Form_UnprcFull"
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
Dim i As Integer
Dim qry As String

Dim oExcel As Object
Dim oBook  As Object
Dim oSheet As Object

Private Sub settingGrid()
    With grid
        .rows = 3
        .Cols = 6
        .FixedRows = 2
        .FixedCols = 0
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
            .Row = 1
            .Col = i
            .CellAlignment = flexAlignCenterCenter
        Next
        .ColWidth(.Cols - 1) = 2000
        .MergeRow(1) = True
        .MergeCells = flexMergeRestrictRows
        .TextMatrix(0, 0) = "No"
        .TextMatrix(1, 0) = .TextMatrix(0, 0)
        .ColWidth(0) = 500
        .TextMatrix(0, 1) = "Part Number"
        .TextMatrix(1, 1) = .TextMatrix(0, 1)
        .ColWidth(1) = 2500
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Part Name"
        .TextMatrix(1, 2) = .TextMatrix(0, 2)
        .ColWidth(2) = 3500
        .TextMatrix(0, 3) = "Prod Plan"
        .TextMatrix(0, 4) = "Processed"
        .TextMatrix(0, 5) = "Unprocessed"
        .TextMatrix(1, 3) = "Qty"
        .TextMatrix(1, 4) = .TextMatrix(1, 3)
        .TextMatrix(1, 5) = .TextMatrix(1, 3)
       
    End With
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

Private Sub cmbRev_Click()
    Dim a As Integer
    qry = "select lcd_itemdid,item_name,sum(neqty) terplot,max(lc_pp) pp,max(lc_pp)-sum(neqty) sisa  from mpp_gen_d a inner join mst_item b on a.lcd_itemdid=b.item_id " _
    & " where fltpp_doc ='" & txtdoc & "'  and lc_pp>0 and fltpp_rev=" & cmbRev _
    & " group by lcd_itemdid,item_name having ceil(sum(neqty))<max(lc_pp) " _
    & " order by 1 asc"
    Set RsBantu = Con.Execute(qry)
    grid.rows = 2
    a = 2
    If RsBantu.RecordCount > 0 Then
        With grid
            .rows = 2 + RsBantu.RecordCount
            While Not RsBantu.EOF
                .TextMatrix(a, 0) = a - 1
                .TextMatrix(a, 1) = RsBantu("lcd_itemdid")
                .TextMatrix(a, 2) = RsBantu("item_name")
                .TextMatrix(a, 3) = FormatNumber(RsBantu("pp"), 0)
                .TextMatrix(a, 4) = FormatNumber(RsBantu("terplot"), 0)
                .TextMatrix(a, 5) = FormatNumber(RsBantu("sisa"), 0)
                a = a + 1
                RsBantu.MoveNext
            Wend
        End With
    End If
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


Private Sub cmdExport_Click()
    If grid.rows < 2 Then MsgBox "nothing to be exported": Exit Sub
    If grid.TextMatrix(1, 0) = "" Then Exit Sub
    Clipboard.Clear
    With grid
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = .rows - 1
        Clipboard.SetText .Clip
    End With
    
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Dim spreasheet      As String
        If optMic.Value Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        oSheet.Cells(1, 1) = "Time Export : " & Now
        With oExcel.ActiveWorkbook.ActiveSheet
            .Range("A2").Select 'Select Cell A1 (will paste from here, to different cells)
            .Paste              'Paste clipboard contents
        End With
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName ', xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub cmdLu_Click()
   
    PopUp_MLDOC.Show 1
    txtdoc.Text = PopUp_MLDOC.lu_nodoc
    qry = "select distinct fltpp_rev from mpp_gen_d where fltpp_doc='" & txtdoc & "' order by 1 asc"
    Set RsBantu = Con.Execute(qry)
    cmbRev.Clear
    While Not RsBantu.EOF
        cmbRev.AddItem RsBantu("fltpp_rev")
        RsBantu.MoveNext
    Wend
End Sub

Private Sub Form_Activate()
    FocusTab Me
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

Private Sub Form_Load()
    AddTab Me
    activeTheme Skin1, Me
    settingGrid
    Width = 8460
    Height = 6525
    Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbRev.Width = Label1.Width
    cmbRev.Left = txtdoc.Left
    cmbRev.Top = Label1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText grid.Clip
        MsgBox "Copied"
    End If
End Sub
