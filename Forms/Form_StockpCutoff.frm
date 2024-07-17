VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_StockpCutoff 
   Caption         =   "Stock at that Time"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9510
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
   ScaleHeight     =   5280
   ScaleWidth      =   9510
   Begin VB.CheckBox ckCust 
      Caption         =   "ALL"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   600
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton cmdCust 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtCust 
      BackColor       =   &H00FFFFC0&
      Height          =   390
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find Next"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.OptionButton OptO 
      Caption         =   "Other"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.OptionButton OptM 
      Caption         =   "Microsoft"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   600
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dttgl 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   202833923
      CurrentDate     =   43125
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7080
      OleObjectBlob   =   "Form_StockpCutoff.frx":0000
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_StockpCutoff.frx":0234
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7223
      _Version        =   393216
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_StockpCutoff.frx":0294
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form_StockpCutoff"
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
Dim posisisFind As Long
Dim oExcel As Object
Dim oBook  As Object
Dim oSheet As Object
Public custID As String

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

Private Sub settingGrid()
    With grid
        .rows = 2
        .Cols = 4
        .TextMatrix(0, 0) = "No"
        .ColWidth(0) = 500
        .TextMatrix(0, 1) = "Part Number"
        .ColWidth(1) = 2500
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Part Name"
        .ColWidth(2) = 3500
        .TextMatrix(0, 3) = "Qty"
        .ColWidth(3) = 2500
    End With
End Sub

Private Sub ckcUST_Click()
    If ckcUST.Value = vbChecked Then
        txtCust = ""
    End If
End Sub

Private Sub cmdCust_Click()
    GetForm = Me.Name
    popUp_Customer.Show 1
    ckcUST.Value = vbUnchecked
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
        If OptM.Value Then
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

Private Sub cmdView_Click()
    Dim tgl As String
    Dim i As Integer
    Dim WHre As String
    If ckcUST.Value = vbUnchecked Then
        WHre = " and cust_id='" & custID & "'"
    End If
    tgl = Format(dttgl, "yyyy-MM-dd")
    qry = "select rtrim(item_id,' ') item_id, item_name ,coalesce(ith_qty,0) ith_qty from " _
    & " (SELECT ith.ith_item_id, sum(ith.ith_qty) AS ith_qty " _
    & " FROM ith where ith_date<='" & tgl & "' and ith_item_id not like '%TEST%'" _
    & " GROUP BY ith.ith_item_id ) v1 right join mst_item a on v1.ith_item_id=a.item_id where pfm_id='10' " & WHre & " order by 1 asc"
    Set RsBantu = Con.Execute(qry)
    With grid
        .rows = 1
        .rows = 1 + RsBantu.RecordCount
        i = 1
        While Not RsBantu.EOF
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = RsBantu("item_id")
            .TextMatrix(i, 2) = RsBantu("item_name")
            .TextMatrix(i, 3) = FormatNumber(RsBantu("ith_qty"), 0)
            i = i + 1
            RsBantu.MoveNext
        Wend
    End With
    
End Sub

Private Sub Command1_Click()
    Dim xf As Double, pos As Integer
    Dim ttlrows As Double
    Dim stringCari As String
    With grid
        ttlrows = .rows - 1
        If posisisFind + 1 >= ttlrows Then
            posisisFind = 2
        Else
            posisisFind = 1 + posisisFind
        End If
        For xf = posisisFind To ttlrows
            stringCari = LCase$(.TextMatrix(xf, 1))
            pos = InStr(stringCari, LCase$(txtFindNext))
            If pos > 0 Then
                .Row = xf
                .Col = 2
                .TopRow = xf
                posisisFind = xf
                Exit For
            End If
        Next
        If pos = 0 Then posisisFind = 2
    End With
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
    settingGrid
    Call BukaKoneksi
    Call activeTheme(Skin1, Me)
    Me.Width = 9750
    Me.Height = 5850
    dttgl = Now
'    Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        PicFIND.Visible = True
        txtFindNext.SetFocus
    ElseIf KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText grid.Clip
        MsgBox "Copied"
    End If
End Sub

Private Sub Label14_Click()
    PicFIND.Visible = False
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MousePointer = 15
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lX As Integer, lY As Single
    If Button = vbLeftButton Then
        PicFIND.Left = PicFIND.Left + (x / 15 - lX)
        PicFIND.Top = PicFIND.Top + (Y / 15 - lY)
    Else
        lX = x / 15: lY = Y / 15
    End If
End Sub

Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
MousePointer = 0
End Sub

Private Sub txtFindNext_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        PicFIND.Visible = False
    ElseIf KeyAscii = 13 Then
        Command1_Click
    End If
End Sub
