VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_ReportMnuLoading 
   Caption         =   "Report of Menu Loading"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
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
   ScaleHeight     =   5700
   ScaleWidth      =   10800
   Begin VB.ComboBox cmbFiletype 
      Height          =   405
      ItemData        =   "F_ReportMnuLoading.frx":0000
      Left            =   9120
      List            =   "F_ReportMnuLoading.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdlu_findDoc 
      Caption         =   "..."
      Height          =   375
      Left            =   4400
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox CmbDocument 
      Height          =   405
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox cmbPeriod 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox txtRevision 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid agrid 
      Height          =   4575
      Left            =   45
      TabIndex        =   3
      Top             =   1080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8070
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "F_ReportMnuLoading.frx":0023
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_ReportMnuLoading.frx":0085
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportMnuLoading.frx":00EB
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "F_ReportMnuLoading.frx":0151
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "F_ReportMnuLoading.frx":0385
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "F_ReportMnuLoading.frx":03E1
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "F_ReportMnuLoading.frx":044B
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "F_ReportMnuLoading"
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
Dim HKWs As Variant
Private oExcel      As Object 'Excel.Application
Private oBook       As Object 'Excel.Workbook
Private oSheet      As Object 'Excel.Worksheet

Private Sub settingFG()
    With agrid
        .Cols = 9: .ColWidth(0) = 700: .ColWidth(1) = 900: .ColWidth(2) = 3000: .ColWidth(3) = 2500
        .ColWidth(5) = 700: .ColWidth(6) = 1000: .ColWidth(4) = 2000
        .rows = 5
        .FixedRows = 2
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "MC ID":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 1
        .TextMatrix(0, i) = "Tonage":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 2
        .TextMatrix(0, i) = "Part No":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 3
        .TextMatrix(0, i) = "Part Name":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 4
        .TextMatrix(0, i) = "Mold Number":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 5
        .TextMatrix(0, i) = "Qty":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 6
        .TextMatrix(0, i) = "Need Day MC":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 700
        
        i = 7
        .TextMatrix(0, i) = "% MC":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
         i = 8
        .TextMatrix(0, i) = "Type":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellAlignment = flexAlignCenterCenter
            .Row = 1
            .CellAlignment = flexAlignCenterCenter
        Next
    End With
End Sub

Private Sub agrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText agrid.Clip
    End If
End Sub

Private Sub cmbPeriod_Click()
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    Screen.MousePointer = 11
    Const kolomtOs As String = "no_mach, ton_mach ,fltpp_hkw, lcd_itemdid,lc_itemname , cap_p_day, fltpp_ym, lcvsmach,neday,lcvsmach,lc_subcont,neqty,reg_mold "
    qry = "select " & kolomtOs & " from mpp_gen_d where fltpp_doc='" & CmbDocument & "'" _
        & " and fltpp_rev='" & txtRevision & "' and fltpp_ym='" & cmbPeriod & "' " _
    & " order by no_mach asc, lc_customer asc, lcd_itemdid asc"
    Set RsGet = Con.Execute(qry)
    agrid.rows = 2
    If RsGet.RecordCount > 0 Then
        agrid.rows = RsGet.RecordCount + 2
        HKWs = RsGet("fltpp_hkw")
        SkinLabel4.Caption = "HKW : " & RsGet("fltpp_hkw")
        i = 2
        While Not RsGet.EOF
            With agrid
                 .TextMatrix(i, 0) = RsGet("no_mach")
                 .TextMatrix(i, 1) = RsGet("ton_mach")
                 .TextMatrix(i, 2) = RsGet("lcd_itemdid")
                 .TextMatrix(i, 3) = RsGet("lc_itemname")
                 .TextMatrix(i, 4) = RsGet("reg_mold")
                 .TextMatrix(i, 5) = FormatNumber(RsGet("neqty"), 0) 'FormatNumber(RsGet("neday") * RsGet("cap_p_day"), 0)
                 .TextMatrix(i, 6) = RsGet("neday")
                 .TextMatrix(i, 7) = RsGet("lcvsmach")
                 .TextMatrix(i, 8) = RsGet("lc_subcont")
            End With
            i = i + 1
            RsGet.MoveNext
        Wend
    End If
    Screen.MousePointer = 0
        
End Sub

Private Sub cmbPeriod_DropDown()
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    qry = "select distinct on (fltpp_ym) fltpp_ym from mpp_gen_d  where fltpp_doc='" & CmbDocument & "'" _
    & " and fltpp_rev='" & txtRevision & "'"
    Set RsGet = Con.Execute(qry)
    cmbPeriod.Clear
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            cmbPeriod.AddItem RsGet(0)
            RsGet.MoveNext
        Wend
    End If
End Sub

Private Sub cmdExport_Click()
    Dim spreasheet      As String
    If cmbFiletype.ListIndex = 0 Then
        spreasheet = "Excel.Application"
    Else
        spreasheet = "Ket.Application"
    End If
    If agrid.rows < 2 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        Dim k As Integer
        With oSheet
            .Cells(1, 1) = "LTPP Document : " & CmbDocument
            .Cells(2, 1) = "Revision : " & txtRevision
            .Cells(3, 1) = "Period : " & cmbPeriod
            .Cells(4, 1) = "HKW : " & HKWs
            .Range(.Cells(1, 1), .Cells(4, 1)).Font.Bold = True
            .Columns(3).NumberFormat = "@"
            
        End With
        With agrid
            For i = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If i = 0 Then
                        oSheet.Cells(i + 6, k + 1).Font.Bold = True
                    Else
                        oSheet.Cells(i + 5, k + 1) = .TextMatrix(i, k)
                    End If
                Next
            Next
        End With
        oSheet.Columns("C:D").AutoFit
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub cmdlu_findDoc_Click()
    PopUp_MLDOC.Show 1
    CmbDocument.Text = PopUp_MLDOC.lu_nodoc
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub CmbDocument_DropDown()
    qry = " select * from " _
        & "(select distinct on (fltpp_doc) fltpp_doc  from mpp_gen_d ) v1 " _
        & " order by right(fltpp_doc,4) asc,substring(fltpp_doc from 17 for 2)"
    Set RsGet = Con.Execute(qry)
    CmbDocument.Clear
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            CmbDocument.AddItem RsGet(0)
            RsGet.MoveNext
        Wend
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



Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingFG
    Me.Height = 6240
    Me.Width = 11040
    cmbFiletype.ListIndex = 0
    Call WheelHook(Me.hwnd)
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    txtRevision.Top = SkinLabel2.Top
    txtRevision.Left = SkinLabel5.Left
    cmbPeriod.Left = SkinLabel6.Left
    cmbPeriod.Top = SkinLabel1.Top
    
    cmbFiletype.Left = cmdExport.Left
    cmbFiletype.Top = SkinLabel4.Top
    cmbFiletype.Width = cmdExport.Width
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

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub txtRevision_DropDown()
    If Len(CmbDocument) < 2 Then CmbDocument.SetFocus: Exit Sub
    qry = "select distinct on (fltpp_rev) fltpp_rev from mpp_gen_d  where fltpp_doc='" & CmbDocument & "'"
    Set RsGet = Con.Execute(qry)
    txtRevision.Clear
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            txtRevision.AddItem RsGet(0)
            RsGet.MoveNext
        Wend
    End If
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hwnd, xpos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, xpos, Ypos
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
End Sub


