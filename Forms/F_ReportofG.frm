VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_ReportofG 
   Caption         =   "Report of Generated LoadCap Data"
   ClientHeight    =   5670
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
   ScaleHeight     =   5670
   ScaleWidth      =   10800
   Begin VB.ComboBox cmbFiletype 
      Height          =   405
      ItemData        =   "F_ReportofG.frx":0000
      Left            =   9120
      List            =   "F_ReportofG.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox CmbDocument 
      Height          =   405
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdFindLoadcap 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
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
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid agrid 
      Height          =   4575
      Left            =   45
      TabIndex        =   0
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
      Left            =   5640
      OleObjectBlob   =   "F_ReportofG.frx":0023
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_ReportofG.frx":0085
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportofG.frx":00EB
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "F_ReportofG.frx":0151
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "F_ReportofG.frx":0385
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "F_ReportofG.frx":03E9
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   6480
      OleObjectBlob   =   "F_ReportofG.frx":0453
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "F_ReportofG"
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
Dim i As Double
Dim colo As Double
Dim qry As String
Dim HKWs As Variant
Private oExcel      As Object 'Excel.Application
Private oBook       As Object 'Excel.Workbook
Private oSheet      As Object 'Excel.Worksheet


Private Sub settingFG()
    With agrid
        .Cols = 8: .ColWidth(0) = 700: .ColWidth(1) = 900: .ColWidth(2) = 3000: .ColWidth(3) = 2500
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
        .TextMatrix(0, i) = "Qty":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = "Need Day MC":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 700
        
        i = 6
        .TextMatrix(0, i) = "% MC":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 7
        .TextMatrix(0, i) = "PP":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 0
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellAlignment = flexAlignCenterCenter
            .Row = 1
            .CellAlignment = flexAlignCenterCenter
        Next
    End With
End Sub

Private Sub cmbPeriod_Click()
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    Screen.MousePointer = 11
    Const kolomtOs As String = "no_mach, ton_mach ,fltpp_hkw, lcd_itemdid,lc_itemname , cap_p_day, a.fltpp_ym, lcvsmach,neday,lcvsmach,lc_pp "
    qry = "select " & kolomtOs & " from loadcap_generate_d a inner join " _
        & " loadcap_generate_h b on a.lcd_itemdid=b.lc_itemid and " _
        & " a.fltpp_doc=b.fltpp_doc and a.fltpp_ym=b.fltpp_ym and a.fltpp_rev=b.fltpp_rev where a.fltpp_doc='" & CmbDocument & "'" _
        & " and a.fltpp_rev='" & txtRevision & "' and a.fltpp_ym='" & cmbPeriod & "' and lc_pp>0 and b.lc_subcont='no'" _
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
                 .TextMatrix(i, 4) = FormatNumber(RsGet("neday") * RsGet("cap_p_day"), 0)
                 .TextMatrix(i, 5) = RsGet("neday")
                 .TextMatrix(i, 6) = RsGet("lcvsmach")
                 .TextMatrix(i, 7) = RsGet("lc_pp")
            End With
            i = i + 1
            RsGet.MoveNext
        Wend
    End If
    '# YANG SUBCONT
    With agrid
        .rows = 1 + .rows
        .TextMatrix(.rows - 1, 0) = "SUBCONT"
        For colo = 0 To .Cols - 1
            .Col = colo
            .Row = .rows - 1
            .CellBackColor = vbGreen
        Next
        qry = "select " & kolomtOs & " from loadcap_generate_d a inner join " _
            & " loadcap_generate_h b on a.lcd_itemdid=b.lc_itemid and " _
            & " a.fltpp_doc=b.fltpp_doc and a.fltpp_ym=b.fltpp_ym and a.fltpp_rev=b.fltpp_rev where a.fltpp_doc='" & CmbDocument & "'" _
            & " and a.fltpp_rev='" & txtRevision & "' and a.fltpp_ym='" & cmbPeriod & "' and lc_pp>0 and b.lc_subcont='yes'" _
        & " order by no_mach asc, lc_customer asc, lcd_itemdid asc"
        Set RsGet = Con.Execute(qry)
        If RsGet.RecordCount > 0 Then
            While Not RsGet.EOF
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = RsGet("no_mach")
                .TextMatrix(.rows - 1, 1) = RsGet("ton_mach")
                .TextMatrix(.rows - 1, 2) = RsGet("lcd_itemdid")
                .TextMatrix(.rows - 1, 3) = RsGet("lc_itemname")
                .TextMatrix(.rows - 1, 4) = FormatNumber(RsGet("neday") * RsGet("cap_p_day"), 0)
                .TextMatrix(.rows - 1, 5) = RsGet("neday")
                .TextMatrix(.rows - 1, 6) = RsGet("lcvsmach")
                .TextMatrix(.rows - 1, 7) = RsGet("lc_pp")
                RsGet.MoveNext
            Wend
        End If
        '# YANG OVERLOAD
        .rows = 1 + .rows
        .TextMatrix(.rows - 1, 0) = "Unprocessed"
        For colo = 0 To .Cols - 1
            .Col = colo
            .Row = .rows - 1
            .CellBackColor = vbGreen
        Next
        qry = "select " & kolomtOs & ",lc_sisa_pp from loadcap_generate_d a inner join " _
            & " loadcap_generate_h b on a.lcd_itemdid=b.lc_itemid and " _
            & " a.fltpp_doc=b.fltpp_doc and a.fltpp_ym=b.fltpp_ym and a.fltpp_rev=b.fltpp_rev where a.fltpp_doc='" & CmbDocument & "'" _
            & " and a.fltpp_rev='" & txtRevision & "' and a.fltpp_ym='" & cmbPeriod & "' and lc_pp>0 and b.lc_subcont='no' and neday=0 AND lc_sisa_pp>0" _
        & " order by no_mach asc, lc_customer asc, lcd_itemdid asc"
        Set RsGet = Con.Execute(qry)
        If RsGet.RecordCount > 0 Then
            While Not RsGet.EOF
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = RsGet("no_mach")
                .TextMatrix(.rows - 1, 1) = RsGet("ton_mach")
                .TextMatrix(.rows - 1, 2) = RsGet("lcd_itemdid")
                .TextMatrix(.rows - 1, 3) = RsGet("lc_itemname")
                .TextMatrix(.rows - 1, 4) = FormatNumber(RsGet("lc_sisa_pp"), 0)
                .TextMatrix(.rows - 1, 5) = RsGet("neday")
                .TextMatrix(.rows - 1, 6) = RsGet("lcvsmach")
                .TextMatrix(.rows - 1, 7) = RsGet("lc_pp")
                RsGet.MoveNext
            Wend
        End If
    End With
    Screen.MousePointer = 0
        
End Sub

Private Sub cmbPeriod_DropDown()
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    qry = "select distinct on (fltpp_ym) fltpp_ym from loadcap_generate_d where fltpp_doc='" & CmbDocument & "'" _
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

Private Sub cmdFindLoadcap_Click()
    popup_loadcap.Show 1
    CmbDocument.Text = popup_loadcap.docSelcd
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub CmbDocument_DropDown()
    qry = "select * from " _
    & " (select distinct on (fltpp_doc) fltpp_doc from loadcap_generate_d) v1 " _
    & " order by right(fltpp_doc,4) asc,substring(fltpp_doc from 17 for 2) "
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
    cmbFiletype.Top = SkinLabel6.Top
    cmbFiletype.Left = cmdExport.Left
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
    qry = "select distinct on (fltpp_rev) fltpp_rev from loadcap_generate_d where fltpp_doc='" & CmbDocument & "'"
    Set RsGet = Con.Execute(qry)
    txtRevision.Clear
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            txtRevision.AddItem RsGet(0)
            RsGet.MoveNext
        Wend
    End If
End Sub
