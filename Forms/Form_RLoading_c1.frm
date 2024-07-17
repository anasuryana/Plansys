VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_RLoading_c1 
   Caption         =   "Data of Loading and Capacity"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
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
   ScaleHeight     =   6450
   ScaleWidth      =   13200
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox cmbMachine 
      Height          =   390
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6840
      OleObjectBlob   =   "Form_RLoading_c1.frx":0000
      Top             =   840
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   390
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox txtRevision 
      Height          =   390
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox CmbDocument 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdlu_docno 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9340
      _Version        =   393216
      MergeCells      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "Form_RLoading_c1.frx":0234
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_RLoading_c1.frx":0296
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_RLoading_c1.frx":02FC
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "Form_RLoading_c1.frx":0362
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel lbl_hkw 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "Form_RLoading_c1.frx":03BE
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   8040
      OleObjectBlob   =   "Form_RLoading_c1.frx":0428
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   7800
      OleObjectBlob   =   "Form_RLoading_c1.frx":048C
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form_RLoading_c1"
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
Dim i As Long
Dim qry As String
Public cust_id As String

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
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    Screen.MousePointer = 11
    Const kolomtOs As String = "no_mach, ton_mach ,fltpp_hkw, lcd_itemdid,lc_itemname , cap_p_day, a.fltpp_ym, lcvsmach,neday,lc_pp,cust_name,lc_fprodtvty "
    Dim HKWs As String
    Dim nom As Integer
    qry = "select " & kolomtOs & " from loadcap_generate_d a inner join " _
        & " loadcap_generate_h b on a.lcd_itemdid=b.lc_itemid and " _
        & " a.fltpp_doc=b.fltpp_doc and a.fltpp_ym=b.fltpp_ym and a.fltpp_rev=b.fltpp_rev " _
        & " inner join mst_item c on a.lcd_itemdid=c.item_id" _
        & " inner join r_customer d on c.cust_id=d.cust_id" _
        & " where a.fltpp_doc='" & CmbDocument & "'" _
        & " and a.fltpp_rev='" & txtRevision & "' and a.fltpp_ym='" & cmbPeriod & "' and lc_pp>0 and b.lc_subcont='no'" _
        & " and no_mach='" & cmbMachine & "'" _
    & " order by lc_customer asc, lcd_itemdid asc"
    
    Set RsGet = Con.Execute(qry)
    grid1.rows = 4
    If RsGet.RecordCount > 0 Then
        grid1.rows = RsGet.RecordCount + 4
        HKWs = RsGet("fltpp_hkw")
        SkinLabel4.Caption = "HKW : " & RsGet("fltpp_hkw")
        i = 3
        nom = 1
        While Not RsGet.EOF
            With grid1
                 .TextMatrix(i, 0) = RsGet("cust_name")
                 .TextMatrix(i, 1) = nom
                 .TextMatrix(i, 2) = RsGet("lcd_itemdid")
                 .TextMatrix(i, 3) = RsGet("lc_itemname")
                 .TextMatrix(i, 4) = "Injection"
                 
                 .TextMatrix(i, 6) = RsGet("lc_fprodtvty") * 100 & "%"
            End With
            i = i + 1
            nom = nom + 1
            RsGet.MoveNext
        Wend
    End If
    Screen.MousePointer = 0
End Sub

Sub loadMachnine()
    qry = "SELECT no_mach FROM loadcap_mst_mach order by 1 ASC"
    Set RsBantu = Con.Execute(qry)
    cmbMachine.Clear
    
    While Not RsBantu.EOF
        cmbMachine.AddItem RsBantu(0)
        RsBantu.MoveNext
    Wend
    
End Sub

Private Sub cmdlu_docno_Click()
    popup_loadcap.Show 1
    CmbDocument.Text = popup_loadcap.docSelcd
End Sub

Private Sub Command1_Click()
    popUp_Customer.Show 1
    
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

Private Sub settingFG()
    With grid1
        .rows = 4
        .FixedRows = 3
        .FixedCols = 0
        .Cols = 14
        .WordWrap = True
        .TextMatrix(0, 0) = "Customer"
        .TextMatrix(1, 0) = .TextMatrix(0, 0)
        .TextMatrix(2, 0) = .TextMatrix(0, 0)
        .MergeCol(0) = True
        
        .TextMatrix(0, 1) = "No."
        .TextMatrix(1, 1) = .TextMatrix(0, 1)
        .TextMatrix(2, 1) = .TextMatrix(0, 1)
        .MergeCol(1) = True
        .ColWidth(1) = 500
        
        .TextMatrix(0, 2) = "Part Number"
        .TextMatrix(1, 2) = .TextMatrix(0, 2)
        .TextMatrix(2, 2) = .TextMatrix(0, 2)
        .MergeCol(2) = True
        .ColWidth(2) = 3000
        .ColAlignment(2) = flexAlignLeftCenter
        
        .TextMatrix(0, 3) = "Part Name"
        .TextMatrix(1, 3) = .TextMatrix(0, 3)
        .TextMatrix(2, 3) = .TextMatrix(0, 3)
        .MergeCol(3) = True
        .ColWidth(3) = 3000
        
        .TextMatrix(0, 4) = "Process Name"
        .TextMatrix(1, 4) = .TextMatrix(0, 4)
        .TextMatrix(2, 4) = .TextMatrix(0, 4)
        .MergeCol(4) = True
        
        .TextMatrix(0, 5) = "Total Cavity/tool"
        .TextMatrix(1, 5) = .TextMatrix(0, 5)
        .TextMatrix(2, 5) = "a"
        .MergeCol(5) = True
        
        .TextMatrix(0, 6) = "Efficiency (%)"
        .TextMatrix(1, 6) = .TextMatrix(0, 6)
        .TextMatrix(2, 6) = "b"
        .MergeCol(6) = True
        
        .TextMatrix(0, 7) = "CT/(detik)"
        .TextMatrix(1, 7) = .TextMatrix(0, 7)
        .TextMatrix(2, 7) = "c"
        .MergeCol(7) = True
        
        .TextMatrix(0, 8) = "Loading"
        .TextMatrix(1, 8) = "Forecast Tertinggi (pcs)"
        .TextMatrix(2, 8) = "d"
        .MergeCol(8) = True
        .RowHeight(1) = 900
        
        .TextMatrix(0, 9) = "Loading"
        .TextMatrix(1, 9) = "Waktu dibutuhkan (detik)"
        .TextMatrix(2, 9) = "e = c*d*(2-b)/a"
        .MergeCol(9) = True
        .MergeRow(0) = True
        .ColWidth(9) = 1800
        
        .TextMatrix(0, 10) = "Waktu Kerja (kapasitas) Tersedia"
        .TextMatrix(1, 10) = "Waktu Kerja Per hari (Jam)"
        .TextMatrix(2, 10) = "f"
        .MergeCol(10) = True
       
        
        .TextMatrix(0, 11) = "Waktu Kerja (kapasitas) Tersedia"
        .TextMatrix(1, 11) = "Hari kerja dalam satu bulan"
        .TextMatrix(2, 11) = "g"
        .MergeCol(11) = True
        
                
        .TextMatrix(0, 12) = "Waktu Kerja (kapasitas) Tersedia"
        .TextMatrix(1, 12) = "Hari Kerja Per bulan (detik)"
        .TextMatrix(2, 12) = "h = 3600*f*g"
        .MergeCol(12) = True
        
        .TextMatrix(0, 13) = ""
        .TextMatrix(1, 13) = "% Loading terhadap kapasitas"
        .TextMatrix(2, 13) = "i = e/g"
        .MergeCol(13) = True
        
    End With
End Sub


Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(Skin1, Me)
    Call settingFG
    Call loadMachnine
    Me.Height = 7020
    Me.Width = 13440
    cmbMachine.ListIndex = 0
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    txtRevision.Left = CmbDocument.Left
    txtRevision.Top = SkinLabel2.Top
    cmbPeriod.Top = SkinLabel1.Top
    cmbPeriod.Left = lbl_hkw.Left
    
    cmbMachine.Left = txtCustomer.Left
    cmbMachine.Top = SkinLabel5.Top
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
