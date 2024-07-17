VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Insertion 
   Caption         =   "Insertion"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14220
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
   ScaleHeight     =   5730
   ScaleWidth      =   14220
   Begin VB.CommandButton cmdInit 
      Caption         =   "Initialize"
      Height          =   375
      Left            =   11400
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current MPS Documents"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   14055
      Begin VB.TextBox txtPeriod 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtLTPPDocRev 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtLTPPDoc 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtMPSrev 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtMPSDoc 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "Form_Insertion.frx":0000
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "Form_Insertion.frx":0078
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   6240
         OleObjectBlob   =   "Form_Insertion.frx":00F2
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   6240
         OleObjectBlob   =   "Form_Insertion.frx":016C
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "Form_Insertion.frx":01E8
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSync 
      Caption         =   "Sync"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12960
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_Insertion.frx":024C
      TabIndex        =   1
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_Insertion.frx":02AC
      Top             =   0
   End
   Begin MSComctlLib.ListView LV2 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form_Insertion"
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
Dim rsSync As ADODB.Recordset
Dim lisitm As ListItem
Dim in_Itemid As String
Dim in_machine As String
Dim in_mold As String
Dim in_customer As String
Dim in_tonage As String
Dim in_period As String
Dim in_ItemNm As String
Dim in_ltppdoc As String
Dim in_menuload_rev As Byte
Dim in_cvt As Byte
Dim in_ct As Single
Dim in_ct2 As Single
Dim in_manpower As Byte
Dim in_cappday As Long
Dim hkw As String
Dim in_factor As Single
Dim in_stok As Long
Dim in_stokwip As Long
Dim in_fc As Long
Dim in_ito As Single
Dim in_shiftUsage As Byte
Dim in_hourpshift As Single
Dim in_neqty As Single
Dim in_cavStd As Byte
Dim in_Timeupdate   As String
Dim in_prcProdplan As Single
Dim in_rstatemach As String


Private Sub cmdInit_Click()
    qry = "SELECT sum_nedady FROM mpp_gen_d " _
    & " WHERE fltpp_doc='" & txtLTPPDoc & "' AND fltpp_rev=" & txtLTPPDocRev & "" _
    & " and no_mach='" & in_machine & "' and lcd_itemdid='" & in_Itemid & "'" _
    & " and reg_mold='" & in_mold & "' limit 1"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        MsgBox "the data is already exists"
        Exit Sub
    End If
    Set RsBantu = Nothing
    
    qry = "SELECT lc_fprodtvty,lc_stockqty,lc_stockwip,lc_fc,lc_ito,shiftusg, " _
    & "hourpshift,neqty,prc_prodplan,rstate_mach,fltpp_hkw FROM mpp_gen_d " _
    & " WHERE fltpp_doc='" & txtLTPPDoc & "' AND fltpp_rev=" & txtLTPPDocRev & "" _
    & " and no_mach='" & in_machine & "' and lcd_itemdid='" & in_Itemid & "'" _
    & " limit 1"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        hkw = RsBantu("fltpp_hkw")
        in_factor = RsBantu("lc_fprodtvty")
        in_stok = RsBantu("lc_stockqty")
        in_stokwip = RsBantu("lc_stockwip")
        in_fc = RsBantu("lc_fc")
        in_ito = RsBantu("lc_ito")
        in_shiftUsage = RsBantu("shiftusg")
        in_hourpshift = RsBantu("hourpshift")
        in_prcProdplan = RsBantu("prc_prodplan")
        in_rstatemach = RsBantu("rstate_mach")
    Else
        qry = "select * from mpp_gen_d where fltpp_doc='" & txtLTPPDoc & "' and fltpp_rev='" & txtLTPPDocRev & "'"
        Set RsTemp = Con.Execute(qry)
        hkw = RsTemp("fltpp_hkw")
        in_factor = RsTemp("lc_fprodtvty")
        in_stok = 0
        in_stokwip = 0
        in_fc = 0
        in_ito = 0
        in_shiftUsage = RsTemp("shiftusg")
        in_hourpshift = RsTemp("hourpshift")
        in_prcProdplan = RsTemp("prc_prodplan")
        in_rstatemach = RsTemp("rstate_mach")
    End If
    Set RsBantu = Nothing
End Sub

Private Sub cmdSync_Click()
On Error GoTo ExE
    Dim rsLC_d As ADODB.Recordset
    
    '==============PERIKSA DATA UNIK================='
    qry = "select count(*) mpp_gen_d where lcd_itemdid='" & in_Itemid & "' and " _
    & " no_mach='" & in_machine & "' and no_mach='" & in_machine & "' and " _
    & " fltpp_doc='" & in_ltppdoc & "' and fltpp_rev=" & in_menuload_rev
    Set RsBantu = Con.Execute(qry)
    If RsBantu(0) > 0 Then
        MsgBox "The data is already exist, the process was canceled"
        Exit Sub
    End If

    Set rsLC_d = New ADODB.Recordset
    rsLC_d.Open "select * from mpp_gen_d where fltpp_doc='" & txtLTPPDoc & "' and fltpp_rev = " & txtLTPPDocRev & " limit 1", Con, adOpenDynamic, adLockOptimistic
    rsLC_d.AddNew
    rsLC_d!lcd_itemdid = in_Itemid 'm
    rsLC_d!lc_customer = in_customer 'm
    rsLC_d!lc_subcont = "no" 'm
    rsLC_d!no_mach = in_machine 'm
    rsLC_d!ton_mach = Val(in_tonage) 'm
    rsLC_d!reg_mold = in_mold 'm
    rsLC_d!fltpp_period = in_period 'm
    rsLC_d!cav = in_cvt 'm
    rsLC_d!ct = in_ct 'm
    rsLC_d!mpower = in_manpower 'm
    rsLC_d!ct_scnd = in_ct2 'm
    rsLC_d!cap_p_day = in_cappday 'm
    rsLC_d!neday = 0
    rsLC_d!sum_nedady = 0
    rsLC_d!lcvsmach = 0
    rsLC_d!lcneed_mp = 0
    rsLC_d!fltpp_doc = txtLTPPDoc
    rsLC_d!fltpp_rev = txtLTPPDocRev
    rsLC_d!fltpp_ym = txtPeriod
    rsLC_d!rstate_mach = in_rstatemach
    rsLC_d!fltpp_hkw = hkw
    rsLC_d!lc_sisa_pp = 0
    rsLC_d!lc_pp = 0
    rsLC_d!lc_fprodtvty = in_factor
    rsLC_d!lc_itemname = in_ItemNm
    rsLC_d!lc_stockqty = in_stok
    rsLC_d!lc_stockwip = in_stokwip
    rsLC_d!lc_fc = in_fc
    rsLC_d!lc_ito = in_ito
    rsLC_d!shiftusg = in_shiftUsage
    rsLC_d!hourpshift = in_hourpshift
    rsLC_d!neqty = 0
    rsLC_d!cav_std = in_cavStd
    If in_Timeupdate <> "" Then
        rsLC_d!timeupdate = in_Timeupdate
    End If
    rsLC_d!prc_prodplan = in_prcProdplan
    rsLC_d.Update
    MsgBox "Saved successfully"
    Exit Sub
ExE:
    MsgBox Err.Description, vbCritical
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

Private Sub loadDat()
    qry = "select * from (select distinct on (mpp_doc_no) mpp_doc_no,mpp_revisi,ml_ym,ml_rev,ml_doc  from mpp_gen ) v1 order by ml_ym desc, mpp_revisi desc limit 1"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        in_period = RsBantu("ml_ym")
        in_ltppdoc = RsBantu("ml_doc")
        in_menuload_rev = RsBantu("ml_rev")
        txtMPSDoc = RsBantu("mpp_doc_no")
        txtMPSrev = RsBantu("mpp_revisi")
        txtPeriod = RsBantu("ml_ym")
        txtLTPPDocRev = RsBantu("ml_rev")
        txtLTPPDoc = RsBantu("ml_doc")
    End If
    Set RsBantu = Nothing
    

    
    qry = "select idproclc,a.partno,partname,manpower,prod_nomach,mold_no,a.cavity, " _
    & " cavity_std,ct,ct_2,priorit,cust_name,tonage_mach,hour_p_shift," _
    & " shift_usg,faktor_productivity,item_muloq,item_perbox,timeupdate" _
    & " from loadcap_proc a inner join loadcap_mst_product_r b on a.partno=b.partno" _
    & " inner join mst_item c on a.partno=c.item_id" _
    & " inner join r_customer d on c.cust_id=d.cust_id" _
    & " inner join loadcap_mst_mach e on a.prod_nomach=e.no_mach" _
    & " where subcont='no' order by cust_name asc,a.partno asc,priorit asc"
    Set rsSync = Con.Execute(qry)

End Sub

Private Sub getList()
On Error Resume Next
    LV2.ListItems.Clear
    Do Until rsSync.EOF
        If rsSync("ct") = 0 Then
            in_cappday = 0
        Else
            in_cappday = ((60 / rsSync("ct")) * rsSync("cavity") * rsSync("hour_p_shift") * rsSync("shift_usg") * 60) * rsSync("faktor_productivity")
            If rsSync("item_perbox") = 0 Then
                in_cappday = isi(rsSync("item_muloq"), in_cappday, "b")
            Else
                in_cappday = isi(rsSync("item_perbox"), in_cappday, "b")
            End If
        End If
        Set lisitm = LV2.ListItems.Add(, , RTrim(rsSync!idproclc))
            lisitm.SubItems(1) = rsSync!cust_name
            lisitm.SubItems(2) = RTrim(rsSync!partNo)
            lisitm.SubItems(3) = rsSync!partname
            lisitm.SubItems(4) = IIf(IsNull(rsSync!manPower), 0, rsSync!manPower)
            lisitm.SubItems(5) = IIf(IsNull(rsSync!prod_nomach), "", RTrim(rsSync!prod_nomach))
            lisitm.SubItems(6) = rsSync!tonage_mach
            lisitm.SubItems(7) = rsSync!mold_no
            lisitm.SubItems(8) = rsSync!cavity
            lisitm.SubItems(9) = IIf(IsNull(rsSync!cavity_std), 0, rsSync!cavity_std)
            lisitm.SubItems(10) = rsSync!ct
            lisitm.SubItems(11) = rsSync!ct_2
            lisitm.SubItems(12) = rsSync!priorit
            lisitm.SubItems(13) = in_cappday
            lisitm.SubItems(14) = Format(rsSync!timeupdate, "yyyy-MM-dd HH:mm:ss")
        rsSync.MoveNext
    Loop
End Sub

Private Function isi(pMPQ As Double, pCapPDay As Variant, atasBawah As String)
    Dim bReach As Boolean
    Dim MPQ As Long
    bReach = True
    MPQ = pMPQ
    While bReach
        If MPQ * 1 > pCapPDay * 1 Then
            If atasBawah = "a" Then
                isi = MPQ '- pMPQ
            Else
                isi = MPQ - pMPQ
            End If
            bReach = False
        Else
            If MPQ = pCapPDay Then
                isi = pCapPDay
                bReach = False
            Else
                isi = MPQ
            End If
        End If
        MPQ = MPQ * 1 + pMPQ * 1
    Wend
End Function

Private Sub Form_Load()
 On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingFG
    loadDat
    Me.Height = 6300
    Me.Width = 14460
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
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

Private Sub settingFG()
    With LV2
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "id", 0
        .ColumnHeaders.Add , , "Customer", 3000
        .ColumnHeaders.Add , , "Part No", 2700
        .ColumnHeaders.Add , , "Part Name", 2700
        .ColumnHeaders.Add , , "Man Power", 900
        .ColumnHeaders.Add , , "Machine No"
        .ColumnHeaders.Add , , "Tonage", 900
        .ColumnHeaders.Add , , "Mold No"
        .ColumnHeaders.Add , , "Cavity", 700
        .ColumnHeaders.Add , , "Cavity STD", 1000
        .ColumnHeaders.Add , , "Cycle Time (CT)"
        .ColumnHeaders.Add , , "CT 2nd", 1000
        .ColumnHeaders.Add , , "Priority", 1000
        .ColumnHeaders.Add , , "Capacity /day", 3000
        .ColumnHeaders.Add , , "timeupdate", 0
    End With
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Set rsSync = Nothing
End Sub

Private Sub LV2_Click()
    lvToVar
'    MsgBox in_Itemid & " " & in_customer & " " & in_ItemNm
End Sub

Private Sub lvToVar()
    in_customer = LV2.SelectedItem.SubItems(1)
    in_Itemid = LV2.SelectedItem.SubItems(2)
    in_ItemNm = LV2.SelectedItem.SubItems(3)
    in_manpower = LV2.SelectedItem.SubItems(4)
    in_machine = LV2.SelectedItem.SubItems(5)
    in_tonage = LV2.SelectedItem.SubItems(6)
    in_mold = LV2.SelectedItem.SubItems(7)
    in_cvt = LV2.SelectedItem.SubItems(8)
    in_cavStd = LV2.SelectedItem.SubItems(9)
    in_ct = LV2.SelectedItem.SubItems(10)
    in_ct2 = LV2.SelectedItem.SubItems(11)
    in_cappday = LV2.SelectedItem.SubItems(13)
    in_Timeupdate = LV2.SelectedItem.SubItems(14)
End Sub

Private Sub LV2_KeyUp(KeyCode As Integer, Shift As Integer)
    lvToVar
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtfind = FilterIn(txtfind)
        If Len(Trim(txtfind)) > 0 Then
            rsSync.Filter = "partno LIKE '*" & txtfind & "*'"
        Else
            rsSync.Filter = adFilterNone
        End If
        If rsSync.RecordCount > 0 Then
            Call getList
        Else
            rsSync.Filter = adFilterNone
            rsSync.Filter = "partname LIKE '*" & txtfind & "*'"
            Call getList
        End If
    End If
End Sub



