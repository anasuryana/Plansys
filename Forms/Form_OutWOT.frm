VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_OutWOT 
   Caption         =   "Out WOT"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   18240
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.Label TYPEKANBAN 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Out WOT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   30
         Width           =   6135
      End
   End
   Begin VB.PictureBox picBAR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -120
      ScaleHeight     =   330
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   15075
      TabIndex        =   5
      Top             =   360
      Width           =   15135
      Begin VB.PictureBox picStatus 
         Height          =   1335
         Left            =   13320
         ScaleHeight     =   1275
         ScaleWidth      =   1395
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtSerial 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   56.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1350
         Left            =   3240
         MaxLength       =   16
         TabIndex        =   7
         Top             =   1320
         Width           =   9975
      End
      Begin VB.TextBox txtKanbanId 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   750
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "WOT ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   18015
      Begin MSComctlLib.ListView lvDataOut 
         Height          =   4695
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   17775
         _ExtentX        =   31353
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Part No"
            Object.Width           =   5717
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Part Name"
            Object.Width           =   5186
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Req Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Scan Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Suggest"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Scan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   15360
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   375
      Left            =   13080
      OleObjectBlob   =   "Form_OutWOT.frx":0000
      TabIndex        =   15
      Top             =   120
      Width           =   135
   End
   Begin MSComctlLib.ImageList imgListStatus 
      Left            =   840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   95
      ImageHeight     =   87
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_OutWOT.frx":0058
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_OutWOT.frx":2E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_OutWOT.frx":675A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_OutWOT.frx":A0A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_OutWOT.frx":D16F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_OutWOT.frx":DEB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin AriefIn 
      Left            =   240
      OleObjectBlob   =   "Form_OutWOT.frx":10DEF
      Top             =   120
   End
   Begin VB.PictureBox logoBPI150 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      Picture         =   "Form_OutWOT.frx":11023
      ScaleHeight     =   1815
      ScaleWidth      =   1935
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form_OutWOT"
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
Dim listOut     As listItem
Dim rs_Serial   As New ADODB.Recordset
Dim rs_Ith      As New ADODB.Recordset
Dim rs_IthQty   As New ADODB.Recordset
Dim affSerial   As Byte
Dim affIth      As Byte
Dim sDoc        As String
Dim sItem       As String
Dim qry         As String
Dim adiTanggal  As String
Dim sQty        As Integer
Dim sMpq        As Integer
Dim indexLV     As Byte
Dim balBEf      As Single
Dim c_datetx    As Date
Dim AssyNo      As String
Dim stsNOTCLOSED As Boolean

Private Sub cmdnew_Click()
    txtKanbanId = ""
    txtKanbanId.Locked = False
    txtKanbanId.SetFocus
    lvDataOut.ListItems.Clear
    picStatus.Cls
    Set picStatus.Picture = Nothing
    picStatus.ToolTipText = ""
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Sub ResizeControls()
    On Error Resume Next
    Dim i As Byte
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

Private Sub Form_Initialize()
    Me.WindowState = vbNormal
    On Error Resume Next
    Dim i As Byte
    ReDim proportionArray(0 To Controls.Count - 1)
    
    For i = 0 To Controls.Count - 1
         With proportionArray(i)
            .Heightproportion = Controls(i).Height / ScaleHeight
            .WidthProportion = Controls(i).Width / ScaleWidth
            .TopProportion = Controls(i).Top / ScaleHeight
            .LeftProportion = Controls(i).Left / ScaleWidth
         End With
    Next
    If Me.WindowState <> vbMaximized Then
    Me.WindowState = vbMaximized
    End If
End Sub

Private Sub Form_Load()
    AddTab Me
    Call TemaAktif(AriefIn, Me)
    adiTanggal = Format(Now, "yyyy-MM-dd")
    BukaKoneksi_alt
    Me.Width = 18480
    Me.Height = 9045
    
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Function querySO(SOD As String) As String
'    querySO = "select a.item_id,item_name,qty,cl,datetx,assy_no, " _
    & " item_muloq,coalesce(mat_act,0) mat_act,(qty*cl)/1000 material_qty, " _
    & " coalesce(balc_bef,0) balc_bef from serial_barcode_cutting_tube a " _
    & " inner join mst_item b on a.item_id=b.item_id  " _
    & " left join bal_wot c on a.serialno=c.docno and a.item_id=c.mat_id" _
    & " where serialno ='" & SOD & "'"
    querySO = "select a.item_id,item_name,qty,cl,assy_no, " _
    & " item_muloq,(qty*cl)/1000 material_qty " _
    & " from serial_barcode_cutting_tube a " _
    & " inner join mst_item b on a.item_id=b.item_id  " _
    & " where serialno ='" & SOD & "'"
End Function

Private Function fetchIthQty(parItem As String) As Long
    rs_IthQty.Filter = adFilterNone
    If rs_IthQty.RecordCount > 0 Then
        rs_IthQty.Filter = "ith_item_id='" & parItem & "'"
        If rs_IthQty.RecordCount > 0 Then
            fetchIthQty = rs_IthQty("qtyin")
        Else
            fetchIthQty = 0
        End If
    Else
        fetchIthQty = 0
    End If
End Function

Private Sub txtKanbanId_KeyPress(KeyAscii As Integer)
On Error GoTo errScan
    If KeyAscii = 13 Then
        txtKanbanId = Trim(txtKanbanId)
        txtKanbanId = FilterIn(txtKanbanId)
        sDoc = txtKanbanId
        Set RsGet = Con_alt.Execute(querySO(txtKanbanId))
        Set rs_IthQty = Con.Execute("select ith_item_id,abs(coalesce(ith_qty,0)) qtyin from ith where ith_docno='" & txtKanbanId & "' ")
        Call getList
        
'        If stsNOTCLOSED = False Then
'            keMaterial
'        End If
    End If
Exit Sub
errScan:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error: " & Err.Number
End Sub

Private Function isi(pMPQ As Integer, initVal As Single, atasBawah As String) As Integer
    Dim bReach As Boolean
    Dim MPQ As Integer
    bReach = True
    MPQ = pMPQ
    While bReach
        If MPQ * 1 > initVal * 1 Then
            If atasBawah = "a" Then
                isi = MPQ '- pMPQ
            Else
                isi = MPQ - pMPQ
            End If
            bReach = False
        Else
            If MPQ = initVal Then
                isi = initVal
                bReach = False
            Else
                isi = MPQ
            End If
        End If
        MPQ = MPQ * 1 + pMPQ * 1
    Wend
End Function

Private Function getBalcLatest(pitem As String, pmaterial As String) As Double
    Dim detik As Single
    Dim rsBalc As ADODB.Recordset
    c_datetx = DateAdd("s", -1, c_datetx)
    
    detik = Format(c_datetx, "ss")
    If detik > 0 Then
        detik = detik - 1
    End If
    qry = "select balc_aft from bal_wot where mat_id='" & pmaterial & "' " _
    & " and datetx<'" & Format(c_datetx, "yyyy\-MM\-dd hh:mm:ss") & "'" _
    & " order by datetx desc limit 1 "
    
    Set rsBalc = Con.Execute(qry)
    If rsBalc.RecordCount > 0 Then
        getBalcLatest = rsBalc("balc_aft")
        'MsgBox rsBalc("balc_aft") & "woo"
    Else
        getBalcLatest = 0
    End If
    Set rsBalc = Nothing
End Function

Private Sub getList()
    Dim suggestQTY As Single
    Dim balcCur As Single
    Dim REqQty As Single
    Dim ithQty As Single
    lvDataOut.ListItems.Clear
    If Not RsGet.EOF Then
'        If IsNull(RsGet("datetx")) Then
'            c_datetx = Now
'        Else
'            c_datetx = RsGet("datetx")
'        End If
        AssyNo = RsGet("assy_no")
        txtKanbanId.Locked = True
        txtSerial.Enabled = True
        txtSerial.SetFocus
        RsGet.MoveFirst
        Do Until RsGet.EOF
            REqQty = (RsGet!qty * RsGet!cl) / 1000
            ithQty = fetchIthQty(RTrim$(RsGet!item_id)) 'RsGet!qtyin
'            If (RsGet("mat_act") + RsGet("balc_bef")) - RsGet("material_qty") >= 0 Then
'                balBEf = getBalcLatest(AssyNo, RTrim$(RsGet!item_id))
'            Else
'                balBEf = getBalcLatest_v2(AssyNo, RTrim$(RsGet!item_id))
'            End If
'            suggestQTY = RsGet("material_qty") - balBEf
'            If suggestQTY > 0 Then
'                suggestQTY = isi(RsGet("item_muloq"), suggestQTY, "a")
'            Else
'                suggestQTY = 0
'            End If
'            balcCur = balBEf - (REqQty - ithQty)
            suggestQTY = RsGet("material_qty")
            suggestQTY = isi(RsGet("item_muloq"), suggestQTY, "a")
            Set listOut = lvDataOut.ListItems.Add(, , lvDataOut.ListItems.Count + 1)
                listOut.SubItems(1) = RTrim$(RsGet!item_id)
                listOut.SubItems(2) = RTrim$(RsGet!item_name)
                listOut.SubItems(3) = REqQty
                listOut.SubItems(4) = ithQty
                listOut.SubItems(5) = suggestQTY
'                listOut.SubItems(6) = balBEf
'                listOut.SubItems(7) = balcCur
            RsGet.MoveNext
        Loop
        RsGet.Close
    Else
        txtKanbanId = ""
    End If
End Sub

'Private Sub SettingFG()
'    Dim i As Byte
'    With grid1
'        .Cols = 8
'        .Rows = 3
'        .FixedRows = 2
'
'        i = 0
'        .TextMatrix(0, i) = "No."
'        .TextMatrix(1, i) = .TextMatrix(0, i)
'        .ColWidth(i) = 500
'        i = 1
'        .TextMatrix(0, i) = "Part No"
'        .TextMatrix(1, i) = .TextMatrix(0, i)
'        .ColWidth(i) = 3000
'        i = 2
'        .TextMatrix(0, i) = "Part Name"
'        .TextMatrix(1, i) = .TextMatrix(0, i)
'        .ColWidth(i) = 3050
'        i = 3
'        .TextMatrix(0, i) = "Req Qty"
'        .TextMatrix(1, i) = .TextMatrix(0, i)
'        i = 4
'        .TextMatrix(0, i) = "Scan Qty"
'        .TextMatrix(1, i) = .TextMatrix(0, i)
'        i = 5
'        .TextMatrix(0, i) = "Suggest"
'        .TextMatrix(1, i) = .TextMatrix(0, i)
'        i = 6
'        .TextMatrix(0, i) = "Balance WOT Cutting"
'        .TextMatrix(1, i) = "Before"
'        i = 7
'        .TextMatrix(0, i) = "Balance WOT Cutting"
'        .TextMatrix(1, i) = "After"
'    End With
'End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
On Error GoTo errScan
    If KeyAscii = 13 Then
        txtSerial = FilterIn(txtSerial)
        Set rs_Serial = Con.Execute("select a.item_id,qty,item_muloq,sts_out from serial_barcode_packing a inner join mst_item b on a.item_id=b.item_id where serialno='" & txtSerial & "' and sts_in = true")
        If rs_Serial.RecordCount > 0 Then
            rs_Serial.MoveFirst
            sItem = Trim(rs_Serial!item_id)
            sQty = rs_Serial("qty")
            sMpq = rs_Serial("item_muloq")
           
            If rs_Serial!sts_out = 0 Then
                getIndexLV sItem
                If indexLV = 0 Then txtSerial = "": Exit Sub
                
                'jika angka scan dan kelipatan sama
                If sQty + lvDataOut.ListItems(indexLV).SubItems(4) * 1 = isi(sMpq, lvDataOut.ListItems(indexLV).SubItems(3), "a") Then
                    Con.Execute "UPDATE serial_barcode_packing SET scan_out = now(), sts_out = true, user_scanout = '" & getNameUser & "',doc_out = '" & txtKanbanId & "' WHERE serialno = '" & txtSerial & "'", affSerial
                    Set rs_Ith = Con.Execute("SELECT * FROM ith WHERE ith_item_id = '" & RTrim(sItem) & "' AND ith_docno = '" & txtKanbanId & "' AND ith_form = 'OU - WOT'")
                    If rs_Ith.RecordCount = 0 Then
                        'insert to ith
                        qry = "INSERT INTO ith (ith_id, ith_date, ith_form, ith_docno, ith_qtybf, ith_qty, ith_qtyend, ith_item_id, employee) " _
                            & "SELECT (select coalesce(max(ith_id), 0) + 1 from ith where ith_item_id = '" & sItem & "' and ith_date = '" & adiTanggal & "') ith_id, " _
                            & "'" & adiTanggal & "' ith_date, 'OU - WOT' ith_form, '" & txtKanbanId & "' ith_docno, " _
                            & "coalesce(sum(ith_qty), 0) ith_qtybf, -" & sQty & " ith_qty, coalesce(sum(ith_qty), 0) - " & sQty & " ith_qtyend, '" & sItem & "' ith_item_id, '" & getNameUser & "' employee " _
                            & "FROM ith WHERE ith_item_id = '" & sItem & "' AND ith_date <= '" & adiTanggal & "'"
                        Con.Execute qry, affIth
                    Else
                        'update ith
                        qry = "UPDATE ith SET ith_qty = ith_qty - " & sQty & ", ith_qtyend = ith_qtyend - " & sQty & " WHERE ith_docno = '" & txtKanbanId & "' AND ith_form = 'OU - WOT' " _
                            & "AND ith_item_id = '" & sItem & "' AND ith_date = '" & Format(rs_Ith!ith_date, "yyyy-MM-dd") & "' AND ith_id = " & rs_Ith!ith_id
                        Con.Execute qry, affIth
                    End If
                    txtSerial.text = ""
                    If affSerial > 0 Then
                        Set picStatus.Picture = imgListStatus.ListImages(1).Picture 'status scan ok
                    End If
                    Call PlayWaveSoundOK
                    txtKanbanId_KeyPress 13
                Else
                    'MsgBox "broo"
                    If sQty + lvDataOut.ListItems(indexLV).SubItems(4) * 1 < isi(sMpq, lvDataOut.ListItems(indexLV).SubItems(3), "a") Then
                        'MsgBox "waa"
                        Con.Execute "UPDATE serial_barcode_packing SET scan_out = now(), sts_out = true, user_scanin = '" & getNameUser & "',doc_out = '" & txtKanbanId & "' WHERE serialno = '" & txtSerial & "'", affSerial
                        Set rs_Ith = Con.Execute("SELECT * FROM ith WHERE ith_item_id = '" & RTrim(sItem) & "' AND ith_docno = '" & sDoc & "' AND ith_form = 'OU - WOT'")
                        'MsgBox rs_Ith.RecordCount & "nanooo"
                        If rs_Ith.RecordCount = 0 Then
                            'insert to ith
                            qry = "INSERT INTO ith (ith_id, ith_date, ith_form, ith_docno, ith_qtybf, ith_qty, ith_qtyend, ith_item_id, employee) " _
                                & "SELECT (select coalesce(max(ith_id), 0) + 1 from ith where ith_item_id = '" & sItem & "' and ith_date = '" & adiTanggal & "') ith_id, " _
                                & "'" & adiTanggal & "' ith_date, 'OU - WOT' ith_form, '" & sDoc & "' ith_docno, " _
                                & "coalesce(sum(ith_qty), 0) ith_qtybf, -" & sQty & " ith_qty, coalesce(sum(ith_qty), 0) - " & sQty & " ith_qtyend, '" & sItem & "' ith_item_id, '" & getNameUser & "' employee " _
                                & "FROM ith WHERE ith_item_id = '" & sItem & "' AND ith_date <= '" & adiTanggal & "'"
                            Con.Execute qry, affIth
                        Else
                            'update ith
                            MsgBox rs_Ith.RecordCount
                            qry = "UPDATE ith SET ith_qty = ith_qty - " & sQty & ", ith_qtyend = ith_qtyend - " & sQty & " WHERE ith_docno = '" & sDoc & "' AND ith_form = 'OU - WOT' " _
                                & "AND ith_item_id = '" & sItem & "' AND ith_date = '" & Format(rs_Ith!ith_date, "yyyy-MM-dd") & "' AND ith_id = " & rs_Ith!ith_id
                            Con.Execute qry, affIth
                        End If
                        
                        txtSerial.text = ""
                        If affSerial > 0 Then
                            Set picStatus.Picture = imgListStatus.ListImages(1).Picture 'status scan ok
                        End If
                        Call PlayWaveSoundOK
                    Else
                        'qty scan melebihi batas qty pada kanban
                        Call PlayWaveSoundERROR
                        txtSerial.text = ""
                        picStatus.ToolTipText = "Scanned QTY > REQ QTY"
                        Set picStatus.Picture = imgListStatus.ListImages(5).Picture
                        txtSerial.SetFocus
                    End If
                End If
                txtKanbanId_KeyPress 13
            Else
                'serial sudah di scan out
                Call PlayWaveSoundDUPLICATE
                txtSerial.text = ""
                picStatus.ToolTipText = "The serial number is already scanned"
                Set picStatus.Picture = imgListStatus.ListImages(4).Picture
                txtSerial.SetFocus
            End If
        Else
             'serial tidak ada atau belum di scan in
            txtSerial.text = ""
            Set picStatus.Picture = imgListStatus.ListImages(6).Picture
            Call PlayWaveSoundNOTRECEIVE
            txtSerial.SetFocus
        End If
    End If
Exit Sub
errScan:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error: " & Err.Number
End Sub

Private Sub getIndexLV(par_item As String)
    Dim g As Byte
    indexLV = 0
    For g = 1 To lvDataOut.ListItems.Count
'        MsgBox lvDataOut.ListItems(g).SubItems(1) & "=" & par_item
        If lvDataOut.ListItems(g).SubItems(1) = par_item Then
           
            indexLV = g
            Exit For
        End If
    Next
End Sub

Private Function getReqQty(par_item As String) As Double
    Dim b As Byte
    For b = 1 To lvDataOut.ListItems.Count
        If par_item = lvDataOut.ListItems(b).SubItems(1) Then
            getReqQty = lvDataOut.ListItems(b).SubItems(3)
            Exit For
        End If
    Next
End Function


'Private Sub keMaterial()
'On Error GoTo exCep
'    Dim balc As Single
'    Dim balcCurnt As Double
'    Dim rowAFF As Byte
'    Dim rsBalc As ADODB.Recordset
'    Dim i As Byte
'    With lvDataOut
'        If IsNumeric(.ListItems(1).SubItems(5)) Then
'            balc = .ListItems(1).SubItems(5)
'        Else
'            balc = 0
'        End If
''        QRY = "update kanban_material set kanbanscin=now(),scin_usr='" & pUserName & "'" _
''        & " where kanbanid='" & txtKanbanId & "'"
''        Con.Execute QRY, rowAFF
''        If rowAFF > 0 Then
'            For i = 1 To .ListItems.Count
'                qry = "select * from bal_wot where docno='" & txtKanbanId & "' and partno='" & .ListItems(i).SubItems(3) & "' and mat_id='" & .ListItems(i).SubItems(1) & "'"
'
'                Set rsBalc = Con.Execute(qry)
'                If IsNumeric(.ListItems(i).SubItems(5)) And IsNumeric(.ListItems(i).SubItems(4)) And IsNumeric(.ListItems(i).SubItems(3)) Then
'                    balcCurnt = .ListItems(i).SubItems(5) * 1 - (-(.ListItems(i).SubItems(4) * 1)) - .ListItems(i).SubItems(3) * 1
'                End If
'
'                If rsBalc.RecordCount > 0 Then
'                    qry = "update bal_mat set balc_bef=" & .ListItems(i).SubItems(5) * 1 & ", mat_act=-" & .ListItems(i).SubItems(4) * 1 & ",balc_aft=" & balcCurnt & ", mat_sug=" & .ListItems(i).SubItems(5) * 1 & "" _
'                    & " where docno='" & txtKanbanId.text & "' and partno='" & .TextMatrix(i, 0) & "' and mat_id='" & .TextMatrix(i, 3) & "'"
'                    Con.Execute qry
'                Else
'                    qry = "insert into bal_mat values(now(),'" & txtKanbanId & "','" & .TextMatrix(i, 0) & "'," & .TextMatrix(i, 4) * 1 & "," _
'                    & .TextMatrix(i, 5) * 1 & "," & .TextMatrix(i, 6) * 1 & ",-" & .TextMatrix(i, 7) * 1 & "," & balcCurnt & ",'" & .TextMatrix(i, 3) & "')"
'                    Con.Execute qry
'                End If
'            Next
''        End If
'    End With
'    Exit Sub
'exCep:
'    MsgBox Err.Description & "<" & Err.Number & ">", vbInformation, "Maaf: "
'End Sub

Private Function getBalcLatest_v2(pitem As String, pmaterial As String) As Double
    Dim rsBalc As ADODB.Recordset
    qry = "select balc_aft from bal_wot where mat_id='" & pmaterial & "' and docno!='" & txtKanbanId & "'" _
    & " order by datetx desc limit 1 "
    Set rsBalc = Con.Execute(qry)
    If rsBalc.RecordCount > 0 Then
        getBalcLatest_v2 = rsBalc("balc_aft")
    Else
        getBalcLatest_v2 = 0
    End If
End Function


