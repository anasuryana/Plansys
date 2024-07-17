VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_MasterSubcont 
   Caption         =   "Master Subcont"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
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
   ScaleHeight     =   6960
   ScaleWidth      =   6870
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtTotalMesin 
         Height          =   390
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   11
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveU 
         Caption         =   "Save"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtpriority 
         Height          =   390
         Left            =   1680
         TabIndex        =   12
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtCP 
         Height          =   390
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtAlamat 
         Height          =   615
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtNamaSubcont 
         Height          =   390
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtKode 
         Height          =   390
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   5
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Total Mesin"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Priority"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "CP"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Subcont"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Subcont"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3015
      Left            =   45
      TabIndex        =   0
      Top             =   3840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_MasterSubcont.frx":0000
      Top             =   4000
   End
End
Attribute VB_Name = "Form_MasterSubcont"
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
Dim lisitm As ListItem
Dim qry As String
Dim oldCode As String

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

Private Sub settingLV()
    With lv1
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Kode Subcont", 2500
        .ColumnHeaders.Add , , "Nama Subcont", 2000
        .ColumnHeaders.Add , , "Alamat", 5000
        .ColumnHeaders.Add , , "CP"
        .ColumnHeaders.Add , , "Total Mesin"
        .ColumnHeaders.Add , , "Prioritas", 800
    End With
End Sub

Private Sub lvtOform()
    If lv1.ListItems.Count > 0 Then
        oldCode = lv1.SelectedItem.Text
        txtKode = lv1.SelectedItem.Text
        txtNamaSubcont = lv1.SelectedItem.SubItems(1)
        txtAlamat = lv1.SelectedItem.SubItems(2)
        txtCP = lv1.SelectedItem.SubItems(3)
        txtTotalMesin = lv1.SelectedItem.SubItems(4)
        txtpriority = lv1.SelectedItem.SubItems(5)
    End If
End Sub

Private Sub kosong()
    txtKode = ""
    txtNamaSubcont = ""
    txtAlamat = ""
    txtCP = ""
    txtpriority = 0
End Sub

Private Sub cmdDelete_Click()
On Error GoTo errKu
    If txtKode = "" Then MsgBox "Tidak ada data yang terpilih": Exit Sub
    If MsgBox("Apakah anda yakin ingin menghapus " & vbNewLine & " data tersebut", vbQuestion + vbYesNo, "Tentukan") = vbYes Then
        qry = "delete from loadcap_mst_subcont where kodesubcont='" & txtKode & "'"
        Con.Execute qry
        MsgBox "deleted", vbInformation, "Penghapusan"
    End If
    Call LoadDatanya
    Call getList
    Exit Sub
errKu:
    MsgBox Err.Description, vbInformation, "Maaf"
End Sub

Private Sub cmdNew_Click()
    Call kosong
    txtKode.SetFocus
    cmdSaveU.Caption = "Save"
    cmdSaveU.Refresh
End Sub

Private Sub cmdSaveU_Click()
On Error GoTo errKu
    If IsNumeric(txtpriority) = False Then txtpriority.SetFocus: Exit Sub
    If IsNumeric(txtTotalMesin) = False Then txtTotalMesin.SetFocus: Exit Sub
    If cmdSaveU.Caption = "Save" Then
        If txtKode <> "" Then
            Set RsBantu = New ADODB.Recordset
            RsBantu.Open "loadcap_mst_subcont", Con, adOpenKeyset, adLockOptimistic, adCmdTable
            RsBantu.AddNew
            RsBantu!kodesubcont = txtKode
            RsBantu!namasubcont = txtNamaSubcont
            RsBantu!alamat = txtAlamat
            RsBantu!kontakperson = txtCP
            RsBantu!totalmesin_tersedia = txtTotalMesin
            RsBantu!prioritas = txtpriority
            RsBantu.Update
            RsBantu.Close
            Set RsBantu = Nothing
            MsgBox "Tersimpan", vbInformation, "Penyimpanan"
        End If
    Else
        If txtKode <> "" Then
            qry = "select kodesubcont,namasubcont,alamat,kontakperson,prioritas,totalmesin_tersedia from loadcap_mst_subcont " _
            & " where kodesubcont='" & oldCode & "'"
            RsBantu.Open qry, Con, adOpenKeyset, adLockOptimistic, adCmdText
            RsBantu!kodesubcont = txtKode
            RsBantu!namasubcont = txtNamaSubcont
            RsBantu!alamat = txtAlamat
            RsBantu!kontakperson = txtCP
            RsBantu!totalmesin_tersedia = txtTotalMesin
            RsBantu!prioritas = txtpriority
            RsBantu.Update
            RsBantu.Close
            Set RsBantu = Nothing
            MsgBox "Tersimpan", vbInformation, "Perubahan"
        End If
    End If
    Call LoadDatanya
    Call getList
    Exit Sub
errKu:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    settingLV
    Me.Height = 7530
    Me.Width = 7110
    Call LoadDatanya
    Call getList
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
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

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub getList()
    lv1.ListItems.Clear
    Do Until RsGet.EOF
        Set lisitm = lv1.ListItems.Add(, , RTrim(RsGet!kodesubcont))
            lisitm.SubItems(1) = RTrim(RsGet!namasubcont)
            lisitm.SubItems(2) = RTrim(RsGet!alamat)
            lisitm.SubItems(3) = RsGet!kontakperson
            lisitm.SubItems(4) = IIf(IsNull(RsGet!totalmesin_tersedia), 0, RsGet!totalmesin_tersedia)
            lisitm.SubItems(5) = IIf(IsNull(RsGet!prioritas), 0, RsGet!prioritas)
        RsGet.MoveNext
    Loop
End Sub

Private Sub LoadDatanya()
    Set RsGet = Con.Execute("select * from loadcap_mst_subcont order by prioritas asc")
End Sub

Private Sub lv1_Click()
    lvtOform
    cmdSaveU.Caption = "Save"
    cmdSaveU.Refresh
End Sub

Private Sub lv1_DblClick()
    cmdSaveU.Caption = "Update"
    cmdSaveU.Refresh
    txtKode.SetFocus
End Sub

Private Sub lv1_KeyUp(KeyCode As Integer, Shift As Integer)
    lvtOform
    cmdSaveU.Caption = "Save"
    cmdSaveU.Refresh
End Sub
