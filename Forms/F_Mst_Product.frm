VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Mst_Product 
   Caption         =   "Master Product"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "F_Mst_Product.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   13440
   Begin VB.Frame Frame2 
      Caption         =   "Find"
      Height          =   855
      Left            =   10080
      TabIndex        =   53
      Top             =   3480
      Width           =   3255
      Begin VB.TextBox txtFind 
         Height          =   360
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.TextBox txtUsedMold 
      Height          =   360
      Left            =   7440
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtFaktorProd 
      Height          =   360
      Left            =   7440
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtFaktorTotalMold 
      Height          =   360
      Left            =   7440
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   855
      Left            =   3480
      TabIndex        =   49
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   855
      Left            =   1800
      TabIndex        =   48
      Tag             =   "s"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add new"
      Height          =   855
      Left            =   120
      TabIndex        =   47
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtHourPshift 
      Height          =   360
      Left            =   7440
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtShift 
      Height          =   360
      Left            =   7440
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4335
      Left            =   120
      TabIndex        =   44
      Top             =   4440
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   7646
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
      NumItems        =   21
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Part No"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Part Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Man Power"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time Second Process"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Machine No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ALT MCH1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ALT MCH2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ALT MCH3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ALT MCH4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ALT MCH5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ALT MCH6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ALT MCH7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Cavity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Cycle Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Subcont"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Shift"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Hour per Shift"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Total Mold"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Used Mold"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Productivity Factor"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alternative Machines"
      Height          =   3375
      Left            =   8760
      TabIndex        =   29
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   17
         Top             =   3360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   42
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   16
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   40
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   15
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   38
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   14
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   36
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   34
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   32
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1680
         TabIndex        =   18
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   30
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_Mst_Product.frx":06EA
         TabIndex        =   31
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_Mst_Product.frx":075A
         TabIndex        =   33
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_Mst_Product.frx":07CA
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_Mst_Product.frx":083A
         TabIndex        =   37
         Top             =   1920
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_Mst_Product.frx":08AA
         TabIndex        =   39
         Top             =   2400
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_Mst_Product.frx":091A
         TabIndex        =   41
         Top             =   2880
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_Mst_Product.frx":098A
         TabIndex        =   43
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdFindMachine 
      Caption         =   "..."
      Height          =   375
      Left            =   4560
      TabIndex        =   28
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtMchineNo 
      Enabled         =   0   'False
      Height          =   360
      Left            =   2160
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "..."
      Height          =   375
      Left            =   4560
      TabIndex        =   26
      Top             =   120
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_Mst_Product.frx":09FA
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtSecondProcess 
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "SUBCONT"
      Height          =   255
      Left            =   5520
      TabIndex        =   19
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtCT 
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtCavity 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtManPower 
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtItemName 
      Enabled         =   0   'False
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtItemId 
      Enabled         =   0   'False
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   1680
      OleObjectBlob   =   "F_Mst_Product.frx":0A66
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_Mst_Product.frx":0C9A
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_Mst_Product.frx":0D02
      TabIndex        =   22
      Top             =   1080
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_Mst_Product.frx":0D6A
      TabIndex        =   23
      Top             =   1560
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_Mst_Product.frx":0DCC
      TabIndex        =   24
      Top             =   2040
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_Mst_Product.frx":0E40
      TabIndex        =   25
      Top             =   2520
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_Mst_Product.frx":0EBC
      TabIndex        =   27
      Top             =   3000
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "F_Mst_Product.frx":0F26
      TabIndex        =   45
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "F_Mst_Product.frx":0F86
      TabIndex        =   46
      Top             =   600
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "F_Mst_Product.frx":0FF4
      TabIndex        =   50
      Top             =   1560
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "F_Mst_Product.frx":105E
      TabIndex        =   51
      Top             =   1080
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "F_Mst_Product.frx":10DA
      TabIndex        =   52
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "F_Mst_Product"
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
Public c_machine_no As String
Private qry         As String
Private kebijakanSC As String
Private lisitm      As ListItem
Private id          As String

Private Sub LoadDatanya()
'On Error GoTo errFind
    Set RsGet = Con.Execute("select * from loadcap_mst_product")
    Call getList
'    If LV.ListItems.Count > 0 Then
'        LV.SetFocus
'    End If
'errFind:
'    If Err.Number <> 0 Then
'        MsgBox "Error... (" & Err.Description & ")", vbCritical, Err.Number
'    End If
End Sub
Private Sub LoadDatanya_v2()
    RsGet.MoveFirst
    If Len(Trim(txtFind)) > 0 Then
        RsGet.Filter = "partno LIKE '*" & txtFind & "*'"
    Else
        RsGet.Filter = adFilterNone
    End If
    Call getList
End Sub
Private Sub getList()
    LV.ListItems.Clear
    
    Do Until RsGet.EOF
        Set lisitm = LV.ListItems.Add(, , RTrim(RsGet!lc_idproduct))
            lisitm.SubItems(1) = RTrim(RsGet!partno)
            lisitm.SubItems(2) = RTrim(RsGet!partname)
            lisitm.SubItems(3) = RsGet!manpower
            lisitm.SubItems(4) = RsGet!time_sec_proc
            lisitm.SubItems(5) = RsGet!prod_nomach
            lisitm.SubItems(6) = RsGet!alt1_prod_nomach
            lisitm.SubItems(7) = RsGet!alt2_prod_nomach
            lisitm.SubItems(8) = RsGet!alt3_prod_nomach
            lisitm.SubItems(9) = RsGet!alt4_prod_nomach
            lisitm.SubItems(10) = RsGet!alt5_prod_nomach
            lisitm.SubItems(11) = RsGet!alt6_prod_nomach
            lisitm.SubItems(12) = RsGet!alt7_prod_nomach
            lisitm.SubItems(13) = RsGet!cavity
            lisitm.SubItems(14) = RsGet!cycletime
            lisitm.SubItems(15) = RsGet!kebijkan_subc
            lisitm.SubItems(16) = RsGet!shift_usg
            lisitm.SubItems(17) = RsGet!hour_p_shift
            lisitm.SubItems(18) = IIf(IsNull(RsGet!jml_mold), 0, RsGet!jml_mold)
            lisitm.SubItems(19) = IIf(IsNull(RsGet!jml_mold_digunakan), 0, RsGet!jml_mold_digunakan)
            lisitm.SubItems(20) = IIf(IsNull(RsGet!faktor_productivity), 0, RsGet!faktor_productivity)
        RsGet.MoveNext
    Loop
End Sub

Sub ResizeControls()
    On Error Resume Next
    Dim i As Integer
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

Private Sub Check1_Click()
    If Check1.value = vbChecked Then
        kebijakanSC = "yes"
    Else
        kebijakanSC = "no"
    End If
End Sub

Private Sub kosong()
    txtItemId = ""
    txtItemName = ""
    txtManPower = ""
    txtCavity = ""
    txtCT = ""
    txtSecondProcess = ""
    txtMchineNo = ""
    txtFaktorTotalMold = 0
    txtFaktorProd = 0
'    txtShift = ""
'    txtHourPshift = ""
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
End Sub

Private Sub cmdAdd_Click()
    kosong
    cmdFind.SetFocus
    cmdSave.Tag = "s"
End Sub

Private Sub cmdDelete_Click()
    If Len(txtMchineNo) > 1 Then
        qry = "delete from loadcap_mst_product where lc_idproduct=" & id
        Con.Execute qry
        MsgBox "Deleted", vbInformation, "Good"
        LoadDatanya
    End If
End Sub

Private Sub cmdFind_Click()
    GetForm = Me.Name
    PopUp_Item_Sup.Show 1
End Sub

Private Sub cmdFindMachine_Click()
    GetForm = Me.Name
    PopUp_machine.Show 1
    txtMchineNo = c_machine_no
End Sub

Private Sub cmdSave_Click()
On Error GoTo ER_exc
    If cmdSave.Tag = "s" Then
        If IsNumeric(txtCavity) = False Or IsNumeric(txtManPower) = False _
            Or IsNumeric(txtCT) = False Or IsNumeric(txtSecondProcess) = False _
            Or Len(txtItemName) < 2 Or IsNumeric(txtFaktorProd) = False Or IsNumeric(txtFaktorTotalMold) = False _
            Or IsNumeric(txtUsedMold) = False _
        Then
            Exit Sub
        End If
        If MsgBox("Save  ?", vbQuestion + vbYesNo) = vbYes Then
            BukaKoneksi
            qry = "insert into loadcap_mst_product values('" & txtItemId & "' " _
                & ",'" & txtItemName & "'," & txtManPower & "," & txtSecondProcess & "" _
                & ",'" & txtMchineNo & "','" & Text1 & "','" & Text2 & "','" & Text3 & "'" _
                & ",'" & Text4 & "','" & Text5 & "','" & Text6 & "','" & Text7 & "'" _
                & "," & txtCavity & "," & txtCT & ",'" & kebijakanSC & "'," & txtShift & "," _
                & txtHourPshift & ",DEFAULT," & txtFaktorTotalMold & "," & txtFaktorProd & "," & txtUsedMold & ")"
                Con.Execute qry
                MsgBox "Saved "
        End If
    Else
        If MsgBox("Update  ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        BukaKoneksi
        qry = "update loadcap_mst_product set manpower=" & txtManPower & "" _
            & ",time_sec_proc=" & txtSecondProcess & ",prod_nomach='" & txtMchineNo & "'" _
            & ",alt1_prod_nomach='" & Text1 & "',alt2_prod_nomach='" & Text2 & "'" _
            & ",alt3_prod_nomach='" & Text3 & "',alt4_prod_nomach='" & Text4 & "'" _
            & ",alt5_prod_nomach='" & Text5 & "',alt6_prod_nomach='" & Text6 & "'" _
            & ",alt7_prod_nomach='" & Text7 & "',cavity=" & txtCavity & ",cycletime=" & txtCT _
            & ",kebijkan_subc='" & kebijakanSC & "',jml_mold=" & txtFaktorTotalMold & "" _
            & ",jml_mold_digunakan=" & txtUsedMold & "" _
            & " where lc_idproduct=" & id
        Con.Execute qry
        MsgBox "Updated ", vbInformation, "Good !"
    End If
    updatesHIFT
    LoadDatanya
    Exit Sub
ER_exc:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub updatesHIFT()
    Dim aqry As String
    aqry = "update loadcap_mst_product set shift_usg=" & txtShift & ",hour_p_shift=" & txtHourPshift & ",faktor_productivity=" & txtFaktorProd
    Con.Execute aqry
End Sub

Private Sub Command1_Click()
    GetForm = Me.Name
    PopUp_machine.Show 1
    Text1 = c_machine_no
End Sub

Private Sub Command2_Click()
     GetForm = Me.Name
    PopUp_machine.Show 1
    Text2 = c_machine_no
End Sub

Private Sub Command3_Click()
 GetForm = Me.Name
    PopUp_machine.Show 1
    Text3 = c_machine_no
End Sub

Private Sub Command4_Click()
     GetForm = Me.Name
    PopUp_machine.Show 1
    Text4 = c_machine_no
End Sub

Private Sub Command5_Click()
     GetForm = Me.Name
    PopUp_machine.Show 1
    Text5 = c_machine_no
End Sub

Private Sub Command6_Click()
     GetForm = Me.Name
    PopUp_machine.Show 1
    Text6 = c_machine_no
End Sub

Private Sub Command7_Click()
     GetForm = Me.Name
    PopUp_machine.Show 1
    Text7 = c_machine_no
End Sub

Private Sub Form_Activate()
    FocusTab Me
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

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        DelTab Me
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Me.Height = 9405
    Me.Width = 13680
    kebijakanSC = "no"
    LoadDatanya
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub lvtOform()
    If LV.ListItems.Count > 0 Then
        id = LV.SelectedItem.Text
        txtItemId = LV.SelectedItem.SubItems(1)
        txtItemName = LV.SelectedItem.SubItems(2)
        txtManPower = LV.SelectedItem.SubItems(3)
        txtSecondProcess = LV.SelectedItem.SubItems(4)
        txtMchineNo = LV.SelectedItem.SubItems(5)
        Text1 = LV.SelectedItem.SubItems(6)
        Text2 = LV.SelectedItem.SubItems(7)
        Text3 = LV.SelectedItem.SubItems(8)
        Text4 = LV.SelectedItem.SubItems(9)
        Text5 = LV.SelectedItem.SubItems(10)
        Text6 = LV.SelectedItem.SubItems(11)
        Text7 = LV.SelectedItem.SubItems(12)
        txtCavity = LV.SelectedItem.SubItems(13)
        txtCT = LV.SelectedItem.SubItems(14)
        If LV.SelectedItem.SubItems(15) = "yes" Then
            Check1.value = 1
        Else
            Check1.value = 0
        End If
        txtShift = LV.SelectedItem.SubItems(16)
        txtHourPshift = LV.SelectedItem.SubItems(17)
        txtFaktorTotalMold = LV.SelectedItem.SubItems(18)
        txtUsedMold = LV.SelectedItem.SubItems(19)
        txtFaktorProd = LV.SelectedItem.SubItems(20)
        Check1.Refresh
    End If
End Sub

Private Sub LV_Click()
    cmdSave.Tag = "s"
    lvtOform
End Sub

Private Sub LV_DblClick()
    cmdSave.Tag = "u"
    txtManPower.SetFocus
End Sub

Private Sub LV_KeyUp(KeyCode As Integer, Shift As Integer)
    cmdSave.Tag = "s"
    lvtOform
End Sub

Private Sub txtFaktorTotalMold_Change()
    If IsNumeric(txtFaktorTotalMold) Then
        txtUsedMold = txtFaktorTotalMold
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFind = FilterIn(txtFind)
        LoadDatanya_v2
    End If
End Sub
