VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Mst_Mesin 
   Caption         =   "Master Machine"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "F_Mst_Mesin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   13050
   Begin VB.PictureBox picPop1 
      BackColor       =   &H00C0FFC0&
      Height          =   2775
      Left            =   2880
      ScaleHeight     =   2715
      ScaleWidth      =   6915
      TabIndex        =   27
      Top             =   2520
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtFindPart 
         Height          =   345
         Left            =   600
         TabIndex        =   29
         Top             =   360
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvPopPart 
         Height          =   1815
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   6615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   32
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lbltotalfind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Find"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Part"
      Height          =   1335
      Left            =   7800
      TabIndex        =   24
      Top             =   1080
      Width           =   2535
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdDeleteItem 
         Caption         =   "Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Material"
      Height          =   1335
      Left            =   10440
      TabIndex        =   19
      Top             =   1080
      Width           =   2535
      Begin VB.CommandButton CmdDeleteMat 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddmat 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbMat 
         Height          =   360
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView lv2 
      Height          =   3735
      Left            =   10440
      TabIndex        =   18
      Top             =   2520
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6588
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Active"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtSpec 
      Height          =   360
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "F_Mst_Mesin.frx":06EA
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   5040
      OleObjectBlob   =   "F_Mst_Mesin.frx":0748
      Top             =   120
   End
   Begin VB.TextBox txtMerk 
      Height          =   360
      Left            =   6240
      MaxLength       =   50
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3735
      Left            =   45
      TabIndex        =   10
      Top             =   2520
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      ToolTipText     =   "Delete user data"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Save changes"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   50
      TabIndex        =   7
      ToolTipText     =   "Add new user"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox CheckSMED 
      Caption         =   "SMED"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtmachine_tonage 
      Height          =   360
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtmachine_name 
      Height          =   360
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtmachine_no 
      Height          =   360
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtLine 
      Height          =   360
      Left            =   4920
      MaxLength       =   40
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_Mst_Mesin.frx":097C
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_Mst_Mesin.frx":09E6
      TabIndex        =   13
      Top             =   600
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5400
      OleObjectBlob   =   "F_Mst_Mesin.frx":0A54
      TabIndex        =   14
      Top             =   600
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   5400
      OleObjectBlob   =   "F_Mst_Mesin.frx":0AB4
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_Mst_Mesin.frx":0B16
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
   End
   Begin MSComctlLib.ListView lv3 
      Height          =   3735
      Left            =   7800
      TabIndex        =   23
      Top             =   2520
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6588
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_Mst_Mesin"
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
Private qry As String
Private idS As String
Private idS_part As String
Private rsLoc As ADODB.Recordset
Private rsMat As ADODB.Recordset
Private rsMat2 As ADODB.Recordset
Private li As ListItem

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

Private Sub settingLV()
    With lv1
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "mcid", 0
        .ColumnHeaders.Add , , "Line", 0
        .ColumnHeaders.Add , , "Machine No"
        .ColumnHeaders.Add , , "Machine Name", 2500
        .ColumnHeaders.Add , , "Tonage"
        .ColumnHeaders.Add , , "Brand"
        .ColumnHeaders.Add , , "Material Used", 0
        .ColumnHeaders.Add , , "SMED"
        .ColumnHeaders.Add , , "Specification", 2500
        .ColumnHeaders.Add , , "Active"
    End With
    With LV2
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "idmatused", 0
        .ColumnHeaders.Add , , "idmch", 0
        .ColumnHeaders.Add , , "Material"
    End With
    With lv3
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "idpartused", 0
        .ColumnHeaders.Add , , "idmat", 0
        .ColumnHeaders.Add , , "Part"
        .ColumnHeaders.Add , , "Customer"
    End With
    With lvPopPart
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxes = True
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Item Id", 2600
        .ColumnHeaders.Add , , "Item Name", 3200
        .ColumnHeaders.Add , , "Customer", 3000
        .ColumnHeaders.Add , , "Machine", 0
        .ColumnHeaders.Add , , "Machine ALT", 2000
    End With
End Sub

Private Sub gridtoform()
On Error GoTo Excp
    If lv1.ListItems.Count > 0 Then
        idS = lv1.SelectedItem.Text
        txtLine = lv1.SelectedItem.ListSubItems(1).Text
        txtmachine_no = lv1.SelectedItem.ListSubItems(2).Text
        txtmachine_name = lv1.SelectedItem.ListSubItems(3).Text
        txtmachine_tonage = lv1.SelectedItem.ListSubItems(4).Text
        txtMerk = lv1.SelectedItem.ListSubItems(5).Text
        txtSpec = lv1.SelectedItem.ListSubItems(8).Text
        If lv1.SelectedItem.ListSubItems(7).Text = "SMED" Then
            CheckSMED.Value = vbChecked
        Else
            CheckSMED.Value = vbUnchecked
        End If
        CheckSMED.Refresh
        If lv1.SelectedItem.ListSubItems(9).Text = 1 Then
            Check1.Value = vbChecked
        Else
            Check1.Value = vbUnchecked
        End If
        Check1.Refresh
        qry = "select * from loadcap_matused where material_mch='" & txtmachine_no & "' order by 2 asc"
        Set rsMat2 = Con.Execute(qry)
        LV2.ListItems.Clear
        If rsMat2.RecordCount > 0 Then
            While Not rsMat2.EOF
                Set li = LV2.ListItems.Add(, , rsMat2(0))
                li.SubItems(1) = rsMat2(2).Value
                li.SubItems(2) = rsMat2(1).Value
                rsMat2.MoveNext
            Wend
        End If
        qry = "select idpartrun,part_used,part_mch,cust_name from loadcap_partused a inner join mst_item b on a.part_used=b.item_id inner join r_customer c on b.cust_id=c.cust_id where part_mch='" & txtmachine_no & "' order by 2 asc"
        Set rsMat2 = Con.Execute(qry)
        lv3.ListItems.Clear
        If rsMat2.RecordCount > 0 Then
            While Not rsMat2.EOF
                Set li = lv3.ListItems.Add(, , rsMat2(0))
                li.SubItems(1) = rsMat2(2).Value
                li.SubItems(2) = rsMat2(1).Value
                li.SubItems(3) = rsMat2(3).Value
                rsMat2.MoveNext
            Wend
        End If
    End If
    Exit Sub
Excp:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub loadData()
    qry = "select * from loadcap_mst_mach order by no_mach asc"
    Set rsLoc = New ADODB.Recordset
    rsLoc.Open qry, Con, adOpenStatic, adLockOptimistic
    If rsLoc.RecordCount > 0 Then
        lv1.ListItems.Clear
        Dim xx As ListItem
        While Not rsLoc.EOF
            Set xx = lv1.ListItems.Add(, , rsLoc("idmst_mach"))
            xx.SubItems(1) = IIf(IsNull(rsLoc("line_mach")), "", rsLoc("line_mach"))
            xx.SubItems(2) = IIf(IsNull(rsLoc("no_mach")), "", rsLoc("no_mach"))
            xx.SubItems(3) = IIf(IsNull(rsLoc("name_mach")), "", rsLoc("name_mach"))
            xx.SubItems(4) = IIf(IsNull(rsLoc("tonage_mach")), "", rsLoc("tonage_mach"))
            xx.SubItems(5) = IIf(IsNull(rsLoc("brand_mach")), "", rsLoc("brand_mach"))
            xx.SubItems(6) = IIf(IsNull(rsLoc("material_used")), "", rsLoc("material_used"))
            xx.SubItems(7) = IIf(IsNull(rsLoc("smed_mach")), "", rsLoc("smed_mach"))
            xx.SubItems(8) = IIf(IsNull(rsLoc("specification")), "", rsLoc("specification"))
            xx.SubItems(9) = IIf(IsNull(rsLoc("state_mach")), 0, rsLoc("state_mach"))
            rsLoc.MoveNext
        Wend
    Else
        lv1.ListItems.Clear
    End If
End Sub

Private Sub KosongkanForm()
    txtLine = ""
    txtmachine_no = ""
    txtmachine_name = ""
    txtMerk = ""
    txtmachine_tonage = ""
'    txtmaterial_used = ""
    txtSpec = ""
End Sub

'Private Sub Check2_Click()
'    Dim a As Long
'    If lvPopPart.ListItems.Count < 1 Then Exit Sub
'    If Check2.Value Then
'        For a = 1 To lvPopPart.ListItems.Count
'            lvPopPart.ListItems(a).Checked = True
'        Next
'    Else
'        For a = 1 To lvPopPart.ListItems.Count
'            lvPopPart.ListItems(a).Checked = False
'        Next
'    End If
'End Sub

Private Sub cmdAdd_Click()
    txtmachine_no.SetFocus
    KosongkanForm
End Sub

Private Sub generateMachineNo()
    qry = "select idmst_mach from loadcap_mst_mach order by 1 desc limit 1"
    Set rsLoc = New Recordset
    rsLoc.Open qry, Con, adOpenStatic, adLockOptimistic
    If rsLoc.RecordCount > 0 Then
        Dim temP As String
        temP = Val(rsLoc(0).Value) + 1
        idS = Right("0000" & temP, 4)
    Else
        idS = "0001"
    End If
End Sub

Private Sub setUncheck()
    Dim e As Integer
    For e = 1 To lvPopPart.ListItems.Count
        If lvPopPart.ListItems(e).Checked Then
            lvPopPart.ListItems(e).Checked = False
        End If
    Next
End Sub

Private Sub cmdAddItem_Click()
    If txtmachine_no <> "" Then
        picPop1.Visible = True
        txtFindPart_KeyPress 13
    Else
        MsgBox "pilih mesin terlebih dahulu ", vbInformation, "Informasi"
        lv1.SetFocus
    End If
End Sub

Private Sub cmdAddmat_Click()
On Error GoTo eXb
    If txtmachine_no <> "" Then
        qry = "insert into loadcap_matused values(default,'" & cmbMat & "','" & txtmachine_no & "')"
        Con.Execute qry
        MsgBox "Saved successfully", vbInformation
    End If
    Exit Sub
eXb:
    MsgBox Err.Description, vbInformation, Err.Number
End Sub

Private Sub cmdDelete_Click()
On Error GoTo cEror
    If MsgBox("Delete ?", vbQuestion + vbYesNo, "WARNING") = vbYes Then
        qry = "delete from loadcap_mst_mach where idmst_mach='" & idS & "'"
        Con.Execute qry
        MsgBox "deleted...", vbInformation
        loadData
    End If
    Exit Sub
cEror:
    MsgBox Err.Description
End Sub

Private Sub cmdDeleteItem_Click()
    If Len(idS_part) > 0 Then
        Dim a As Integer
        Dim ttl As Integer
        ttl = 0
        For a = 1 To lv3.ListItems.Count
            If lv3.ListItems(a).Selected Then
                ttl = ttl + 1
            End If
        Next
        If ttl > 1 Then
            If MsgBox("Are you sure ?", vbQuestion + vbYesNo, "Multiple") = vbYes Then
                For a = 1 To lv3.ListItems.Count
                    If lv3.ListItems(a).Selected Then
                        qry = "DELETE FROM loadcap_partused where idpartrun = " & lv3.ListItems(a).Text
                        Con.Execute qry
                    End If
                Next
                MsgBox "Deleted successfully", vbInformation
            End If
        Else
            If MsgBox("Are you sure ?", vbQuestion + vbYesNo) = vbYes Then
                qry = "DELETE FROM loadcap_partused where idpartrun = " & idS_part
                Con.Execute qry
                MsgBox "Deleted successfully", vbInformation
            End If
        End If
    End If
    gridtoform
End Sub

Private Sub CmdDeleteMat_Click()
    If Len(idS) > 0 Then
        If MsgBox("Are you sure ?", vbQuestion + vbYesNo) = vbYes Then
            qry = "DELETE from loadcap_matused where idmater = " & idS
            Con.Execute qry
            MsgBox "Deleted successfully", vbInformation
        End If
    End If
End Sub

'Private Sub cmdOK_Click()
'    Dim rowKe As Double
'    With lvPopPart
'        For rowKe = 1 To .ListItems.Count
'            If .ListItems(rowKe).Checked Then
'                qry = "select count(*) from loadcap_partused where part_used='" & .ListItems(rowKe) & "' and part_mch='" & txtmachine_no & "'"
'                Set RsGet = Con.Execute(qry)
'                If RsGet(0) = 0 Then
'                    qry = "insert into loadcap_partused values(default,'" & .ListItems(rowKe) & "','" & txtmachine_no & "')"
'                    Con.Execute qry
'                End If
'            End If
'        Next
'    End With
'    picPop1.Visible = False
'End Sub

Private Sub cmdSave_Click()
    On Error GoTo Exct
    If Len(txtmachine_name) > 0 And Len(txtmachine_no) > 0 Then
        Dim statusSMED As String, statusMCH As Boolean
        statusSMED = IIf(CheckSMED.Value, "SMED", "NOSMED")
        statusMCH = IIf(Check1.Value, True, False)
        BukaKoneksi
        If cmdSave.Caption = "Save" Then
            generateMachineNo
            Set rsLoc = New ADODB.Recordset
            rsLoc.Open "loadcap_mst_mach", Con, adOpenKeyset, adLockOptimistic, adCmdTable
            rsLoc.AddNew
            rsLoc!idmst_mach = idS
            rsLoc!line_mach = txtLine
            rsLoc!no_mach = txtmachine_no
            rsLoc!name_mach = txtmachine_name
            rsLoc!tonage_mach = txtmachine_tonage
            rsLoc!smed_mach = statusSMED
            rsLoc!brand_mach = txtMerk
            rsLoc!specification = txtSpec
            rsLoc!state_mach = statusMCH
            rsLoc.Update
            MsgBox "Saved successfully"
            rsLoc.Close
            Set rsLoc = Nothing
        Else
            qry = "update loadcap_mst_mach set line_mach='" & txtLine & "', no_mach='" & txtmachine_no & "', name_mach='" & txtmachine_name & "'" _
                & " ,tonage_mach=" & txtmachine_tonage & ", smed_mach='" & statusSMED & "'" _
                & " ,brand_mach='" & txtMerk & "',specification='" & txtSpec & "',state_mach=" & statusMCH _
               & " where idmst_mach='" & idS & "' "
            Con.Execute qry
            MsgBox "Update successfully"
            cmdSave.Caption = "Save"
        End If
    End If
    loadData
    Exit Sub
Exct:
    MsgBox Err.Description, vbCritical, "Maaf " & Err.Number
End Sub

Private Sub LoadDatanya()
'    qry = "select item_id,item_name,cust_name,prod_nomach,alternatif from mst_item a inner join r_customer b on a.cust_id=b.cust_id left join loadcap_proc c " _
        & " on a.item_id=c.partno left join (select partno,string_agg(prod_nomach,',') alternatif from loadcap_proc group by partno) ad on c.partno=ad.partno " _
        & " where pfm_id = '10' " _
        & " group by item_id,prod_nomach,cust_name,alternatif " _
        & " order by 1 asc"
    qry = "select item_id,item_name,cust_name,alternatif from mst_item a inner join r_customer b on a.cust_id=b.cust_id left join loadcap_proc c " _
        & " on a.item_id=c.partno left join (select partno,string_agg(prod_nomach,',') alternatif from loadcap_proc group by partno) ad on c.partno=ad.partno " _
        & " where pfm_id = '10' and alternatif LIKE '%" & txtmachine_no & "%'" _
        & " group by item_id,cust_name,alternatif " _
        & " order by 1 asc"
    Set RsGet = Con.Execute(qry)
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

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    settingLV
    loadData
    Me.Height = 6840
    Me.Width = 13275
    qry = "select distinct on (item_name) item_name from mst_item where type_id ='02'"
    Set rsMat = Con.Execute(qry)
    While Not rsMat.EOF
        cmbMat.AddItem rsMat(0)
        rsMat.MoveNext
    Wend
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    LV2.ColumnHeaders(3).Width = LV2.Width
'    lv3.ColumnHeaders(4).Width = lv3.Width - lv3.ColumnHeaders(3).Width
    cmbMat.Width = cmdAddmat.Width + CmdDeleteMat.Width + 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Call WheelUnHook(Me.hwnd)
        DelTab Me
    End If
End Sub

Private Sub Label2_Click()
    picPop1.Visible = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Static lX As Integer, lY As Single
    If Button = vbLeftButton Then
        picPop1.Left = picPop1.Left + (x - lX)
        picPop1.Top = picPop1.Top + (Y - lY)
    Else
        lX = x: lY = Y
    End If
End Sub

Private Sub lv1_Click()
    gridtoform
    picPop1.Visible = False
End Sub

Private Sub lv1_DblClick()
    cmdSave.Caption = "Update"
    cmdSave.Refresh
    gridtoform
    txtmachine_no.SetFocus
End Sub

Private Sub lv1_KeyUp(KeyCode As Integer, Shift As Integer)
    gridtoform
End Sub

Private Sub lv2toForm()
    If LV2.ListItems.Count > 0 Then
        idS = LV2.SelectedItem.Text
    Else
        idS = ""
    End If
End Sub

Private Sub lv3toForm()
    If lv3.ListItems.Count > 0 Then
        idS_part = lv3.SelectedItem.Text
    Else
        idS_part = ""
    End If
End Sub

Private Sub LoadDatanya_V2()
    If Len(Trim(txtFindPart)) > 0 Then
        RsGet.Filter = "item_id LIKE '*" & txtFindPart & "*' and alternatif LIKE '*" & txtmachine_no & "*' "
        If RsGet.RecordCount = 0 Then
            RsGet.Filter = adFilterNone
            RsGet.Filter = "item_name LIKE '*" & txtFindPart & "*' and alternatif LIKE '*" & txtmachine_no & "*'"
            If RsGet.RecordCount = 0 Then
                RsGet.Filter = adFilterNone
                RsGet.Filter = "cust_name LIKE '*" & txtFindPart & "*' and alternatif LIKE '*" & txtmachine_no & "*'"
                If RsGet.RecordCount = 0 Then
                    RsGet.Filter = adFilterNone
                    RsGet.Filter = "alternatif LIKE '*" & txtFindPart & "*' and alternatif LIKE '*" & txtmachine_no & "*'"
                End If
            End If
        End If
    Else
        RsGet.Filter = adFilterNone
    End If
    Call getList

End Sub

Private Sub getList()
    lvPopPart.ListItems.Clear
    lbltotalfind.Caption = RsGet.RecordCount & " row(s) found"
    Do Until RsGet.EOF
        Set li = lvPopPart.ListItems.Add(, , RTrim(RsGet!item_id))
            li.SubItems(1) = RTrim(RsGet!item_name)
            li.SubItems(2) = RTrim(RsGet!cust_name)
            'li.SubItems(3) = RTrim(IIf(IsNull(RsGet!prod_nomach), "-", RsGet!prod_nomach))
            li.SubItems(4) = RTrim(IIf(IsNull(RsGet!alternatif), "-", RsGet!alternatif))
        RsGet.MoveNext
    Loop
End Sub

Private Sub LV2_Click()
    lv2toForm
End Sub

Private Sub lv3_Click()
    lv3toForm
End Sub

Private Sub lv3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 1 Then
        Dim a As Integer
        For a = 1 To lv3.ListItems.Count
            lv3.ListItems(a).Selected = True
        Next
    End If
End Sub


Private Sub lvPopPart_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim mesin As String
    If Item.Checked = False Then
        qry = "DELETE FROM loadcap_partused WHERE part_used='" & Item.Text & "' AND part_mch='" & txtmachine_no & "'"
        Con.Execute qry
    Else
        qry = "SELECT part_mch FROM loadcap_partused WHERE part_used='" & Item.Text & "' "
        Set RsGet = Con.Execute(qry)
        If RsGet.RecordCount > 0 Then
            While Not RsGet.EOF
                mesin = mesin & RsGet(0) & ", "
                RsGet.MoveNext
            Wend
        End If
        Set RsGet = Nothing
        If Len(mesin) > 0 Then
            If MsgBox("The item has been planned on the following machine " & mesin & vbNewLine & " commit to check ?", vbQuestion + vbYesNo) = vbYes Then
                qry = "select count(*) from loadcap_partused where part_used='" & Item.Text & "' and part_mch='" & txtmachine_no & "'"
                Set RsGet = Con.Execute(qry)
                If RsGet(0) = 0 Then
                    qry = "insert into loadcap_partused values(default,'" & Item.Text & "','" & txtmachine_no & "')"
                    Con.Execute qry
                End If
            Else
                qry = "DELETE FROM loadcap_partused WHERE part_used='" & Item.Text & "' AND part_mch='" & txtmachine_no & "'"
                Con.Execute qry
                lvPopPart.ListItems(Item.Index).Checked = False
            End If
        Else
            qry = "insert into loadcap_partused values(default,'" & Item.Text & "','" & txtmachine_no & "')"
            Con.Execute qry
        End If
    End If
End Sub

Private Sub txtFindPart_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LoadDatanya
        LoadDatanya_V2
        gridtoform
        setCheckedAdded
    End If
    
End Sub

Private Sub setCheckedAdded()
    Dim u As Integer, e As Integer
    For u = 1 To lv3.ListItems.Count
        For e = 1 To lvPopPart.ListItems.Count
            If lv3.ListItems(u).ListSubItems(2).Text = lvPopPart.ListItems(e).Text Then
                lvPopPart.ListItems(e).Checked = True
            End If
        Next
    Next
End Sub


Private Sub txtLine_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtmachine_name_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtmachine_no_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtmachine_tonage_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA & HURUFCEGAH, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtmaterial_used_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtMerk_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtSpec_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
