VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PopUp_socMPP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SO Data List"
   ClientHeight    =   4440
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "PopUp_socMPP.frx":0000
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6588
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Find"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   1320
      OleObjectBlob   =   "PopUp_socMPP.frx":005E
      Top             =   240
   End
End
Attribute VB_Name = "PopUp_socMPP"
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
Dim qry As String
Dim listSOC As ListItem
Public lu_SOCID As String
Public lu_custName As String

Private Sub settingLV()
    With lv1
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
'        .ColumnHeaders.Add , , "Item ID", 2324.977
        .ColumnHeaders.Add , , "SOC ID", 3165.166
        .ColumnHeaders.Add , , "Customer", 2900
    End With
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

Private Sub Form_Load()
On Error GoTo errLoad
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    settingLV
    LoadDatanya
    LoadDatanya_V2
    Call getList
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub LoadDatanya_V2()
    If Len(Trim(txtFind)) > 0 Then
        RsGet.Filter = "soc_id LIKE '*" & txtFind & "*'"
    Else
        RsGet.Filter = adFilterNone
    End If
    If RsGet.RecordCount > 0 Then
        Call getList
    Else
        RsGet.Filter = adFilterNone
        RsGet.Filter = "cust_name LIKE '*" & txtFind & "*'"
        Call getList
    End If
    
End Sub

Private Sub LoadDatanya()
    qry = "select distinct on (aka.soc_id) soc_id, cust_name from " _
        & "(select a.soc_id,x4.cust_name " _
         & " from soc a " _
         & " inner join mst_item x1 on a.item_id = x1.item_id " _
         & " inner join r_prodfam x2 on x1.pfm_id = x2.pfm_id " _
         & " inner join r_unit_measure x3 on x1.um_id = x3.um_id " _
         & " inner join r_customer x4 on a.cust_id = x4.cust_id " _
         & " left join sod b on a.soc_id = b.sod_socid and a.item_id = b.item_id and inv_status = true " _
         & " where lower(a.soc_id) LIKE '%" & LCase(txtFind) & "%' " _
         & " group by a.soc_id, a.cust_id, a.soc_date, a.soc_cust_po, a.item_id, x1.item_name, a.soc_reqdate, a.soc_reqqty, " _
         & " x4.cust_name , X2.pfm_name, x3.um_name " _
         & " having sum(coalesce(b.sod_scanqty, 0))<a.soc_reqqty order by a.soc_id, a.item_id) aka"
    Set RsGet = Con.Execute(qry)
    Me.Caption = "SO Data List [" & RsGet.RecordCount & " row(s) found]"
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub lv1_DblClick()
    If Not lv1.SelectedItem Is Nothing Then
        Set listSOC = lv1.SelectedItem
            lu_SOCID = RTrim(listSOC.Text)
            lu_custName = listSOC.SubItems(1)
            Unload Me
    End If
End Sub

Private Sub lv1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lv1_DblClick
End Sub

Private Sub OKButton_Click()
    On Error GoTo errFind
    LoadDatanya
    LoadDatanya_V2
    If RsGet.RecordCount > 0 Then
        lv1.SetFocus
    End If
Exit Sub
errFind:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error: " & Err.Number
End Sub

Private Sub getList()
    lv1.ListItems.Clear
    If Not RsGet.EOF Then
        Do Until RsGet.EOF
            Set listSOC = lv1.ListItems.Add(, , RTrim(RsGet!soc_id))
            listSOC.SubItems(1) = RTrim(RsGet!cust_name)

            RsGet.MoveNext
        Loop
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then
        OKButton_Click
    End If
End Sub
