VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form popup_wono 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Order List"
   ClientHeight    =   5925
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "popup_wono.frx":0000
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtFind 
      Height          =   390
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9340
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "WO Id"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lot No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Id"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Item Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Machine"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Mould No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Plan Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Plan Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "MPP Rev"
         Object.Width           =   2540
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "popup_wono.frx":0064
      Top             =   360
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "popup_wono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim qry As String

Private Sub Form_Load()
    activeTheme Skin1, Me
End Sub

Private Sub loadWO()
    Dim li As ListItem
    Dim ke As Integer
    Const kols As String = "wo_no,partno,item_name,coalesce(lotno,'') lotno,mesinno,moldno,issudate,coalesce(qty,0) qty,mpprev"
    qry = "SELECT " & kols & " FROM worko a inner join mst_item b on a.partno=b.item_id" _
    & " where lower(wo_no) like  '%" & LCase(txtFind) & "%' order by partno asc"
    Set RsBantu = Con.Execute(qry)
    lv1.ListItems.Clear
    If RsBantu.RecordCount > 0 Then
        While Not RsBantu.EOF
            ke = ke + 1
            Set li = lv1.ListItems.Add(, , ke)
            li.SubItems(1) = RsBantu("wo_no")
            li.SubItems(2) = RsBantu("lotno")
            li.SubItems(3) = RsBantu("partno")
            li.SubItems(4) = RsBantu("item_name")
            li.SubItems(5) = RsBantu("mesinno")
            li.SubItems(6) = RsBantu("moldno")
            li.SubItems(7) = RsBantu("qty")
            li.SubItems(8) = Format(RsBantu("issudate"), "yyyy-MM-dd")
            li.SubItems(9) = RsBantu("mpprev")
            RsBantu.MoveNext
        Wend
    End If
    Set RsBantu = Nothing
End Sub

Private Sub lv1_DblClick()
    OKButton_Click
End Sub

Private Sub lv1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OKButton_Click
    End If
End Sub

Private Sub OKButton_Click()
    If GetForm = "Form_CancelWO" Then
        If lv1.ListItems.Count = 0 Then Exit Sub
        If lv1.SelectedItem.SubItems(2) = "" Then
            MsgBox "The data is already canceled"
            Exit Sub
        End If
        Form_CancelWO.txtWoid = lv1.SelectedItem.SubItems(1)
        Form_CancelWO.lbllotno = lv1.SelectedItem.SubItems(2)
        Form_CancelWO.lblitemid = lv1.SelectedItem.SubItems(3)
        Form_CancelWO.lblitemname = lv1.SelectedItem.SubItems(4)
        Form_CancelWO.lblmachine = lv1.SelectedItem.SubItems(5)
        Form_CancelWO.lblmoldno = lv1.SelectedItem.SubItems(6)
        Form_CancelWO.lblplanqty = lv1.SelectedItem.SubItems(7)
        Form_CancelWO.lblplandate = "Plan date : " & lv1.SelectedItem.SubItems(8)
        Form_CancelWO.lblrevisiMPP = "MPP Revision : " & lv1.SelectedItem.SubItems(9)
        GetForm = ""
        Unload Me
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        loadWO
    End If
End Sub
