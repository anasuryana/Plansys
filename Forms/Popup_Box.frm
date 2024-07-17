VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Popup_Box 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box List"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3360
      OleObjectBlob   =   "Popup_Box.frx":0000
      Top             =   360
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Box ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Box Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Color"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "L"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "W"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "H"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Type ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Search"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtfind 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Popup_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qry As String

Private Sub cmdfind_Click()
    qry = "SELECT * FROM mst_box WHERE lower(boxid) like '%" & LCase(txtFind) & "%'"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount = 0 Then
        qry = "SELECT * FROM mst_box WHERE lower(boxname) like '%" & LCase(txtFind) & "%'"
        Set RsBantu = Con.Execute(qry)
        If RsBantu.RecordCount = 0 Then
            qry = "SELECT * FROM mst_box WHERE lower(color) like '%" & LCase(txtFind) & "%'"
            Set RsBantu = Con.Execute(qry)
        End If
    End If
    getList
End Sub

Private Sub getList()
    lv1.ListItems.Clear
    If RsBantu.RecordCount > 0 Then
        Dim li As ListItem
        While Not RsBantu.EOF
            Set li = lv1.ListItems.Add(, , RsBantu("boxid"))
            li.SubItems(1) = RsBantu("boxname")
            li.SubItems(2) = RsBantu("color")
            li.SubItems(3) = RsBantu("l")
            li.SubItems(4) = RsBantu("w")
            li.SubItems(5) = RsBantu("h")
            li.SubItems(6) = RsBantu("typeid")
            RsBantu.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
    activeTheme Skin1, Me
End Sub

Private Sub lv1_DblClick()
    F_Mst_Product_v2.typebox = lv1.SelectedItem.SubItems(6)
    F_Mst_Product_v2.txtBox = lv1.SelectedItem.SubItems(1)
    Unload Me
End Sub

Private Sub lv1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lv1_DblClick
    End If
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdfind_Click
    End If
End Sub
