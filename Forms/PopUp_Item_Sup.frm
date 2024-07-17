VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PopUp_Item_Sup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item ID"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "PopUp_Item_Sup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin sknFD 
      Left            =   0
      OleObjectBlob   =   "PopUp_Item_Sup.frx":000C
      Top             =   0
   End
   Begin VB.PictureBox picFrame 
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5475
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   240
      Width           =   6615
      Begin VB.CommandButton cmdFind 
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   4335
      End
      Begin MSComctlLib.ListView lvItemId 
         Height          =   4575
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Id"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   3952
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unit"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Item ID "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PopUp_Item_Sup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim listItemID  As ListItem

Private Sub cmdfind_Click()
On Error GoTo errFind
    Dim qry As String
    If GetForm = "F_Mst_Product" Or GetForm = "F_Mst_Product_v2" Or GetForm = "Form_STE" Or GetForm = "Form_Forecast" Then
        qry = "SELECT item_id, item_name, um_name FROM mst_item a INNER JOIN r_unit_measure b ON a.um_id = b.um_id WHERE pfm_id = '10' and " _
    & "upper(item_id) LIKE '" & UCase(txtFilter) & "%' and stscode_id='01' ORDER BY item_id"
    Else
        qry = "SELECT item_id, item_name, um_name FROM mst_item a INNER JOIN r_unit_measure b ON a.um_id = b.um_id WHERE pfm_id = '06' AND " _
    & "upper(item_id) LIKE '" & UCase(txtFilter) & "%' ORDER BY item_id"
    End If
    Set RsGet = Con.Execute(qry)
    Call getList
    If lvItemId.ListItems.Count > 0 Then
        lvItemId.SetFocus
    End If
errFind:
    If Err.Number <> 0 Then
        MsgBox "Error... (" & Err.Description & ")", vbCritical, Err.Number
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call activeTheme(sknFD, Me)
    LV_FlatHeaders Me.hwnd, lvItemId.hwnd
    Call BukaKoneksi
Exit Sub
errLoad:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error Load [" & Err.Number & "]"
End Sub

Private Sub getList()
    lvItemId.ListItems.Clear
    Do Until RsGet.EOF
        Set listItemID = lvItemId.ListItems.Add(, , RTrim(RsGet!item_id))
            listItemID.SubItems(1) = RTrim(RsGet!item_name)
            listItemID.SubItems(2) = RTrim(RsGet!um_name)
        RsGet.MoveNext
    Loop
End Sub

Private Sub lvItemId_DblClick()
On Error Resume Next
    If GetForm = "F_Mst_Product_v2" Then
        If Not Me.lvItemId.SelectedItem Is Nothing Then
            Set listItemID = Me.lvItemId.SelectedItem
                F_Mst_Product_v2.txtItemid = RTrim(listItemID.Text)
                F_Mst_Product_v2.txtItemName = RTrim(listItemID.SubItems(1))
                GetForm = ""
                Unload Me
                F_Mst_Product_v2.txtTotalMold.SetFocus
        End If
    ElseIf GetForm = "Form_STE" Then
        If Not Me.lvItemId.SelectedItem Is Nothing Then
            Set listItemID = Me.lvItemId.SelectedItem
                Form_STE.txtPartNo = RTrim(listItemID.Text)
                GetForm = ""
                Unload Me
                Form_STE.txtMachine.SetFocus
        End If
    ElseIf GetForm = "Form_Forecast" Then
        If Not Me.lvItemId.SelectedItem Is Nothing Then
            Set listItemID = Me.lvItemId.SelectedItem
                Form_Forecast.txtItemid = RTrim(listItemID.Text)
                GetForm = ""
                Unload Me
                Form_Forecast.txtQty.SetFocus
        End If
    End If
End Sub

Private Sub lvItemId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lvItemId_DblClick
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdfind_Click
    End If
End Sub
