VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_SetSparepart 
   Caption         =   "Setting Sparepart"
   ClientHeight    =   10320
   ClientLeft      =   3120
   ClientTop       =   645
   ClientWidth     =   14025
   Icon            =   "Form_SetSparepart.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10320
   ScaleWidth      =   14025
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin skn 
      Left            =   9960
      OleObjectBlob   =   "Form_SetSparepart.frx":000C
      Top             =   240
   End
   Begin VB.PictureBox FrameFind 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   10395
      TabIndex        =   1
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton cmdFind 
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtAssy 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Assy No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lvSparepart 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   15266
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Assy No"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Assy Name"
         Object.Width           =   7832
      EndProperty
   End
End
Attribute VB_Name = "Form_SetSparepart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim listSparepart As ListItem

Private Sub cmdFind_Click()
On Error GoTo errFind
    getDataSparepart txtAssy
Exit Sub
errFind:
    If Err.Number <> "" Then
        MsgBox Err.Description, vbCritical, "Error Find: " & Err.Number
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call activeTheme(skn, Me)
    'Call BukaKoneksi
    LV_FlatHeaders Me.hWnd, lvSparepart.hWnd
    getDataSparepart
Exit Sub
errLoad:
    If Err.Number <> "" Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub getDataSparepart(Optional ItemID As String)
    Set RsGet = Con.Execute("select item_id, item_name, st_sparepart from mst_item where pfm_id = '06' and item_id like '" & ItemID & "%' order by item_id")
    lvSparepart.ListItems.Clear
    If Not RsGet.EOF Then
        Do Until RsGet.EOF
                Set listSparepart = lvSparepart.ListItems.Add(, , lvSparepart.ListItems.Count + 1)
                    listSparepart.SubItems(1) = RTrim(RsGet!item_id)
                    listSparepart.SubItems(2) = RTrim(RsGet!item_name)
                If RsGet!st_sparepart = 1 Then
                    listSparepart.Checked = True
                End If
                RsGet.MoveNext
        Loop
    End If
    RsGet.Close
End Sub

Private Sub lvSparepart_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo errCheck
    If Item.Checked Then
        Con.Execute "update mst_item set st_sparepart = true where item_id = '" & Item.ListSubItems(1) & "'"
    Else
        Con.Execute "update mst_item set st_sparepart = false where item_id = '" & Item.ListSubItems(1) & "'"
    End If
Exit Sub
errCheck:
    If Err.Number <> "" Then
        MsgBox Err.Description, vbCritical, "Error Change: " & Err.Number
    End If
End Sub

Private Sub txtAssy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdFind_Click
    End If
End Sub
