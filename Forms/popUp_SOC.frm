VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form popUp_SOC 
   Caption         =   "SOC"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8895
   Icon            =   "popUp_SOC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "popUp_SOC.frx":000C
      Top             =   0
   End
   Begin VB.PictureBox picFrame 
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5475
      ScaleWidth      =   8355
      TabIndex        =   3
      Top             =   240
      Width           =   8415
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   5655
      End
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
         Left            =   7320
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ListView lvSOC 
         Height          =   4575
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SOC ID"
            Object.Width           =   5362
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SOC Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SOC PO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cust ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cust Name"
            Object.Width           =   5010
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cust Address"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SOC ID"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "popUp_SOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim listSOC As ListItem

Private Sub cmdFind_Click()
On Error GoTo errFind
    Set RsGet = Con.Execute("select distinct soc_id, soc_date, soc_cust_po, a.cust_id, b.cust_name, b.cust_address from soc a inner join r_customer b on a.cust_id = b.cust_id where upper(soc_id) like '" & UCase(txtFilter) & "%'")
    Call getList
    If RsGet.RecordCount > 0 Then
        lvSOC.SetFocus
    End If
    RsGet.Close
Exit Sub
errFind:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error: " & Err.Number
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call activeTheme(skinFD, Me)
Exit Sub
errLoad:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error: " & Err.Number
End Sub

Private Sub getList()
    lvSOC.ListItems.Clear
    If Not RsGet.EOF Then
        Do Until RsGet.EOF
            Set listSOC = lvSOC.ListItems.Add(, , RsGet!soc_id)
                listSOC.SubItems(1) = RsGet!soc_date
                listSOC.SubItems(2) = RsGet!soc_cust_po
                listSOC.SubItems(3) = RsGet!cust_ID
                listSOC.SubItems(4) = RsGet!cust_name
                listSOC.SubItems(5) = RsGet!cust_address
            RsGet.MoveNext
        Loop
    End If
End Sub
    
Private Sub lvSOC_DblClick()
On Error Resume Next
    
    If GetForm = "Form_KanbanFGCreate" Then
        If Not Me.lvSOC.SelectedItem Is Nothing Then
            Set listSOC = Me.lvSOC.SelectedItem
                Form_KanbanFGCreate.txtSoId = RTrim(listSOC.Text)
                Form_KanbanFGCreate.txtCustID = RTrim(listSOC.SubItems(3))
                Form_KanbanFGCreate.txtCustName = RTrim(listSOC.SubItems(4))
                GetForm = ""
                Unload Me
        End If
    End If
End Sub

Private Sub lvSOC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lvSOC_DblClick
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdFind_Click
    End If
End Sub
