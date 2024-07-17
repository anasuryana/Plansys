VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form popUp_Customer 
   Caption         =   "Customer"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8970
   Icon            =   "popUp_Customer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFrame 
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5475
      ScaleWidth      =   8355
      TabIndex        =   3
      Top             =   240
      Width           =   8415
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
      Begin MSComctlLib.ListView lvCustomer 
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cust Id"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cust Name"
            Object.Width           =   5010
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cust Address"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "popUp_Customer.frx":000C
      Top             =   0
   End
End
Attribute VB_Name = "popUp_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim listCustomer As ListItem

Private Sub cmdfind_Click()
On Error GoTo errFind
    Set RsGet = Con.Execute("select * from r_customer where upper(cust_name) like '%" & UCase(txtFilter) & "%'")
    Call getList
    If RsGet.RecordCount > 0 Then
        lvCustomer.SetFocus
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
    lvCustomer.ListItems.Clear
    If Not RsGet.EOF Then
        Do Until RsGet.EOF
            Set listCustomer = lvCustomer.ListItems.Add(, , RsGet!cust_id)
                listCustomer.SubItems(1) = RsGet!cust_name
                listCustomer.SubItems(2) = RsGet!cust_address
            RsGet.MoveNext
        Loop
    End If
End Sub
    
Private Sub lvCustomer_DblClick()
On Error Resume Next
    
    If GetForm = "Form_MovAVG" Then
        If Not Me.lvCustomer.SelectedItem Is Nothing Then
            Set listCustomer = Me.lvCustomer.SelectedItem
                Form_MovAVG.custID = Trim$(listCustomer.Text)
                Form_MovAVG.txtCustID = Trim$(listCustomer.SubItems(1))
                GetForm = ""
                Unload Me
        End If
    End If
    If GetForm = "Form_Forecast" Then
        If Not Me.lvCustomer.SelectedItem Is Nothing Then
            Set listCustomer = Me.lvCustomer.SelectedItem
                Form_Forecast.cust_id = Trim$(listCustomer.Text)
                Form_Forecast.txtCust = Trim$(listCustomer.SubItems(1))
                GetForm = ""
                Unload Me
        End If
    End If
    If GetForm = "Form_ReportFC2" Then
        If Not Me.lvCustomer.SelectedItem Is Nothing Then
            Set listCustomer = Me.lvCustomer.SelectedItem
                Form_ReportFC2.custID = Trim$(listCustomer.Text)
                Form_ReportFC2.txtCust = Trim$(listCustomer.SubItems(1))
                GetForm = ""
                Unload Me
        End If
    End If
    If GetForm = "Form_ReportFC4" Then
        If Not Me.lvCustomer.SelectedItem Is Nothing Then
            Set listCustomer = Me.lvCustomer.SelectedItem
                Form_ReportFC4.custID = Trim$(listCustomer.Text)
                Form_ReportFC4.txtCustomer = Trim$(listCustomer.SubItems(1))
                GetForm = ""
                Unload Me
        End If
    End If
    If GetForm = "Form_StockpCutoff" Then
        If Not Me.lvCustomer.SelectedItem Is Nothing Then
            Set listCustomer = Me.lvCustomer.SelectedItem
                Form_StockpCutoff.custID = Trim$(listCustomer.Text)
                Form_StockpCutoff.txtCust = Trim$(listCustomer.SubItems(1))
                GetForm = ""
                Unload Me
        End If
    End If
    
    If GetForm = "Form_RLoading_c1" Then
        If Not Me.lvCustomer.SelectedItem Is Nothing Then
            Set listCustomer = Me.lvCustomer.SelectedItem
                Form_RLoading_c1.cust_id = Trim$(listCustomer.Text)
                Form_RLoading_c1.txtCustomer = Trim$(listCustomer.SubItems(1))
                GetForm = ""
                Unload Me
        End If
    End If
End Sub

Private Sub lvCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lvCustomer_DblClick
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdfind_Click
    End If
End Sub

