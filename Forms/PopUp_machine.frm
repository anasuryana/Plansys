VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PopUp_machine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Machine List"
   ClientHeight    =   5595
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   5595
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "PopUp_machine.frx":0000
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8705
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Machine ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Machine No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Machine Name"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tonage"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Find"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin sknFD 
      Left            =   0
      OleObjectBlob   =   "PopUp_machine.frx":005E
      Top             =   0
   End
End
Attribute VB_Name = "PopUp_machine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim listItemID  As ListItem

Private Sub getList()
    LV.ListItems.Clear
    Do Until RsGet.EOF
        Set listItemID = LV.ListItems.Add(, , RTrim(RsGet!idmst_mach))
            listItemID.SubItems(1) = RTrim(RsGet!no_mach)
            listItemID.SubItems(2) = RTrim(RsGet!name_mach)
            listItemID.SubItems(3) = RsGet!tonage_mach
        RsGet.MoveNext
    Loop
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call activeTheme(sknFD, Me)
    LV_FlatHeaders Me.hwnd, LV.hwnd
    Call BukaKoneksi
Exit Sub
errLoad:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error Load [" & Err.Number & "]"
End Sub

Private Sub LV_DblClick()
    On Error Resume Next
  
    If GetForm = "F_Mst_Product_v2" Then
        If Not Me.LV.SelectedItem Is Nothing Then
            Set listItemID = Me.LV.SelectedItem
                F_Mst_Product_v2.c_machine_no = RTrim(listItemID.SubItems(1))
                GetForm = ""
                Unload Me
        Else
            F_Mst_Product_v2.c_machine_no = ""
                GetForm = ""
                Unload Me
        End If
    End If
    If GetForm = "Form_STE" Then
        If Not Me.LV.SelectedItem Is Nothing Then
            Set listItemID = Me.LV.SelectedItem
                Form_STE.txtMachine = RTrim(listItemID.SubItems(1))
                GetForm = ""
                Unload Me
        Else
            Form_STE.txtMachine = ""
                GetForm = ""
                Unload Me
        End If
    End If
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LV_DblClick
    End If
End Sub

Private Sub OKButton_Click()
    On Error GoTo errFind
    Set RsGet = Con.Execute("select idmst_mach, no_mach, name_mach, tonage_mach from loadcap_mst_mach where upper(specification) like '%" & UCase(txtFind) & "%' ORDER BY no_mach asc")
    Call getList
    If LV.ListItems.Count > 0 Then
        LV.SetFocus
    End If
errFind:
    If Err.Number <> 0 Then
        MsgBox "Error... (" & Err.Description & ")", vbCritical, Err.Number
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OKButton_Click
    End If
End Sub
