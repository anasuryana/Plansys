VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PopUp_MLDOC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doc List"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4260
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
   ScaleHeight     =   3195
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   45
      OleObjectBlob   =   "PopUp_MLDOC.frx":0000
      TabIndex        =   3
      Top             =   120
      Width           =   780
   End
   Begin VB.TextBox txtfind 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Doc No"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "PopUp_MLDOC.frx":0062
      Top             =   0
   End
End
Attribute VB_Name = "PopUp_MLDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private li As ListItem
Private rsa_doc As New ADODB.Recordset
Public lu_nodoc As String

Private Sub loadData()
    Dim qry As String
    qry = "select * from " _
        & " (select distinct on (fltpp_doc) fltpp_doc from mpp_gen_d ) v1 " _
        & " where fltpp_doc like '%" & txtfind & "%'" _
        & " order by right(fltpp_doc,4) asc,substring(fltpp_doc from 17 for 2) asc"
    Set rsa_doc = Con.Execute(qry)
    lv1.ListItems.Clear
    If rsa_doc.RecordCount > 0 Then
        While Not rsa_doc.EOF
            Set li = lv1.ListItems.Add(, , rsa_doc(0))
            rsa_doc.MoveNext
        Wend
    End If
End Sub

Private Sub lvToVar()
    lu_nodoc = lv1.SelectedItem.Text
End Sub

Private Sub Form_Load()
    On Error GoTo AE
    BukaKoneksi
    lu_nodoc = ""
    activeTheme skinFD, Me
    Exit Sub
AE:
    MsgBox Err.Description
End Sub

Private Sub lv1_Click()
    lvToVar
End Sub

Private Sub lv1_DblClick()
    lvToVar
    Unload Me
End Sub

Private Sub lv1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Unload Me
    End If
End Sub

Private Sub lv1_KeyUp(KeyCode As Integer, Shift As Integer)
    lvToVar
End Sub

Private Sub OKButton_Click()
    If lv1.ListItems.Count > 0 Then
        lu_nodoc = lv1.ListItems(1).Text
    End If
    Unload Me
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtfind = FilterIn(txtfind)
        loadData
    End If
End Sub
