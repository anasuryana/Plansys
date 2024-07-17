VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form popup_loadcap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Capacity Docs"
   ClientHeight    =   3570
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "popup_loadcap.frx":0000
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   735
      Left            =   50
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtFind 
         Height          =   405
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2655
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "double click or press enter to select document"
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "LTPP Document"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "popup_loadcap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public docSelcd As String

Private Sub loadData()
    Dim qry As String
    Dim li As ListItem
    Dim rsa_doc As New ADODB.Recordset
    qry = "select * from " _
    & " (select distinct on (fltpp_doc) fltpp_doc from loadcap_generate_d where fltpp_doc like '%" & txtFind & "%') v1 " _
    & " order by right(fltpp_doc,4) asc,substring(fltpp_doc from 17 for 2) "
    Set rsa_doc = Con.Execute(qry)
    lv1.ListItems.Clear
    If rsa_doc.RecordCount > 0 Then
        While Not rsa_doc.EOF
            Set li = lv1.ListItems.Add(, , rsa_doc(0))
            rsa_doc.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Wah
    activeTheme Skin1, Me
    BukaKoneksi
    Exit Sub
Wah:
    MsgBox Err.Description
End Sub

Private Sub lv1_DblClick()
    docSelcd = lv1.SelectedItem.Text
    Unload Me
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFind = FilterIn(txtFind)
        loadData
    End If
End Sub
