VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_SafetyStock 
   Caption         =   "Setting Safety Stock"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12480
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   2640
      Width           =   11655
      Begin VB.CommandButton cmdFind 
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5880
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   4695
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7080
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvItem 
         Height          =   5775
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   10186
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1236
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Assy No"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Assy Name"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Safety Stock (%)"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Prod Plan (%)"
            Object.Width           =   3351
         EndProperty
      End
      Begin VB.Label lblItemFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Item ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         Caption         =   "Result :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   11595
      TabIndex        =   7
      Top             =   120
      Width           =   11655
      Begin VB.ComboBox cmbFiletype 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form_SafetyStock.frx":0000
         Left            =   9000
         List            =   "Form_SafetyStock.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtProdplan 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   18
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9000
         TabIndex        =   17
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9000
         TabIndex        =   15
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtItemId 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtQty 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   " Prod Plan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   " %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   " %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Assy No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "SAFETY STOCK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   11655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Assy Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   " Safety Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
   End
   Begin ACTIVESKINLibCtl.Skin sknFD 
      Left            =   9960
      OleObjectBlob   =   "Form_SafetyStock.frx":0023
      Top             =   5280
   End
End
Attribute VB_Name = "Form_SafetyStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim listData As ListItem
Dim oSheet As Object
        Dim oBook As Object
        Dim oExcel As Object
Private Sub cmdAll_Click()
    txtFilter = ""
    Call cmdFind_Click
End Sub

Private Sub cmdFind_Click()
On Error GoTo errFind
    Set RsGet = Con.Execute("SELECT item_id, item_name, CAST(COALESCE(prc_safetystock, '0') AS integer) prc_safetystock " _
        & ",prc_prodplan FROM mst_item WHERE upper(item_id) like '" & UCase(txtFilter) & "%' AND pfm_id = '10' " _
        & "ORDER BY item_id ASC")
    Call getList
Exit Sub
errFind:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error Find [" & Err.Number & "]"
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo errUpdate
    Dim rowAffected As Long
    If txtItemId <> "" Then
        Con.Execute "UPDATE mst_item SET prc_safetystock = " & Val(txtQty) _
            & ",prc_prodplan=" & Val(txtProdplan) & " WHERE item_id = '" & txtItemId & "'", rowAffected
        If rowAffected > 0 Then MsgBox "Update Succesfull...", vbInformation, "Update Result: " & rowAffected
        txtFilter = txtItemId
        Call cmdFind_Click
    End If
Exit Sub
errUpdate:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error Update [" & Err.Number & "]"
End Sub

Private Sub cmdUpload_Click()
    On Error GoTo Exc
    Dim spreasheet As String
    Dim qry As String
    If cmbFiletype.ListIndex = 0 Then
        spreasheet = "Excel.Application"
    Else
        spreasheet = "Ket.Application"
    End If
    
    Dim urlFILE As String, ada As Boolean
    Dim iROW As Double
    Const NamaTabel As String = "mst_item"
    Dim li As ListItem
    With CommonDialog1
        .Filter = ""
        .ShowOpen
        urlFILE = .FileName
    End With
    If urlFILE <> "" Then
        Set oExcel = CreateObject(spreasheet)
        oExcel.Workbooks.Open urlFILE
        Set oBook = oExcel.Workbooks(1)
        Set oSheet = oBook.Worksheets(1)
        ada = True
        iROW = 2
        BukaKoneksi
        Screen.MousePointer = 11
        While ada
            If oSheet.Cells(iROW, 1) <> "" Then
                qry = "UPDATE " & NamaTabel & " SET prc_safetystock = " & Val(oSheet.Cells(iROW, 2)) _
            & ",prc_prodplan=" & Val(oSheet.Cells(iROW, 3)) & " WHERE item_id = '" & oSheet.Cells(iROW, 1) & "'"
            Con.Execute qry
            Else
                ada = False
            End If
            iROW = iROW + 1
        Wend
        Screen.MousePointer = 0
        MsgBox "Uploaded !", vbInformation, "Upload Status"
        
        
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    End If
    Exit Sub
Exc:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    activeTheme sknFD, Me
    LV_FlatHeaders Me.hWnd, lvItem.hWnd
    BukaKoneksi
    cmbFiletype.ListIndex = 0
Exit Sub
errLoad:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error Load [" & Err.Number & "]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 0 Then DelTab Me
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Sub txtProdplan_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48) And (KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48) And (KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then
        KeyAscii = 0
    End If
End Sub

Private Sub lvItem_Click()
On Error GoTo errItemClick
    If lvItem.ListItems.Count <> 0 Then
    If lvItem.SelectedItem.Text <> "" Then
        txtItemId = lvItem.SelectedItem.SubItems(1)
        txtItemName = lvItem.SelectedItem.SubItems(2)
        txtQty = lvItem.SelectedItem.SubItems(3)
        txtProdplan = lvItem.SelectedItem.SubItems(4)
    End If
    End If
Exit Sub
errItemClick:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error Click [" & Err.Description & "]"
End Sub

Private Sub getList()
    lvItem.ListItems.Clear
    If Not RsGet.EOF Then
    With RsGet
        Do Until .EOF
            Set listData = lvItem.ListItems.Add(, , lvItem.ListItems.Count + 1)
                listData.SubItems(1) = RTrim(RsGet.Fields(0))
                listData.SubItems(2) = RTrim(RsGet.Fields(1))
                listData.SubItems(3) = RsGet.Fields(2)
                listData.SubItems(4) = RsGet.Fields(3)
            .MoveNext
        Loop
        lblResult = "Result : " & lvItem.ListItems.Count
    End With
    End If
End Sub

