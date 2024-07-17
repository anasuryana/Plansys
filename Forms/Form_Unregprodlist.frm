VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_Unregprodlist 
   Caption         =   "Unregistered Product List"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7995
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   7995
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4695
      Left            =   45
      TabIndex        =   1
      Top             =   960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8281
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   855
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtFind 
         Height          =   390
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   3975
      End
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_Unregprodlist.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "Form_Unregprodlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type CtrlProportion
    Heightproportion As Single
    WidthProportion As Single
    TopProportion As Single
    LeftProportion As Single
End Type
Dim proportionArray() As CtrlProportion
Dim i As Integer
Dim qry As String

Sub ResizeControls()
    On Error Resume Next
    For i = 0 To Controls.Count - 1
        With proportionArray(i)
            ' move and resize controls
            Controls(i).Move .LeftProportion * ScaleWidth, _
            .TopProportion * ScaleHeight, _
            .WidthProportion * ScaleWidth, _
            .Heightproportion * ScaleHeight
        End With
    Next
End Sub

Private Sub loadData()
    qry = "select item_id,item_name,item_regdate from loadcap_mst_product_r a " _
        & " right join mst_item b on a.partno=b.item_id " _
        & " where a.partno is null and pfm_id='10' and left(item_id,4)!='TEST' order by item_regdate desc"
    Set RsGet = Con.Execute(qry)
    If RsGet.RecordCount = 0 Then Unload Me
End Sub

Private Sub getList()
    If RsGet.RecordCount > 0 Then
        With grid1
            .rows = 1
            .rows = RsGet.RecordCount + 1
            i = 1
            While Not RsGet.EOF
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Trim(RsGet("item_id"))
                .TextMatrix(i, 2) = RsGet("item_name")
                .TextMatrix(i, 3) = Format(RsGet("item_regdate"), "dd MMM yyyy")
                i = 1 + i
                RsGet.MoveNext
            Wend
        End With
    Else
        
    End If
End Sub

Private Sub settingGrid()
    With grid1
        .rows = 2
        .Cols = 4
        .TextMatrix(0, 0) = "No"
        .ColWidth(0) = 500
        .TextMatrix(0, 1) = "Part Number"
        .ColWidth(1) = 2500
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Part Name"
        .ColWidth(2) = 3500
        .TextMatrix(0, 3) = "Registered Date"
        .ColWidth(3) = 2500
    End With
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub Form_Initialize()
    Me.WindowState = vbNormal
    On Error Resume Next
    
    ReDim proportionArray(0 To Controls.Count - 1)
    
    For i = 0 To Controls.Count - 1
         With proportionArray(i)
            .Heightproportion = Controls(i).Height / ScaleHeight
            .WidthProportion = Controls(i).Width / ScaleWidth
            .TopProportion = Controls(i).Top / ScaleHeight
            .LeftProportion = Controls(i).Left / ScaleWidth
         End With
    Next

End Sub

Private Sub Form_Load()
    AddTab Me
    settingGrid
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call loadData
    Call getList
    Me.Width = 8235
    Me.Height = 6255
    
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText grid1.Clip
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
On Error GoTo errKu
    If KeyAscii = 13 Then
        If txtFind = "" Then
            RsGet.Filter = adFilterNone
        Else
            txtFind = FilterIn(txtFind)
            RsGet.Filter = adFilterNone
            RsGet.Filter = "item_id LIKE '*" & txtFind & "*'"
            If RsGet.RecordCount = 0 Then
                RsGet.Filter = adFilterNone
                RsGet.Filter = "item_name LIKE '*" & txtFind & "*'"
            End If
        End If
        getList
    End If
    Exit Sub
errKu:
    loadData
End Sub
