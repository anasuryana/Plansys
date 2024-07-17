VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form popUp_HistoryProc 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "History of Data"
   ClientHeight    =   3135
   ClientLeft      =   2835
   ClientTop       =   3765
   ClientWidth     =   7830
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
   ScaleHeight     =   3135
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   0
   End
End
Attribute VB_Name = "popUp_HistoryProc"
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
Dim lisitm As ListItem
Dim qry As String

Private Sub LoadDatanya()
    qry = "select full_name,datelg,partno,prod_nomach,mold_no,cavity,ct,ct_2,manpower,priorit,subcont " _
        & " from loadcap_lg_proc a " _
        & " inner join hr_employee b on a.userlg=b.empno" _
        & " where idproclclg = " & F_Mst_Product_v2.idPROC & "" _
        & " order by datelg asc"
    Set RsGet = Con.Execute(qry)
    Me.Caption = "History of Data [" & RsGet.RecordCount & " row(s) found]"
End Sub

Private Sub settingLV()
    With lv1
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "By"
        .ColumnHeaders.Add , , "When", 3000
        .ColumnHeaders.Add , , "Part No", 3000
        .ColumnHeaders.Add , , "Machine"
        .ColumnHeaders.Add , , "Mold No"
        .ColumnHeaders.Add , , "Cavity"
        .ColumnHeaders.Add , , "CT"
        .ColumnHeaders.Add , , "CT 2"
        .ColumnHeaders.Add , , "Man Power"
        .ColumnHeaders.Add , , "Priority"
        .ColumnHeaders.Add , , "Subcont"
    End With
End Sub

Private Sub getList()
    lv1.ListItems.Clear
    Do Until RsGet.EOF
        Set lisitm = lv1.ListItems.Add(, , RTrim(RsGet!full_name))
            lisitm.SubItems(1) = RTrim(RsGet!datelg)
            lisitm.SubItems(2) = RTrim(RsGet!partno)
            lisitm.SubItems(3) = IIf(IsNull(RsGet!prod_nomach), "", RsGet!prod_nomach)
            lisitm.SubItems(4) = RsGet!mold_no
            lisitm.SubItems(5) = IIf(IsNull(RsGet!cavity), 0, RsGet!cavity)
            lisitm.SubItems(6) = IIf(IsNull(RsGet!ct), 0, RsGet!ct)
            lisitm.SubItems(7) = IIf(IsNull(RsGet!ct_2), 0, RsGet!ct_2)
            lisitm.SubItems(8) = IIf(IsNull(RsGet!manpower), 0, RsGet!manpower)
            lisitm.SubItems(9) = IIf(IsNull(RsGet!priorit), 0, RsGet!priorit)
            lisitm.SubItems(10) = IIf(IsNull(RsGet!subcont), 0, RsGet!subcont)
        RsGet.MoveNext
    Loop
End Sub

Sub ResizeControls()
    On Error Resume Next
    Dim i As Integer
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

Private Sub Form_Initialize()
    Me.WindowState = vbNormal
    Dim i As Integer
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
    settingLV
    LoadDatanya
    getList
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub
