VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_ReportFC4 
   Caption         =   "Actual Forecast /Customer"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11235
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
   ScaleHeight     =   6420
   ScaleWidth      =   11235
   Begin VB.CheckBox ckcUST 
      Caption         =   "All"
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   600
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtCustomer 
      Height          =   390
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   3600
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find Next"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbFiletype 
      Height          =   390
      ItemData        =   "Form_ReportFC4.frx":0000
      Left            =   9480
      List            =   "Form_ReportFC4.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   157089795
      CurrentDate     =   43117
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7440
      OleObjectBlob   =   "Form_ReportFC4.frx":0023
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9340
      _Version        =   393216
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_ReportFC4.frx":0257
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   157089795
      CurrentDate     =   43117
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   3480
      OleObjectBlob   =   "Form_ReportFC4.frx":02BB
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_ReportFC4.frx":0317
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Label2"
      Height          =   375
      Left            =   9480
      TabIndex        =   14
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form_ReportFC4"
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
Dim i As Byte
Dim ar_nmBulan(1 To 12) As String
Dim posisisFind As Long
Dim oExcel As Object
Dim oBook  As Object
Dim oSheet As Object
Public custID As String

Private Sub ckcUST_Click()
    txtCustomer = ""
End Sub

Private Sub cmdExport_Click()
Clipboard.Clear
    With grid1
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = .rows - 1
        Clipboard.SetText .Clip
    End With
    If grid1.rows < 2 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Dim spreasheet      As String
        If cmbFiletype.ListIndex = 0 Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        oSheet.Cells(1, 1) = "Time Export : " & Now
        With oExcel.ActiveWorkbook.ActiveSheet
            .Range("A2").Select 'Select Cell A1 (will paste from here, to different cells)
            .Paste              'Paste clipboard contents
        End With
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName ', xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub cmdView_Click()
    Dim qry As String
    Dim r As Long
    Dim R2 As Integer
    Dim bulan As Byte
    Dim bln As Byte
    Dim bedaBulan As Byte
    Dim tahunA  As Integer
    Dim tahunI  As Integer
    Dim bulanA  As String
    Dim bulanB  As String
    Dim it      As Integer
    Dim ik      As Byte
    Dim itTotal As Long
    Dim totalR  As Long
    Dim iter    As Long
    Dim wCust   As String
    
    If ckcUST.Value = vbUnchecked And txtCustomer <> "" Then
        wCust = " and v1.cust_id='" & custID & "' "
    End If
    
    tahunA = Format(dtfrom, "yyyy")
    bulanA = Format(dtfrom, "MM")
    bulanB = Format(dtTo, "MM")
    bedaBulan = DateDiff("m", dtfrom, dtTo) + 1 'digit, rata kanan, merge
    
    'qry = "select item_id,item_name,cust_name,a.cust_id from mst_item a inner join " _
    & " r_customer b on a.cust_id=b.cust_id where pfm_id='10' order by item_id asc"
    qry = "select v1.cust_id,cust_name,v1.item_id,item_name from mst_item a left join (select distinct on (cust_id,item_id) cust_id,item_id from forecast_mod where qty>0) v1 on a.item_id=v1.item_id " _
    & " inner join r_customer c on v1.cust_id=c.cust_id " _
     & " where pfm_id='10' and v1.item_id not like '%TEST%' " & wCust & " order by cust_name asc"
    
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        r = 2
        With grid1
            .rows = r
            .rows = r + RsBantu.RecordCount
            While Not RsBantu.EOF
                .TextMatrix(r, 0) = Trim$(RsBantu("cust_id"))
                .TextMatrix(r, 1) = RsBantu("cust_name")
                .TextMatrix(r, 2) = RsBantu("item_id")
                .TextMatrix(r, 3) = RsBantu("item_name")
                
                r = r + 1
                RsBantu.MoveNext
            Wend
            .Cols = 4 + bedaBulan
            .FixedCols = 2
            bulan = bulanA
            tahunI = tahunA
            For r = 4 To .Cols - 1
                .ColWidth(r) = 1600
                .Col = r
                If bulan > 12 Then
                    bulan = 1:
                    tahunI = tahunI + 1
                End If
                .TextMatrix(0, r) = tahunI
                .TextMatrix(1, r) = Left$(ar_nmBulan(bulan), 3)
                .Row = 1
                If .TextMatrix(1, r) = "Jan" Then
                    .CellBackColor = RGB(0, 170, 255)
                ElseIf .TextMatrix(1, r) = "Feb" Then
                    .CellBackColor = RGB(15, 170, 255)
                ElseIf .TextMatrix(1, r) = "Mar" Then
                    .CellBackColor = RGB(20, 170, 255)
                ElseIf .TextMatrix(1, r) = "Apr" Then
                    .CellBackColor = RGB(30, 170, 255)
                ElseIf .TextMatrix(1, r) = "May" Then
                    .CellBackColor = RGB(50, 170, 255)
                ElseIf .TextMatrix(1, r) = "Jun" Then
                    .CellBackColor = RGB(70, 170, 255)
                ElseIf .TextMatrix(1, r) = "Jul" Then
                    .CellBackColor = RGB(90, 170, 255)
                ElseIf .TextMatrix(1, r) = "Aug" Then
                    .CellBackColor = RGB(120, 170, 255)
                ElseIf .TextMatrix(1, r) = "Sep" Then
                    .CellBackColor = RGB(140, 170, 255)
                ElseIf .TextMatrix(1, r) = "Oct" Then
                    .CellBackColor = RGB(160, 170, 255)
                ElseIf .TextMatrix(1, r) = "Nov" Then
                    .CellBackColor = RGB(180, 170, 255)
                ElseIf .TextMatrix(1, r) = "Dec" Then
                    .CellBackColor = RGB(200, 170, 255)
                End If
                bulan = bulan + 1
            Next
        End With
    End If

    qry = "select cust_id,a.item_id,period,sum(qty) qty,period_h from forecast_mod a  " _
    & " where period>='" & tahunA & bulanA & "' and period<='" & tahunI & bulanB & "' and coalesce(qty,0)>0" _
    & " group by a.item_id,period,period_h,cust_id"
   
    Set RsBantu = Con.Execute(qry)
    totalR = RsBantu.RecordCount
    If totalR > 0 Then
        iter = 1
        With grid1
            While Not RsBantu.EOF
                For it = 2 To .rows - 1
                    For ik = 4 To .Cols - 1
                        If RsBantu("cust_id") = .TextMatrix(it, 0) And RsBantu("item_id") = .TextMatrix(it, 2) And RsBantu("period") = .TextMatrix(0, ik) & NameToID(.TextMatrix(1, ik)) _
                           And RsBantu("qty") > 0 Then
                            .TextMatrix(it, ik) = FormatNumber(RsBantu("qty"), 0)
                        End If
                    Next
                Next
                
                pb.Value = (iter / totalR) * 100
                pb.ToolTipText = pb.Value & "%"
                iter = iter + 1
                RsBantu.MoveNext
            Wend
            .rows = .rows + 1
            .TextMatrix(.rows - 1, 0) = "Total"
            .TextMatrix(.rows - 1, 1) = "Total"
            .MergeRow(.rows - 1) = True
            For ik = 4 To .Cols - 1
                itTotal = 0
                For it = 2 To .rows - 2
                    If IsNumeric(.TextMatrix(it, ik)) Then
                        itTotal = itTotal + (.TextMatrix(it, ik) * 1)
                    End If
                Next
                .TextMatrix(.rows - 1, ik) = FormatNumber(itTotal, 0)
            Next
        End With
    End If
End Sub

Private Function checkMonth(prNo As String, bln As String) As Boolean
    Dim rsC As Byte
    Dim hasil As Boolean
    hasil = False
    For rsC = 1 To UBound(ar_nmBulan)
        If rsC = Val(prNo) And bln = Left$(ar_nmBulan(rsC), 3) Then
            hasil = True
        End If
    Next
    checkMonth = hasil
End Function

Private Sub Command1_Click()
    Dim xf As Double, pos As Integer
    Dim ttlrows As Double
    Dim stringCari As String
    With grid1
        ttlrows = .rows - 1
        If posisisFind + 1 >= ttlrows Then
            posisisFind = 2
        Else
            posisisFind = 1 + posisisFind
        End If
        For xf = posisisFind To ttlrows
            stringCari = LCase$(.TextMatrix(xf, 2))
            pos = InStr(stringCari, LCase$(txtFindNext))
            If pos > 0 Then
                .Row = xf
                .Col = 2
                .TopRow = xf
                posisisFind = xf
                Exit For
            End If
        Next
        If pos = 0 Then posisisFind = 2
    End With
End Sub

Private Sub Command2_Click()
    GetForm = Me.Name
    popUp_Customer.Show 1
    ckcUST.Value = vbUnchecked
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

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

Private Function NameToID(nmbulan As String) As String
    Dim s As Byte
    For s = 1 To 12
        If nmbulan = Left$(ar_nmBulan(s), 3) Then
            NameToID = Right$("00" & s, 2)
            Exit For
        End If
    Next
End Function

Private Sub Form_Load()
    AddTab Me
    settingFG
    activeTheme Skin1, Me
    Height = 6990
    Width = 11475
    ar_nmBulan(1) = "January"
    ar_nmBulan(2) = "February"
    ar_nmBulan(3) = "March"
    ar_nmBulan(4) = "April"
    ar_nmBulan(5) = "May"
    ar_nmBulan(6) = "Juni"
    ar_nmBulan(7) = "July"
    ar_nmBulan(8) = "August"
    ar_nmBulan(9) = "September"
    ar_nmBulan(10) = "October"
    ar_nmBulan(11) = "November"
    ar_nmBulan(12) = "December"
    cmbFiletype.ListIndex = 0
    Call WheelHook(Me.hwnd)
    dtfrom = Now
    dtTo = Now
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hwnd, xpos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, xpos, Ypos
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbFiletype.Width = Label5.Width
    cmbFiletype.Top = cmdView.Top
    cmbFiletype.Left = Label5.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub settingFG()
    With grid1
        .Cols = 4
        .rows = 3
        .FixedRows = 2
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(1) = flexAlignLeftCenter

        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        
         i = 0
        .TextMatrix(0, i) = "Custid"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 0
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, 0)
        
         i = 1
        .TextMatrix(0, i) = "Customer"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, 1)
        
        i = 2
        .TextMatrix(0, i) = "Part No"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, 2)
        
        

        i = 3
        .TextMatrix(0, i) = "Part Name"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 3400
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, 3)
                      
    End With
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        PicFIND.Visible = True
        txtFindNext.SetFocus
    ElseIf KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText grid1.Clip
        MsgBox "Copied"
    End If
End Sub

Private Sub Label14_Click()
    PicFIND.Visible = False
End Sub

Private Sub txtFindNext_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        PicFIND.Visible = False
    ElseIf KeyAscii = 13 Then
        Command1_Click
    End If
End Sub


