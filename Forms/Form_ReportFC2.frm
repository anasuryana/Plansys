VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_ReportFC2 
   Caption         =   "Forecast History"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11835
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
   ScaleHeight     =   6555
   ScaleWidth      =   11835
   Begin VB.CheckBox Check1 
      Caption         =   "ALL"
      Height          =   375
      Left            =   11160
      TabIndex        =   31
      Top             =   120
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "..."
      Height          =   375
      Left            =   10560
      TabIndex        =   30
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtCust 
      BackColor       =   &H00FFFFC0&
      Height          =   390
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   5400
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   120
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   0
         Width           =   4215
      End
   End
   Begin MSComCtl2.DTPicker dtfrom 
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   157155331
      CurrentDate     =   43116
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbFiletype 
      Height          =   390
      ItemData        =   "Form_ReportFC2.frx":0000
      Left            =   10080
      List            =   "Form_ReportFC2.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   600
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ComboBox cmbMonth2 
      Height          =   390
      ItemData        =   "Form_ReportFC2.frx":0023
      Left            =   3960
      List            =   "Form_ReportFC2.frx":004B
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   390
      ItemData        =   "Form_ReportFC2.frx":00B1
      Left            =   1560
      List            =   "Form_ReportFC2.frx":00D9
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox cmbThn2 
      Height          =   390
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox PicEdit 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   3240
      ScaleHeight     =   1035
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtEdit 
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   2655
      End
      Begin VB.Label Label2 
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
         Left            =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Edit"
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
         Width           =   2565
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6240
      OleObjectBlob   =   "Form_ReportFC2.frx":013F
      Top             =   480
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cmbThn 
      Height          =   390
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_ReportFC2.frx":0373
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   3600
      OleObjectBlob   =   "Form_ReportFC2.frx":03D7
      TabIndex        =   9
      Top             =   840
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_ReportFC2.frx":0433
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   3600
      OleObjectBlob   =   "Form_ReportFC2.frx":049F
      TabIndex        =   20
      Top             =   120
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   158072835
      CurrentDate     =   43116
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   6960
      OleObjectBlob   =   "Form_ReportFC2.frx":04FB
      TabIndex        =   28
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label2"
      Height          =   375
      Left            =   10080
      TabIndex        =   17
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "Form_ReportFC2"
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
Dim ar_nmBulan(1 To 12) As String
Dim i As Long
Public custID As String
Dim u As Byte
Dim rsView As ADODB.Recordset
Dim oExcel As Object
Dim oBook  As Object
Dim oSheet As Object
Dim posisisFind As Long
Dim bklik As Long
Dim kklik As Long
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

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then txtCust.Text = ""
End Sub

Private Sub cmbMonth_Click()
    cmbMonth2.ListIndex = cmbMonth.ListIndex
End Sub

Private Sub cmbMonth2_Click()
    If cmbMonth2.ListIndex < cmbMonth.ListIndex Then
        If cmbThn2.ListIndex = cmbThn.ListIndex Then
            cmbMonth2.ListIndex = cmbMonth.ListIndex
        End If
    End If
End Sub

Private Sub cmbThn_Click()
    cmbThn2.ListIndex = cmbThn.ListIndex
End Sub

Private Sub cmbThn2_Click()
    If cmbThn2.ListIndex < cmbThn.ListIndex Then
        cmbThn2.ListIndex = cmbThn.ListIndex
    End If
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

Private Sub cmdLookup_Click()
    GetForm = Me.Name
    popUp_Customer.Show 1
    Check1.Value = vbUnchecked
End Sub

Private Sub cmdView_Click()
On Error GoTo Excp
    Dim qry         As String
    Dim indexmonth  As String
    Dim ndxMonth    As String
    Dim i2          As Integer
    Dim k           As Byte
    Dim ttlBar      As Single
    Dim a           As Long
    Dim thn_        As Integer
    Dim c_prd1      As Date
    Dim c_prd2      As Date
    Dim diffMonth   As Byte
    Dim yearD       As String
    Dim yearH       As String
    Dim nMonth      As Byte
    Dim bedaBulan   As Byte
    Dim mulaibulan  As Byte
    Dim mulaitahun  As Integer
    Dim iterMnth    As Byte
    Dim iterYear    As Integer
    Dim POSs        As Byte
    Dim whereClause As String
    Dim whereDet    As String
    PB1.Visible = True
    If Check1.Value = vbUnchecked Then
        whereClause = " where a.cust_id='" & custID & "' "
        whereDet = " and a.cust_id='" & custID & "' "
    End If
    
    bedaBulan = DateDiff("m", dtfrom, dtTo) + 1
    If bedaBulan > 24 Then
        MsgBox "the maximum period of sequence of date (by month) is 12"
        dtfrom.SetFocus
        Exit Sub
    End If
    
    mulaibulan = CByte(Format(dtfrom, "m"))
    mulaitahun = CInt(Format(dtfrom, "yyyy"))
   
    
    c_prd1 = DateSerial(cmbThn, cmbMonth.ListIndex + 1, 1)
    c_prd2 = DateSerial(cmbThn2, cmbMonth2.ListIndex + 1, 1)

    diffMonth = DateDiff("m", c_prd1, c_prd2) + 1

    
    thn_ = cmbThn
    qry = "select distinct ON (a.cust_id,a.item_id) a.item_id,item_name,cust_name,a.cust_id from forecast_mod a inner join mst_item b " _
    & " on a.item_id=b.item_id inner join r_customer c on a.cust_id=c.cust_id " & whereClause & " and stscode_id = '01' order by a.cust_id asc ,a.item_id asc"
    Set RsBantu = Con.Execute(qry)
    ttlBar = RsBantu.RecordCount
    
    If RsBantu.RecordCount > 0 Then
        'qry = "select a.item_id,qty,period_h,period,a.cust_id from forecast_mod a inner join mst_item b " _
        & " on a.item_id=b.item_id where substring(period_h from 1 for 4)>='" & cmbThn & "' and " _
        & " substring(period_h from 1 for 4)<='" & cmbThn2 & "' ORDER by period_h asc, period asc"
        
        
        With grid1
            .rows = 2
            .Cols = 5
            .Cols = 5 + diffMonth '12 '
            .FixedCols = 5

            nMonth = cmbMonth.ListIndex + 1
            For u = 5 To .Cols - 1
                .TextMatrix(0, u) = thn_
                .TextMatrix(1, u) = ar_nmBulan(nMonth)
                nMonth = nMonth + 1
                If nMonth > 12 Then
                    thn_ = thn_ + 1
                    nMonth = 1
                End If
                .Col = u
                .Row = 0
                .CellBackColor = RGB(255, 212, 42)
                .Row = 1
                .CellBackColor = RGB(255, 212, 42)
            Next
            
            iterMnth = mulaibulan
            While Not RsBantu.EOF
                a = 1 + a
                iterMnth = mulaibulan
                iterYear = mulaitahun

                For u = 1 To bedaBulan  '1 to UBound(ar_nmBulan)
                    If iterMnth > 12 Then
                        iterMnth = 1
                        iterYear = iterYear + 1
                    End If
                    .rows = .rows + 1
                    .TextMatrix(.rows - 1, 0) = RsBantu("cust_id")
                    .TextMatrix(.rows - 1, 1) = RsBantu("cust_name")
                    .TextMatrix(.rows - 1, 2) = RsBantu("item_id")
                    .TextMatrix(.rows - 1, 3) = RsBantu("item_name")
                    .TextMatrix(.rows - 1, 4) = ar_nmBulan(iterMnth) & "/" & iterYear  'ar_nmBulan(U)
                    iterMnth = iterMnth + 1
                Next
                PB1.Value = (a / ttlBar) * 100
                RsBantu.MoveNext
            Wend
            qry = "select a.item_id,qty,period_h,period,a.cust_id from forecast_mod a inner join mst_item b " _
            & " on a.item_id=b.item_id where substring(period_h from 1 for 4)>='" & mulaitahun & "' and " _
            & " substring(period_h from 1 for 4)<='" & iterYear & "' " & whereDet & "  ORDER by period_h asc, period asc"
            Set rsView = Con.Execute(qry)
            
            rsView.Fields("item_id").Properties("Optimize") = True
            rsView.Fields("period_h").Properties("Optimize") = True
            rsView.Fields("cust_id").Properties("Optimize") = True
            rsView.Fields("period").Properties("Optimize") = True
            
            For i2 = 2 To .rows - 1
                POSs = InStr(1, .TextMatrix(i2, 4), "/")
                indexmonth = Right$("00" & getNumbMonth(Left(.TextMatrix(i2, 4), POSs - 1)), 2) 'possibilty 01,02
                yearH = Right$(.TextMatrix(i2, 4), 4)
'                MsgBox indexMonth & " dan " & yearH
                For k = 5 To .Cols - 1
                    yearD = .TextMatrix(0, k)
                    ndxMonth = Right$("00" & getNumbMonth(.TextMatrix(1, k)), 2)
                    'qry = "period_h='" & yearD & indexMonth & "'" _
                    & " and item_id='" & .TextMatrix(i2, 2) & "' and period='" & yearD & ndxMonth & "' and cust_id='" & .TextMatrix(i2, 0) & "'"
                    qry = "period_h='" & yearH & indexmonth & "'" _
                    & " and item_id='" & .TextMatrix(i2, 2) & "' and period='" & yearD & ndxMonth & "' and cust_id='" & .TextMatrix(i2, 0) & "'"
                    

                    rsView.Filter = qry
                    If rsView.RecordCount > 0 Then
                        .TextMatrix(i2, k) = FormatNumber(rsView("qty"), 0)
                    End If
                Next
            Next
        End With
    Else
        grid1.rows = 2
        grid1.Cols = 5
    End If
    PB1.Value = 0
    PB1.Visible = False
    Exit Sub
Excp:
    MsgBox Err.Description
End Sub

Private Function getNumbMonth(pnilai As String) As Byte
    Dim x As Byte
    For x = 1 To UBound(ar_nmBulan)
        If ar_nmBulan(x) = pnilai Then
            getNumbMonth = x
            Exit For
        End If
    Next
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
                .Col = 3
                .TopRow = xf
                posisisFind = xf
                Exit For
            End If
        Next
        If pos = 0 Then posisisFind = 2
    End With
End Sub

Private Sub dtfrom_Change()
    If dtfrom > Now Then dtfrom.Value = Now
End Sub

Private Sub dtTo_Change()
    If dtTo.Value > Now Then dtTo.Value = Now
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 70 And Shift = 2 Then
        PicFIND.Visible = True
        PicEdit.Visible = False
        txtFindNext.SetFocus
    ElseIf KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText LTrim$(grid1.Clip)
    End If
End Sub

Private Sub Label14_Click()
    PicFIND.Visible = False
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 15
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lX As Integer, lY As Single
    If Button = vbLeftButton Then
        PicFIND.Left = PicFIND.Left + (x / 15 - lX)
        PicFIND.Top = PicFIND.Top + (Y / 15 - lY)
    Else
        lX = x / 15: lY = Y / 15
    End If
End Sub

Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
MousePointer = 0
End Sub

Private Sub Label2_Click()
     PicEdit.Visible = False
     grid1.SetFocus
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 15
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lX As Integer, lY As Single
    If Button = vbLeftButton Then
        PicEdit.Left = PicEdit.Left + (x / 15 - lX)
        PicEdit.Top = PicEdit.Top + (Y / 15 - lY)
    Else
        lX = x / 15: lY = Y / 15
    End If
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 0
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub settingFG()
    With grid1
        .Cols = 5
        .rows = 3
        .FixedRows = 2
        .FixedCols = 0
        .ColAlignment(1) = flexAlignLeftCenter
        .WordWrap = True
        .MergeCells = flexMergeRestrictRows
        
        i = 0
        .TextMatrix(0, i) = "CustID"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 0
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, i)
        .Col = i
        .Row = 0
        .CellBackColor = RGB(255, 212, 42)
        .Row = 1
        .CellBackColor = RGB(255, 212, 42)
        
        i = 1
        .TextMatrix(0, i) = "Customer"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2000
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, i)
        .Col = i
        .Row = 0
        .CellBackColor = RGB(255, 212, 42)
        .Row = 1
        .CellBackColor = RGB(255, 212, 42)
        
        i = 2
        .TextMatrix(0, i) = "Part No"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, i)
        .Col = i
        .Row = 0
        .CellBackColor = RGB(255, 212, 42)
        .Row = 1
        .CellBackColor = RGB(255, 212, 42)

        i = 3
        .TextMatrix(0, i) = "Part Name"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 3000
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, i)
        .Col = i
        .Row = 0
        .CellBackColor = RGB(255, 212, 42)
        .Row = 1
        .CellBackColor = RGB(255, 212, 42)
                      
        i = 4
        .TextMatrix(0, i) = "Issue Date"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 1650
        .MergeCol(i) = True
        .TextMatrix(1, i) = .TextMatrix(0, i)
        .Col = i
        .Row = 0
        .CellBackColor = RGB(255, 212, 42)
        .Row = 1
        .CellBackColor = RGB(255, 212, 42)
        
        .MergeRow(0) = True
    End With
End Sub

Private Sub Form_Load()
    AddTab Me
    settingFG
    activeTheme Skin1, Me
    Height = 7125
    Width = 12075
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
    Dim u As Integer
    Dim tahun As Integer
    Dim dtbefore As Date
    dtbefore = DateAdd("yyyy", -1, Now)
    tahun = Format(dtbefore, "yyyy")
    For u = 1 To 3
        cmbThn.AddItem tahun
        cmbThn2.AddItem tahun
        tahun = tahun + 1
    Next
    cmbMonth.ListIndex = 0
    cmbMonth2.ListIndex = 0
    cmbThn.ListIndex = 0
    Call WheelHook(Me.hwnd)
    cmbFiletype.ListIndex = 0
    dtfrom = Now
    dtTo = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
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

Private Sub Form_Resize()
    ResizeControls
    cmbThn.Width = Label1.Width
    cmbThn.Left = Label1.Left
    cmbThn.Top = SkinLabel2.Top
    cmbMonth.Top = Label6.Top
    cmbMonth.Width = cmbThn.Width
    cmbMonth.Left = cmbThn.Left
    cmbThn2.Width = Label3.Width
    cmbThn2.Left = Label3.Left
    cmbThn2.Top = cmbThn.Top
    cmbMonth2.Top = cmbMonth.Top
    cmbMonth2.Width = cmbMonth.Width
    cmbMonth2.Left = cmbThn2.Left
    cmbFiletype.Width = Label5.Width
    cmbFiletype.Top = cmdExport.Top
    cmbFiletype.Left = Label5.Left
End Sub

Sub SelectAllText(tb As TextBox)

tb.SelStart = 0
tb.SelLength = Len(tb.Text)

End Sub

Private Sub grid1_DblClick()
    If txtCust = "" Then MsgBox "Choose customer first", vbExclamation: Exit Sub
    PicEdit.Visible = True
    bklik = grid1.Row
    kklik = grid1.Col
    If LenB(grid1.Text) > 0 Then
        txtEdit.Text = grid1.Text
        txtEdit.SetFocus
    Else
        txtEdit.SetFocus
    End If
End Sub

Private Sub txtEdit_GotFocus()
    SelectAllText txtEdit
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        PicEdit.Visible = False
    ElseIf KeyAscii = 13 Then
        If IsNumeric(txtEdit) Then
            grid1.Col = kklik
            grid1.Row = bklik
            grid1.Text = txtEdit.Text
            PicEdit.Visible = False
            txtEdit.Text = ""
            insertorUpdateDB grid1.Text, grid1.Col, grid1.Row
            grid1.SetFocus
        Else
            If Len(txtEdit) = 0 Then
                grid1.Text = 0
                PicEdit.Visible = False
            End If
        End If
    End If
End Sub

Private Sub insertorUpdateDB(pnilai As String, x As Long, Y As Long)
    Dim qry As String
    Dim bulan As String
    Dim custCod As String
    Dim bulanper As String
    Dim RsA As ADODB.Recordset
    Dim tahunD As String
    Dim tahunH As String
    Dim POSs   As Byte
    
    With grid1
        POSs = InStr(1, .TextMatrix(Y, 4), "/")
        tahunD = .TextMatrix(0, x)
        tahunH = Right$(.TextMatrix(Y, 4), 4)
        bulan = Right("00" & getNumbMonth(.TextMatrix(1, x)), 2)
        bulanper = Right$("00" & getNumbMonth(Left(.TextMatrix(Y, 4), POSs - 1)), 2) 'Right("00" & getNumbMonth(.TextMatrix(Y, 4)), 2)
        custCod = .TextMatrix(Y, 0) 'CustId
        qry = "select qty from forecast_mod " _
        & " where item_id='" & .TextMatrix(Y, 2) & "' " _
        & " and period='" & tahunD & bulan & "' " _
        & " and cust_id='" & custCod & "' and period_h='" & tahunH & bulanper & "'"
        Set RsA = New ADODB.Recordset
        Set RsA = Con.Execute(qry)
        If RsA.RecordCount > 0 Then
            qry = "update forecast_mod set inputtime=now(), qty=" & pnilai * 1 _
            & " where item_id='" & .TextMatrix(Y, 2) & "' " _
            & " and period='" & tahunD & bulan & "'" _
            & " and cust_id='" & custCod & "' and period_h='" & tahunH & bulanper & "'"
            Con.Execute qry
        Else
            qry = "INSERT INTO forecast_mod values('" & .TextMatrix(Y, 2) & "', " _
            & "'" & tahunD & bulan & "'," & .Text * 1 & ",'" & custCod & "','" & tahunH & bulanper & "',now())"
            Con.Execute qry
        End If
        Set RsA = Nothing
    End With
    
End Sub


Private Sub txtFindNext_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        PicFIND.Visible = False
    ElseIf KeyAscii = 13 Then
        Command1_Click
    End If
End Sub
