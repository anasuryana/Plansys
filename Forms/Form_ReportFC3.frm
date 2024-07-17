VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_ReportFC3 
   Caption         =   "Delivery vs SO vs Forecast"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10185
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
   ScaleHeight     =   5475
   ScaleWidth      =   10185
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbFiletype 
      Height          =   390
      ItemData        =   "Form_ReportFC3.frx":0000
      Left            =   8400
      List            =   "Form_ReportFC3.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel labelload 
      Height          =   255
      Left            =   3480
      OleObjectBlob   =   "Form_ReportFC3.frx":0023
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbThn 
      Height          =   390
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   2160
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   120
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   0
         Width           =   4215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "Form_ReportFC3.frx":007B
      Top             =   360
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8493
      _Version        =   393216
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_ReportFC3.frx":02AF
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   8400
      TabIndex        =   14
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "Form_ReportFC3"
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
            spreasheet = "et.Application"
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
        MsgBox "saved !", vbInformation, "Information"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation
    End If
End Sub

Private Sub cmdView_Click()
    Dim qry As String
    Dim r As Long
    Dim R2 As Integer
    Dim bulan As Byte
    Dim bln As Byte
    Dim ttlBaris As Long
    Dim ttlqty As Long
    
    PB1.Visible = True
    labelload.Visible = True
    qry = "select item_id,item_name,cust_name,a.cust_id from mst_item a inner join " _
    & " r_customer b on a.cust_id=b.cust_id where pfm_id='10' order by item_id asc"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        With grid1
            .Cols = 4 + 12
            .FixedCols = 4
            bulan = 1
            R2 = 1
            For r = 4 To .Cols - 1
                .ColWidth(r) = 1100
                .Col = r
                If bulan > 12 Then bulan = 1: R2 = R2 + 1
                .TextMatrix(0, r) = Left$(ar_nmBulan(bulan), 3)
                .Row = 0
                If .TextMatrix(0, r) = "Jan" Then
                    .CellBackColor = RGB(0, 170, 255)
                ElseIf .TextMatrix(0, r) = "Feb" Then
                    .CellBackColor = RGB(15, 170, 255)
                ElseIf .TextMatrix(0, r) = "Mar" Then
                    .CellBackColor = RGB(20, 170, 255)
                ElseIf .TextMatrix(0, r) = "Apr" Then
                    .CellBackColor = RGB(30, 170, 255)
                ElseIf .TextMatrix(0, r) = "May" Then
                    .CellBackColor = RGB(50, 170, 255)
                ElseIf .TextMatrix(0, r) = "Jun" Then
                    .CellBackColor = RGB(70, 170, 255)
                ElseIf .TextMatrix(0, r) = "Jul" Then
                    .CellBackColor = RGB(90, 170, 255)
                ElseIf .TextMatrix(0, r) = "Aug" Then
                    .CellBackColor = RGB(120, 170, 255)
                ElseIf .TextMatrix(0, r) = "Sep" Then
                    .CellBackColor = RGB(140, 170, 255)
                ElseIf .TextMatrix(0, r) = "Oct" Then
                    .CellBackColor = RGB(160, 170, 255)
                ElseIf .TextMatrix(0, r) = "Nov" Then
                    .CellBackColor = RGB(180, 170, 255)
                ElseIf .TextMatrix(0, r) = "Dec" Then
                    .CellBackColor = RGB(200, 170, 255)
                End If
                bulan = bulan + 1
            Next
            r = 1
            .rows = r
            While Not RsBantu.EOF
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = Trim$(RsBantu("item_id"))
                .TextMatrix(.rows - 1, 1) = RsBantu("item_name")
                .TextMatrix(.rows - 1, 2) = "FC"
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = Trim$(RsBantu("item_id"))
                .TextMatrix(.rows - 1, 1) = RsBantu("item_name")
                .TextMatrix(.rows - 1, 2) = "SO"
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = Trim$(RsBantu("item_id"))
                .TextMatrix(.rows - 1, 1) = RsBantu("item_name")
                .TextMatrix(.rows - 1, 2) = "Delivery"
                RsBantu.MoveNext
            Wend
            
        End With
    End If
   
    qry = "select a.item_id,period,sum(qty) qty,period_h from forecast_mod a  " _
    & " where substring(period from 1 for 4)='" & cmbThn & "' " _
    & " group by a.item_id,period,period_h"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        With grid1
            For bln = 1 To 12
                RsBantu.Filter = "period_h='" & cmbThn & Right("00" & bln, 2) & "'"
                If RsBantu.RecordCount > 0 Then
                    While Not RsBantu.EOF
                        For r = 1 To .rows - 1
                            If .TextMatrix(r, 2) = "FC" Then
                                If RsBantu("item_id") = .TextMatrix(r, 0) Then
                                    For R2 = 4 To .Cols - 1
                                        If checkMonth(Right$(RsBantu("period"), 2), .TextMatrix(0, R2)) Then
                                            If RsBantu("qty") > 0 Then
                                                .TextMatrix(r, R2) = FormatNumber(RsBantu("qty"), 0)
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                        RsBantu.MoveNext
                    Wend
                End If
            Next
            For r = 1 To .rows - 1
                If .TextMatrix(r, 2) = "FC" Then
                    ttlqty = 0
                    For R2 = 4 To .Cols - 1
                        If IsNumeric(.TextMatrix(r, R2)) Then
                            ttlqty = ttlqty + .TextMatrix(r, R2) * 1
                        End If
                    Next
                    .TextMatrix(r, 3) = FormatNumber(ttlqty, 0)
                End If
            Next
        End With
    End If
    labelload.Caption = "Data: SO"
    qry = "SELECT item_id,sum(soc_reqqty) qty,extract(month from soc_reqdate) bulan from soc " _
        & " where extract(year from soc_reqdate)=" & cmbThn _
        & " group by item_id,extract(year from soc_reqdate),extract(month from soc_reqdate) " _
        & " order by 3 asc"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        With grid1
            ttlBaris = .rows - 1
            For r = 1 To .rows - 1
                ttlqty = 0
                PB1.Value = (r / ttlBaris) * 100
                If .TextMatrix(r, 2) = "SO" Then
                    .Row = r
                    .Col = 2
                    .CellBackColor = RGB(255, 212, 127)
                    For R2 = 4 To .Cols - 1
                        .Row = r
                        .Col = R2
                        .CellBackColor = RGB(255, 212, 127)
                        qry = "item_id='" & .TextMatrix(r, 0) & "' and " _
                        & " bulan = " & R2 - 3
                        RsBantu.Filter = qry
                        If RsBantu.RecordCount > 0 Then
                            .TextMatrix(r, R2) = FormatNumber(RsBantu("qty"), 0)
                            ttlqty = ttlqty + RsBantu("qty")
                        End If
                    Next
                    .TextMatrix(r, 3) = FormatNumber(ttlqty, 0)
                End If
            Next
        End With
    End If
    labelload.Caption = "Data: Delivery"
    qry = "SELECT item_id,sum(sod_scanqty) qty,extract(month from inv_date) bulan from sod " _
    & " where extract(year from inv_date)=" & cmbThn _
    & " group by item_id,extract(year from inv_date),extract(month from inv_date) " _
    & " order by 3 asc"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        With grid1
            For r = 1 To .rows - 1
                ttlqty = 0
                PB1.Value = (r / ttlBaris) * 100
                If .TextMatrix(r, 2) = "Delivery" Then
                    .Row = r
                    .Col = 2
                    .CellBackColor = RGB(255, 255, 127)
                    For R2 = 4 To .Cols - 1
                        .Row = r
                        .Col = R2
                        .CellBackColor = RGB(255, 255, 127)
                        qry = "item_id='" & .TextMatrix(r, 0) & "' and " _
                        & " bulan = " & R2 - 3
                        RsBantu.Filter = qry
                        If RsBantu.RecordCount > 0 Then
                            .TextMatrix(r, R2) = FormatNumber(RsBantu("qty"), 0)
                            ttlqty = ttlqty + RsBantu("qty")
                        End If
                    Next
                    .TextMatrix(r, 3) = FormatNumber(ttlqty, 0)
                End If
            Next
        End With
    End If
    PB1.Visible = False
    labelload.Visible = False
    Set RsBantu = Nothing
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
            stringCari = LCase$(.TextMatrix(xf, 0))
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

Private Sub Form_Load()
    AddTab Me
    settingFG
    activeTheme Skin1, Me
    Height = 6045
    Width = 10425
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
        tahun = tahun + 1
    Next
    cmbFiletype.ListIndex = 0
    Call WheelHook(Me.hwnd)
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
    cmbThn.Width = Label1.Width
    cmbThn.Left = Label1.Left
    cmbThn.Top = SkinLabel1.Top
    cmbFiletype.Top = cmdExport.Top
    cmbFiletype.Left = Label2.Left
    cmbFiletype.Width = Label2.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub settingFG()
    With grid1
        .Cols = 4
        .rows = 2
        .FixedRows = 1
        .RowHeight(0) = 500
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(1) = flexAlignLeftCenter

        .MergeCells = flexMergeRestrictRows
   
        i = 0
        .TextMatrix(0, i) = "Part No"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800
        .MergeCol(i) = True

        i = 1
        .TextMatrix(0, i) = "Part Name"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 3400
        .MergeCol(i) = True
        
        i = 2
        .TextMatrix(0, i) = "."
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 1400
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = "Total"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 1400
        .MergeCol(i) = True
                      
    End With
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        PicFIND.Visible = True
        txtFindNext.SetFocus
    ElseIf KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText LTrim$(grid1.Clip)
        
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


