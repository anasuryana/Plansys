VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_PlanVsWO 
   Caption         =   "WO vs Actual"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   11160
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
      Begin ACTIVESKINLibCtl.SkinLabel slkolom 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Form_PlanVsWO.frx":0000
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmbKolom 
         Height          =   345
         ItemData        =   "Form_PlanVsWO.frx":007A
         Left            =   120
         List            =   "Form_PlanVsWO.frx":0084
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find Next"
         Height          =   375
         Left            =   3600
         TabIndex        =   17
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   0
         Width           =   4215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbFiletype 
      Height          =   345
      ItemData        =   "Form_PlanVsWO.frx":009A
      Left            =   9480
      List            =   "Form_PlanVsWO.frx":00A4
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox PicPop 
      BackColor       =   &H0080FF80&
      Height          =   2895
      Left            =   3840
      ScaleHeight     =   2835
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid gridF 
         Height          =   1935
         Left            =   45
         TabIndex        =   10
         Top             =   840
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3413
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   375
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
         Left            =   3360
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Period List"
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
         TabIndex        =   6
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.TextBox txtPeriod 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "..."
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9340
      _Version        =   393216
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.Skin SkinFD 
      Left            =   7320
      OleObjectBlob   =   "Form_PlanVsWO.frx":00BD
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_PlanVsWO.frx":02F1
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   9480
      TabIndex        =   14
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form_PlanVsWO"
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
Dim bulan As String
Dim tahun As String
Dim ar_mesin() As String
Dim ar_mold() As String
Dim ar_part() As String
Dim ar_partname() As String
Dim rsR As ADODB.Recordset
Dim oExcel As Object
Dim oBook  As Object
Dim oSheet As Object
Dim posisisFind As Long

Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhLastDayInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 0)
End Function

Private Sub inAddValue(partNo As String, partname As String, mold As String, mesin As String)
    Dim r As Long
    If (Not ar_mesin) <> -1 Then
        For r = 1 To UBound(ar_mesin)
            If partNo = ar_part(r) And ar_mold(r) = mold And mesin = ar_mesin(r) Then
            
            Exit Sub
            End If
        Next
        
        ReDim Preserve ar_mesin(1 To UBound(ar_mesin) + 1) As String
        ReDim Preserve ar_part(1 To UBound(ar_part) + 1) As String
        ReDim Preserve ar_partname(1 To UBound(ar_partname) + 1) As String
        ReDim Preserve ar_mold(1 To UBound(ar_mold) + 1) As String
        ar_mesin(UBound(ar_mesin)) = mesin
        ar_part(UBound(ar_part)) = partNo
        ar_mold(UBound(ar_mold)) = mold
        ar_partname(UBound(ar_partname)) = partname
    Else
        ReDim ar_part(1 To 1) As String
        ReDim ar_partname(1 To 1) As String
        ReDim ar_mesin(1 To 1) As String
        ReDim ar_mold(1 To 1) As String
        ar_mesin(UBound(ar_mesin)) = mesin
        ar_part(UBound(ar_part)) = partNo
        ar_mold(UBound(ar_mold)) = mold
        ar_partname(UBound(ar_partname)) = partname
    End If
End Sub

Private Sub loadperiod()
    Dim qry As String
    qry = "select extract(year from issudate) tahun,extract(month from issudate) bulan from worko" _
    & " where extract(month from issudate)::text like '%" & FilterIn(txtfind) & "%' " _
    & " group by extract(year from issudate),extract(month from issudate) " _
    & " order by 1 desc , 2 desc limit 3"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount Then
        With gridF
            .rows = 1
            .rows = RsBantu.RecordCount + 1
            i = 1
            While Not RsBantu.EOF
                .TextMatrix(i, 0) = RsBantu(0)
                .TextMatrix(i, 1) = RsBantu(1)
                i = i + 1
                RsBantu.MoveNext
            Wend
        End With
    End If
    Set RsBantu = Nothing
End Sub

Private Sub cmdfind_Click()
    If PicPop.Visible Then
        PicPop.Visible = False
    Else
        PicPop.Visible = True
        txtfind.SetFocus
    End If
End Sub

Private Sub cmdFindNext_Click()
    Dim xf As Double, pos As Integer
    Dim ttlrows As Double
    Dim stringCari As String
    Dim kolom As Byte
    If cmbKolom.ListIndex = 0 Then
        kolom = 0
    Else
        kolom = 1
    End If
    With grid1
        ttlrows = .rows - 1
        If posisisFind + 1 >= ttlrows Then
            posisisFind = 2
        Else
            posisisFind = 1 + posisisFind
        End If
        For xf = posisisFind To ttlrows
            stringCari = LCase$(.TextMatrix(xf, kolom))
            pos = InStr(stringCari, LCase(txtFindNext))
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

Private Sub cmdView_Click()
    If txtPeriod.Text = "" Then txtPeriod.SetFocus
    Label5.Visible = True
    Dim qry As String
    Dim bf_part As String
    Dim bf_mold As String
    Dim bf_mach As String
    Dim tgl_init As Date
    Dim tglakhir As Byte
    Dim k As Byte
    tgl_init = DateSerial(tahun, bulan, 1)
    Dim ro As Long
    Dim rsActual As ADODB.Recordset
    qry = "select a.partno,partname,moldno,mesinno,issudate, qty,coalesce(qtyrfg,0) qtyrfg from worko a inner join loadcap_mst_product_r b on a.partno=b.partno " _
    & " left join ( " _
    & " select item_id,lot_no,sum(qty) qtyrfg from serial_barcode_packing WHERE substring(doc_in from 6 for 7)='RFG-INJ' group by item_id,lot_no " _
    & " ) v1 on a.partno=v1.item_id and a.lotno=v1.lot_no " _
    & " where extract(year from issudate)=" & tahun & " and extract(month from issudate)=" & bulan & " order by mesinno asc, a.partno asc"
    Set rsR = Con.Execute(qry)
    
    qry = "select machine,item_id,prod_date,sum(qty) qty from serial_barcode_packing where " _
    & " extract(year from prod_date)=" & tahun & " and extract(month from prod_date)=" & bulan & " and substring(doc_in from 6 for 7)='RFG-INJ' " _
    & " group by machine,item_id,prod_date " _
    & " order by item_id asc"
    Set rsActual = Con.Execute(qry)
    
    rsActual.Fields("machine").Properties("Optimize") = True
    rsActual.Fields("item_id").Properties("Optimize") = True
    rsActual.Fields("prod_date").Properties("Optimize") = True
    rsActual.Fields("qty").Properties("Optimize") = True
    
    If rsR.RecordCount > 0 Then
        Erase ar_mesin
        Erase ar_mold
        Erase ar_part
        Erase ar_partname
        While Not rsR.EOF
            inAddValue rsR("partno"), rsR("partname"), rsR("moldno"), rsR("mesinno")
            rsR.MoveNext
        Wend

        tglakhir = CByte(Format(dhLastDayInMonth(tgl_init), "dd"))
        grid1.Cols = 6 + tglakhir
        With grid1
            .FixedCols = 6
            For i = 1 To tglakhir
                .TextMatrix(0, 5 + i) = Format(DateSerial(tahun, bulan, i), "dd-mmm")
            Next
        End With
        grid1.rows = 1
        grid1.rows = UBound(ar_mesin) * 3 + 1
        ro = 1
        With grid1
            For i = 1 To UBound(ar_mesin)
                bf_mach = ar_mesin(i)
                bf_mold = ar_mold(i)
                bf_part = ar_part(i)
                .TextMatrix(ro, 0) = ar_mesin(i)
                .TextMatrix(ro, 1) = ar_part(i)
                .TextMatrix(ro, 2) = ar_partname(i)
                .TextMatrix(ro, 3) = ar_mold(i)
                .TextMatrix(ro, 5) = "Plan"
                 For k = 1 To tglakhir
                    rsR.Filter = "partno='" & bf_part & "' and mesinno='" & bf_mach & "' and moldno='" & bf_mold & "' and issudate='" & tahun & "-" & bulan & "-" & k & "'"
                    If rsR.RecordCount > 0 Then
                        .TextMatrix(ro, 5 + k) = rsR("qty")
                        If .TextMatrix(ro, 4) = "" Then
                            .TextMatrix(ro, 4) = FormatNumber(rsR("qty"), 0)
                        Else
                            .TextMatrix(ro, 4) = FormatNumber(.TextMatrix(ro, 4) * 1 + rsR("qty"), 0)
                        End If
                    Else
                        .TextMatrix(ro, 5 + k) = 0
                    End If
                    rsActual.Filter = "item_id='" & bf_part & "' and machine='" & bf_mach & "' and prod_date='" & tahun & "-" & bulan & "-" & k & "'"
                    If rsActual.RecordCount > 0 Then
                        .TextMatrix(ro + 1, 5 + k) = rsActual("qty")
                    Else
                        .TextMatrix(ro + 1, 5 + k) = 0
                    End If
                    If k = 1 Then
                        .TextMatrix(ro + 2, 5 + k) = .TextMatrix(ro, 5 + k) * 1 - 0 '.TextMatrix(ro + 1, 5 + k) * 1
                    Else
                        .TextMatrix(ro + 2, 5 + k) = .TextMatrix(ro, 5 + k) * 1 + .TextMatrix(ro + 2, 5 + k - 1) * 1 - 0 ' .TextMatrix(ro + 1, 5 + k) * 1
                    End If
                    If k = tglakhir Then .TextMatrix(ro + 2, 4) = FormatNumber(.TextMatrix(ro + 2, 5 + k), 0)
                 Next
                ro = ro + 1
                .TextMatrix(ro, 0) = ar_mesin(i)
                .TextMatrix(ro, 1) = ar_part(i)
                .TextMatrix(ro, 2) = ar_partname(i)
                .TextMatrix(ro, 3) = ar_mold(i)
                .TextMatrix(ro, 5) = "Actual"
                ro = ro + 1
                .TextMatrix(ro, 0) = ar_mesin(i)
                .TextMatrix(ro, 1) = ar_part(i)
                .TextMatrix(ro, 2) = ar_partname(i)
                .TextMatrix(ro, 3) = ar_mold(i)
                .TextMatrix(ro, 5) = "Balance"
                ro = ro + 1
            Next
        End With
    End If
    Set rsR = Nothing
    Set rsActual = Nothing
    Label5.Visible = False
End Sub

Private Sub Command1_Click()
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
    MsgBox "ok"
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

Private Sub Form_Resize()
    ResizeControls
    cmbFiletype.Width = Label4.Width
    cmbFiletype.Left = Label4.Left
    cmbFiletype.Top = cmdfind.Top
    cmbKolom.Width = slkolom.Width
    cmbKolom.Left = slkolom.Left
    cmbKolom.Top = txtFindNext.Top
End Sub

Private Sub Form_Load()
On Error GoTo Ex
    AddTab Me
    BukaKoneksi
    Call activeTheme(skinFD, Me)
    settingFG
    Height = 6480
    Width = 11400
    cmbFiletype.ListIndex = 0
    cmbKolom.ListIndex = 0
    Call WheelHook(Me.hwnd)
    Exit Sub
Ex:
    MsgBox Err.Description
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

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub settingFG()
    With grid1
        .Cols = 6
        .rows = 2
        .FixedRows = 1
        .RowHeight(0) = 500
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter

        .MergeCells = flexMergeFree

        i = 0
        .TextMatrix(0, i) = "Machine"
        .ColWidth(i) = 900
        .ColAlignment(i) = flexAlignLeftCenter
        .MergeCol(i) = True

        i = 1
        .TextMatrix(0, i) = "Part No"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800
        .MergeCol(i) = True

        i = 2
        .TextMatrix(0, i) = "Part Name"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2500
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = "Mold No"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 1500
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = "Total"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 1300
        .ColAlignment(i) = flexAlignRightCenter
        
        i = 5
        .TextMatrix(0, i) = "."
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 900
    End With
    With gridF
        .Cols = 2
        .rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, 0) = "Year"
        .TextMatrix(0, 1) = "Month"
    End With
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 70 And Shift = 2 And grid1.rows > 2 Then
        PicFIND.Visible = True
        txtFindNext.SetFocus
    End If
End Sub

Private Sub gridF_DblClick()
    gridF_KeyPress 13
End Sub

Private Sub gridF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With gridF
            bulan = .TextMatrix(.RowSel, 1)
            tahun = .TextMatrix(.RowSel, 0)
        End With
        txtPeriod = tahun & Right("00" & bulan, 2)
        Label2_Click
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
    PicPop.Visible = False
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        loadperiod
    End If
End Sub

Private Sub txtFindNext_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFindNext_Click
    ElseIf KeyAscii = vbKeyEscape Then
        PicFIND.Visible = False
    ElseIf KeyAscii = 1 Then
        txtFindNext.SelStart = 0
        txtFindNext.SelLength = Len(txtFindNext.Text)
    End If
End Sub
