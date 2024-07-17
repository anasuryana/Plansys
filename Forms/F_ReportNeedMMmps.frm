VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_ReportNeedMMmps 
   Caption         =   "Overloading MPS"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12420
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   12420
   Begin VB.ComboBox cmbFiletype 
      Height          =   375
      ItemData        =   "F_ReportNeedMMmps.frx":0000
      Left            =   9600
      List            =   "F_ReportNeedMMmps.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   960
      Width           =   2655
   End
   Begin VB.ComboBox CmbRevision 
      Height          =   375
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox CmbDocument 
      Height          =   375
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5955
      TabIndex        =   3
      Top             =   1440
      Width           =   6015
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Overload Production (Pcs)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6240
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   1440
      Width           =   6135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Overload Production (Day)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdExportLC 
      Caption         =   "Export"
      Height          =   735
      Left            =   9600
      TabIndex        =   0
      ToolTipText     =   "Spreadsheet"
      Top             =   120
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   50
      OleObjectBlob   =   "F_ReportNeedMMmps.frx":0023
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyyMM"
      Format          =   152567811
      CurrentDate     =   42544
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   50
      OleObjectBlob   =   "F_ReportNeedMMmps.frx":0089
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   50
      OleObjectBlob   =   "F_ReportNeedMMmps.frx":00EB
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid agrid 
      Height          =   5415
      Left            =   45
      TabIndex        =   11
      Top             =   2040
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9551
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "F_ReportNeedMMmps.frx":0151
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid angrid2 
      Height          =   5415
      Left            =   6240
      TabIndex        =   12
      Top             =   2040
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9551
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "F_ReportNeedMMmps"
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
Private RsA As ADODB.Recordset
Private rsB As ADODB.Recordset
Dim qry As String
Dim nmbulan() As String
Dim period1 As String
Dim period2 As String
Dim period3 As String
Dim period4 As String
Private oExcel      As Object 'Excel.Application
Private oBook       As Object 'Excel.Workbook
Private oSheet      As Object 'Excel.Worksheet
Dim i As Integer, j As Integer
Dim ttlMesin As Integer
Dim ttlMPPERIOD1 As Variant
Dim ttlMPPERIOD2 As Variant
Dim ttlMPPERIOD3 As Variant
Dim ttlMPPERIOD4 As Variant

Dim m_SortColumn As Integer
Dim m_SortOrder As Variant

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

Private Sub agrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText agrid.Clip
    End If
End Sub

Private Sub agrid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    MsgBox agrid.MouseRow
    If agrid.MouseRow <> 0 Then Exit Sub
    SortByColumn agrid.MouseCol
End Sub

Private Sub CmbDocument_DropDown()
    qry = "select distinct on (fltpp_doc) fltpp_doc from mpp_gen_d where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "'"
    Set RsA = Con.Execute(qry)
    CmbDocument.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            CmbDocument.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub

Private Function nmAngkakeBulan(pis As String) As String
    Dim x As Integer
    For x = 1 To UBound(nmbulan)
        If x = pis Then
            nmAngkakeBulan = nmbulan(x)
            Exit For
        End If
    Next
End Function

Private Sub formatHeaderFG()
    Dim b As Byte
    With agrid
        b = 4
        .TextMatrix(1, b) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 1) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 2) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 3) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
        
        b = 8
        .TextMatrix(1, b) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 1) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 2) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 3) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
        
        b = 12
        .TextMatrix(1, b) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 1) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 2) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, b + 3) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
    End With
    With angrid2
        .TextMatrix(0, 4) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 5) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 6) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 7) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
    End With
End Sub

Private Function setDistinctData(pitem As String) As Boolean
    For j = 1 To agrid.rows - 1
        If pitem = agrid.TextMatrix(j, 1) Then
            setDistinctData = True
            Exit For
        Else
            setDistinctData = False
        End If
    Next
End Function

' Sort by the indicated column.
Private Sub SortByColumn(ByVal sort_column As Integer)
    ' Hide the FlexGrid.
    agrid.Visible = False
    agrid.Refresh

    ' Sort using the clicked column.
    agrid.Col = sort_column
    agrid.ColSel = sort_column
    agrid.Row = 0
    agrid.RowSel = 0

    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If m_SortColumn <> sort_column Then
        m_SortOrder = flexSortGenericAscending
    ElseIf m_SortOrder = flexSortGenericAscending Then
        m_SortOrder = flexSortGenericDescending
    Else
        m_SortOrder = flexSortGenericAscending
    End If
    agrid.Sort = m_SortOrder

    ' Restore the previous sort column's name.
    If m_SortColumn >= 0 Then
        agrid.TextMatrix(0, m_SortColumn) = _
            Mid$(agrid.TextMatrix(0, m_SortColumn), 3)
    End If

    ' Display the new sort column's name.
    m_SortColumn = sort_column
    If m_SortOrder = flexSortGenericAscending Then
        agrid.TextMatrix(0, m_SortColumn) = "> " & _
            agrid.TextMatrix(0, m_SortColumn)
    Else
        agrid.TextMatrix(0, m_SortColumn) = "< " & _
            agrid.TextMatrix(0, m_SortColumn)
    End If

    ' Display the FlexGrid.
    agrid.Visible = True
End Sub

Private Sub CmbRevision_Click()
    If CmbRevision <> "" Then
        Dim capday1 As Variant
        Dim needday1 As Variant, needday2 As Variant, needday3 As Variant, needday4 As Variant
        Dim terproses1 As Double, terproses2 As Double, terproses3 As Double, terproses4 As Double
        period1 = Format(DTPicker1.Value, "yyyyMM")
        period2 = Format(DateAdd("m", 1, DTPicker1.Value), "yyyyMM") 'Left(period1, 4) & Right("00" & Val(Right(period1, 2) + 1), 2)
        period3 = Format(DateAdd("m", 2, DTPicker1.Value), "yyyyMM") 'Left(period2, 4) & Right("00" & Val(Right(period2, 2) + 1), 2)
        period4 = Format(DateAdd("m", 3, DTPicker1.Value), "yyyyMM") 'Left(period3, 4) & Right("00" & Val(Right(period3, 2) + 1), 2)
        
        qry = "select lcd_itemdid,lc_itemname,cav,ct,mpower,max(pp1) pp1,max(bln1) bln1,max(pp2) pp2,max(bln2) bln2,max(pp3) pp3,max(bln3) bln3,max(pp4) pp4,max(bln4) bln4,lc_fprodtvty from " _
        & " (select distinct on (partno) priorit,partno, cav,mpower,lc_fprodtvty,a.ct " _
        & " from mpp_gen_d a  inner join loadcap_proc b on a.lcd_itemdid=b.partno " _
        & " where fltpp_rev=" & CmbRevision & " and fltpp_doc='" & CmbDocument & "' and priorit=1 order by partno, priorit asc) viat1 " _
        & " inner join (select lcd_itemdid,lc_itemname, (case when fltpp_ym='" & period1 & "' then lc_pp end) pp1, (case when fltpp_ym='" & period1 & "' then lc_sisa_pp end) bln1, " _
        & " (case when fltpp_ym='" & period2 & "' then lc_pp end) pp2,(case when fltpp_ym='" & period2 & "' then lc_sisa_pp end) bln2, " _
        & " (case when fltpp_ym='" & period3 & "' then lc_pp end) pp3,(case when fltpp_ym='" & period3 & "' then lc_sisa_pp end) bln3, " _
        & " (case when fltpp_ym='" & period4 & "' then lc_pp end) pp4,(case when fltpp_ym='" & period4 & "' then lc_sisa_pp end) bln4 " _
        & " from mpp_gen_d where lc_sisa_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and lc_subcont='no') viat2 on viat1.partno=viat2.lcd_itemdid " _
        & " group by lcd_itemdid,lc_itemname,cav,ct,mpower,lc_fprodtvty order by lcd_itemdid asc"
       
        Set RsA = Con.Execute(qry)
        If RsA.RecordCount > 0 Then
            i = 1
            agrid.rows = 2
            angrid2.rows = 1
            agrid.rows = 3
            angrid2.rows = 2
            agrid.rows = RsA.RecordCount + i + 1
            angrid2.rows = RsA.RecordCount + i
            formatHeaderFG
            While Not RsA.EOF
                capday1 = (60 / RsA("ct")) * RsA("cav") * 7 * 3 * 60 * RsA("lc_fprodtvty")
                needday1 = RsA("bln1") / capday1
                needday2 = RsA("bln2") / capday1
                needday3 = RsA("bln3") / capday1
                needday4 = RsA("bln4") / capday1
                If IsNull(RsA!pp1) = False Then
                    If IsNull(RsA!bln1) = False Then
                        terproses1 = RsA!pp1 - RsA!bln1
                    Else
                        terproses1 = 0
                    End If
                Else
                    terproses1 = 0
                End If
                If IsNull(RsA!pp2) = False Then
                    If IsNull(RsA!bln2) = False Then
                        terproses2 = RsA!pp2 - RsA!bln2
                    Else
                        terproses2 = 0
                    End If
                Else
                    terproses2 = 0
                End If
                If IsNull(RsA!pp3) = False Then
                    If IsNull(RsA!bln3) = False Then
                        terproses3 = RsA!pp3 - RsA!bln3
                    Else
                        terproses3 = 0
                    End If
                Else
                    terproses3 = 0
                End If
                If IsNull(RsA!pp4) = False Then
                    If IsNull(RsA!bln4) = False Then
                        terproses4 = RsA!pp4 - RsA!bln4
                    Else
                        terproses4 = 0
                    End If
                Else
                    terproses4 = 0
                End If
                    With agrid
                        .TextMatrix(i + 1, 0) = i
                        .TextMatrix(i + 1, 1) = " " & RsA("lcd_itemdid")
                        .TextMatrix(i + 1, 2) = RsA("lc_itemname")
                        .TextMatrix(i + 1, 3) = "Pcs"
                        .TextMatrix(i + 1, 4) = FormatNumber(IIf(IsNull(RsA("pp1")), 0, RsA("pp1")), 0)
                        .TextMatrix(i + 1, 5) = FormatNumber(IIf(IsNull(RsA("pp2")), 0, RsA("pp2")), 0)
                        .TextMatrix(i + 1, 6) = FormatNumber(IIf(IsNull(RsA("pp3")), 0, RsA("pp3")), 0)
                        .TextMatrix(i + 1, 7) = FormatNumber(IIf(IsNull(RsA("pp4")), 0, RsA("pp4")), 0)
                        .TextMatrix(i + 1, 8) = FormatNumber(terproses1, 0)
                        .TextMatrix(i + 1, 9) = FormatNumber(terproses2, 0)
                        .TextMatrix(i + 1, 10) = FormatNumber(terproses3, 0)
                        .TextMatrix(i + 1, 11) = FormatNumber(terproses4, 0)
                        .TextMatrix(i + 1, 12) = FormatNumber(IIf(IsNull(RsA("bln1")), 0, RsA("bln1")), 0)
                        .TextMatrix(i + 1, 13) = FormatNumber(IIf(IsNull(RsA("bln2")), 0, RsA("bln2")), 0)
                        .TextMatrix(i + 1, 14) = FormatNumber(IIf(IsNull(RsA("bln3")), 0, RsA("bln3")), 0)
                        .TextMatrix(i + 1, 15) = FormatNumber(IIf(IsNull(RsA("bln4")), 0, RsA("bln4")), 0)
                    End With
                    With angrid2
                        .TextMatrix(i, 0) = i
                        .TextMatrix(i, 1) = RsA("lcd_itemdid")
                        .TextMatrix(i, 2) = RsA("lc_itemname")
                        .TextMatrix(i, 3) = "Day"
                        .TextMatrix(i, 4) = FormatNumber(needday1, 2)
                        .TextMatrix(i, 5) = FormatNumber(needday2, 2)
                        .TextMatrix(i, 6) = FormatNumber(needday3, 2)
                        .TextMatrix(i, 7) = FormatNumber(needday4, 2)
                    End With
                    i = i + 1
                
                RsA.MoveNext
                
            Wend
        End If
    End If
End Sub

Private Sub CmbRevision_DropDown()
    qry = "select distinct on (fltpp_rev) fltpp_rev from mpp_gen_d where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "' and fltpp_doc='" & CmbDocument & "' and lc_sisa_pp>0 and lc_subcont='no'"
    Set RsA = Con.Execute(qry)
    CmbRevision.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            CmbRevision.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub

Private Sub cmdExportLC_Click()
    Dim spreasheet      As String
    If cmbFiletype.ListIndex = 0 Then
        spreasheet = "Excel.Application"
    Else
        spreasheet = "Ket.Application"
    End If
    If agrid.rows < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        oSheet.Cells(1, 1) = "LTPP DOC : " & CmbDocument
        oSheet.Cells(2, 1) = "Period : " & Format(DTPicker1, "yyyyMM")
        oSheet.Cells(3, 1) = "Revision : " & CmbRevision
        oSheet.Cells(4, 1) = Label1.Caption
        oSheet.Cells(4, 16) = Label2.Caption
        oSheet.Range(oSheet.Cells(4, 1), oSheet.Cells(4, 15)).Merge
        oSheet.Range(oSheet.Cells(4, 16), oSheet.Cells(4, 22)).Merge
        oSheet.Range(oSheet.Cells(5, 4), oSheet.Cells(5, 7)).Merge
        oSheet.Range(oSheet.Cells(5, 8), oSheet.Cells(5, 11)).Merge
        oSheet.Range(oSheet.Cells(5, 12), oSheet.Cells(5, 15)).Merge
        With oSheet
            .Range(.Cells(1, 1), .Cells(5, 22)).Font.Bold = True
            .Columns(1).NumberFormat = "@"
            .Columns(16).NumberFormat = "@"
            .Range("A4").HorizontalAlignment = xlCenter
            .Range("A4").VerticalAlignment = xlCenter
            .Range("P4").HorizontalAlignment = xlCenter
            .Range("P4").VerticalAlignment = xlCenter
            .Range("D5").HorizontalAlignment = xlCenter
            .Range("D5").VerticalAlignment = xlCenter
            .Range("H5").HorizontalAlignment = xlCenter
            .Range("H5").VerticalAlignment = xlCenter
            .Range("L5").HorizontalAlignment = xlCenter
            .Range("L5").VerticalAlignment = xlCenter
        End With
        
        Dim baris As Integer, k As Integer
        baris = 5
        With agrid
            For i = 0 To .rows - 1
                oSheet.Cells(baris, 1) = LTrim(.TextMatrix(i, 1))
                oSheet.Cells(baris, 2) = .TextMatrix(i, 2)
                oSheet.Cells(baris, 3) = .TextMatrix(i, 3)
                If i = 1 Then
                    oSheet.Cells(baris, 4) = DTPicker1.Value
                    oSheet.Cells(baris, 5) = DateAdd("m", 1, DTPicker1.Value)
                    oSheet.Cells(baris, 6) = DateAdd("m", 2, DTPicker1.Value)
                    oSheet.Cells(baris, 7) = DateAdd("m", 3, DTPicker1.Value)
                    oSheet.Cells(baris, 8) = DTPicker1.Value
                    oSheet.Cells(baris, 9) = DateAdd("m", 1, DTPicker1.Value)
                    oSheet.Cells(baris, 10) = DateAdd("m", 2, DTPicker1.Value)
                    oSheet.Cells(baris, 11) = DateAdd("m", 3, DTPicker1.Value)
                    oSheet.Cells(baris, 12) = DTPicker1.Value
                    oSheet.Cells(baris, 13) = DateAdd("m", 1, DTPicker1.Value)
                    oSheet.Cells(baris, 14) = DateAdd("m", 2, DTPicker1.Value)
                    oSheet.Cells(baris, 15) = DateAdd("m", 3, DTPicker1.Value)
                    For k = 4 To 15
                        oSheet.Cells(baris, k).NumberFormat = "mmm-yy"
                    Next
                Else
                    oSheet.Cells(baris, 4) = .TextMatrix(i, 4)
                    oSheet.Cells(baris, 5) = .TextMatrix(i, 5)
                    oSheet.Cells(baris, 6) = .TextMatrix(i, 6)
                    oSheet.Cells(baris, 7) = .TextMatrix(i, 7)
                    oSheet.Cells(baris, 8) = .TextMatrix(i, 8)
                    oSheet.Cells(baris, 9) = .TextMatrix(i, 9)
                    oSheet.Cells(baris, 10) = .TextMatrix(i, 10)
                    oSheet.Cells(baris, 11) = .TextMatrix(i, 11)
                    oSheet.Cells(baris, 12) = .TextMatrix(i, 12)
                    oSheet.Cells(baris, 13) = .TextMatrix(i, 13)
                    oSheet.Cells(baris, 14) = .TextMatrix(i, 14)
                    oSheet.Cells(baris, 15) = .TextMatrix(i, 15)
                End If
                baris = baris + 1
            Next
        End With
        baris = 5
        With angrid2
            For i = 0 To .rows - 1
                oSheet.Cells(baris, 16) = .TextMatrix(i, 1)
                oSheet.Cells(baris, 17) = .TextMatrix(i, 2)
                oSheet.Cells(baris, 18) = .TextMatrix(i, 3)
                If i = 0 Then
                    oSheet.Cells(baris, 19) = DTPicker1.Value
                    oSheet.Cells(baris, 20) = DateAdd("m", 1, DTPicker1.Value)
                    oSheet.Cells(baris, 21) = DateAdd("m", 2, DTPicker1.Value)
                    oSheet.Cells(baris, 22) = DateAdd("m", 3, DTPicker1.Value)
                    For k = 19 To 22
                        oSheet.Cells(baris, k).NumberFormat = "mmm-yy"
                    Next
                Else
                    oSheet.Cells(baris, 19) = .TextMatrix(i, 4)
                    oSheet.Cells(baris, 20) = .TextMatrix(i, 5)
                    oSheet.Cells(baris, 21) = .TextMatrix(i, 6)
                    oSheet.Cells(baris, 22) = .TextMatrix(i, 7)
                End If
                baris = baris + 1
            Next
        End With
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub settingLV()
    With agrid
        .Cols = 16: .ColWidth(0) = 700: .ColWidth(1) = 2500: .ColWidth(2) = 2500
        .rows = 3
        .FixedRows = 2
        .FixedCols = 1
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "No": .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 1
        .TextMatrix(0, i) = "Assy No": .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 2
        .TextMatrix(0, i) = "Assy Name": .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = "Unit": .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .MergeRow(0) = True
        
        i = 4
        .TextMatrix(0, i) = "Prod Plan Qty"
        .Col = i
        .Row = 0
        .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(0, i + 1) = .TextMatrix(0, i)
        .TextMatrix(0, i + 2) = .TextMatrix(0, i)
        .TextMatrix(0, i + 3) = .TextMatrix(0, i)
        
        i = 8
        .TextMatrix(0, i) = "Processed Qty"
        .Col = i
        .Row = 0
        .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(0, i + 1) = .TextMatrix(0, i)
        .TextMatrix(0, i + 2) = .TextMatrix(0, i)
        .TextMatrix(0, i + 3) = .TextMatrix(0, i)
        
        i = 12
        .TextMatrix(0, i) = "Unprocessed Qty"
        .Col = i
        .Row = 0
        .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(0, i + 1) = .TextMatrix(0, i)
        .TextMatrix(0, i + 2) = .TextMatrix(0, i)
        .TextMatrix(0, i + 3) = .TextMatrix(0, i)
    End With
    
    With angrid2
        .Cols = 8: .ColWidth(0) = 700: .ColWidth(1) = 2500: .ColWidth(2) = 2500
        .rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "No":
        
        i = 1
        .TextMatrix(0, i) = "Assy No":
        
        i = 2
        .TextMatrix(0, i) = "Assy Name"
        
        i = 3
        .TextMatrix(0, i) = "Unit"
        
        .MergeRow(0) = True
    End With
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
 On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call settingLV
    Call activeTheme(skinFD, Me)
    Me.Height = 8100
    Me.Width = 12675
    ReDim nmbulan(1 To 12) As String
    nmbulan(1) = "Jan"
    nmbulan(2) = "Feb"
    nmbulan(3) = "Mar"
    nmbulan(4) = "Apr"
    nmbulan(5) = "May"
    nmbulan(6) = "Jun"
    nmbulan(7) = "Jul"
    nmbulan(8) = "Aug"
    nmbulan(9) = "Sep"
    nmbulan(10) = "Oct"
    nmbulan(11) = "Nov"
    nmbulan(12) = "Dec"
    DTPicker1.Value = Now
    cmbFiletype.ListIndex = 0
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    With CmbDocument
        .Left = DTPicker1.Left: .Top = SkinLabel1.Top
    End With
    With CmbRevision
        .Left = DTPicker1.Left: .Top = SkinLabel3.Top
    End With
    cmbFiletype.Left = cmdExportLC.Left
    cmbFiletype.Top = CmbRevision.Top
    cmbFiletype.Width = cmdExportLC.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub Label1_Click()
    MsgBox agrid.rows
End Sub


