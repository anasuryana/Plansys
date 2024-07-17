VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_ReportNeedMM 
   Caption         =   "Overload Production"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12435
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
   ScaleWidth      =   12435
   Begin VB.PictureBox Picture3 
      Height          =   4935
      Left            =   4680
      ScaleHeight     =   4875
      ScaleWidth      =   5835
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
      Begin MSComctlLib.ListView lvex 
         Height          =   3495
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CheckBox ckShow 
      Caption         =   "Show Recap based on Tonage"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   960
      Width           =   3375
   End
   Begin VB.PictureBox picrecapTonage 
      BackColor       =   &H00C0FFFF&
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5835
      ScaleWidth      =   12075
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   12135
      Begin VB.CommandButton Command1 
         Caption         =   "Export"
         Height          =   375
         Left            =   11040
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid gridku 
         Height          =   2295
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4048
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid gridku2 
         Height          =   2895
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5106
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin MSFlexGridLib.MSFlexGrid gridku3 
         Height          =   2895
         Left            =   5160
         TabIndex        =   20
         Top             =   2880
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5106
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Recapitulation based on Tonage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   12135
      End
   End
   Begin VB.ComboBox cmbFiletype 
      Height          =   375
      ItemData        =   "F_ReportNeedMM.frx":0000
      Left            =   9600
      List            =   "F_ReportNeedMM.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdExportLC 
      Caption         =   "Export"
      Height          =   735
      Left            =   9600
      TabIndex        =   12
      ToolTipText     =   "spreadsheet file"
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6240
      ScaleHeight     =   435
      ScaleWidth      =   5955
      TabIndex        =   10
      Top             =   1440
      Width           =   6015
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
         TabIndex        =   11
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5955
      TabIndex        =   7
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
         TabIndex        =   8
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.ComboBox CmbDocument 
      Height          =   375
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.ComboBox CmbRevision 
      Height          =   375
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   735
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
      OleObjectBlob   =   "F_ReportNeedMM.frx":0023
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
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
      Format          =   152829955
      CurrentDate     =   42544
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   50
      OleObjectBlob   =   "F_ReportNeedMM.frx":0089
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   50
      OleObjectBlob   =   "F_ReportNeedMM.frx":00EB
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid agrid 
      Height          =   5415
      Left            =   45
      TabIndex        =   6
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
      OleObjectBlob   =   "F_ReportNeedMM.frx":0151
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid angrid2 
      Height          =   5415
      Left            =   6240
      TabIndex        =   9
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
Attribute VB_Name = "F_ReportNeedMM"
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

Dim rsOVRLD As New ADODB.Recordset

Dim ovhkw1 As Byte
Dim ovhkw2 As Byte
Dim ovhkw3 As Byte
Dim ovhkw4 As Byte

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

Private Sub agrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText Trim(agrid.Clip)
    End If
End Sub

Private Sub agrid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    MsgBox agrid.MouseRow
    If agrid.MouseRow <> 0 Then Exit Sub
    SortByColumn agrid.MouseCol
End Sub



Private Sub CmbDocument_Click()
    qry = "SELECT hkw_1,hkw_2,hkw_3,hkw_4 FROM ltpp_header WHERE ltpp_doc='" & CmbDocument & "'"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        ovhkw1 = RsBantu("hkw_1")
        ovhkw2 = RsBantu("hkw_2")
        ovhkw3 = RsBantu("hkw_3")
        ovhkw4 = RsBantu("hkw_4")
    End If
    
End Sub

Private Sub CmbDocument_DropDown()
    qry = "select distinct on (fltpp_doc) fltpp_doc from loadcap_generate_h where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "'"
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
    With agrid
        .TextMatrix(0, 4) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 5) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 6) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 7) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
    End With
    With angrid2
        .TextMatrix(0, 4) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 5) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 6) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(0, 7) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
    End With
    With gridku
        .TextMatrix(0, 9) = Format(DTPicker1.Value, "mmm-yy") & " (%)"
        .TextMatrix(0, 10) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") & " (%)"
        .TextMatrix(0, 11) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") & " (%)"
        .TextMatrix(0, 12) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") & " (%)"
    End With
    With gridku2
        .TextMatrix(1, 2) = Format(DTPicker1.Value, "mmm-yy") & " (%)"
        .TextMatrix(1, 3) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") & " (%)"
        .TextMatrix(1, 4) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") & " (%)"
        .TextMatrix(1, 5) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") & " (%)"
    End With
    
    With gridku3
        .TextMatrix(1, 2) = Format(DTPicker1.Value, "mmm-yy") & " (%)"
        .TextMatrix(1, 3) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") & " (%)"
        .TextMatrix(1, 4) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") & " (%)"
        .TextMatrix(1, 5) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") & " (%)"
        
        .TextMatrix(1, 6) = Format(DTPicker1.Value, "mmm-yy")
        .TextMatrix(1, 7) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy")
        .TextMatrix(1, 8) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy")
        .TextMatrix(1, 9) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy")
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
        Dim PN_to_FD As String
        Dim c_cap_p_day As Double
        period1 = Format(DTPicker1.Value, "yyyyMM")
        period2 = Format(DateAdd("m", 1, DTPicker1.Value), "yyyyMM") 'Left(period1, 4) & Right("00" & Val(Right(period1, 2) + 1), 2)
        period3 = Format(DateAdd("m", 2, DTPicker1.Value), "yyyyMM") 'Left(period2, 4) & Right("00" & Val(Right(period2, 2) + 1), 2)
        period4 = Format(DateAdd("m", 3, DTPicker1.Value), "yyyyMM") 'Left(period3, 4) & Right("00" & Val(Right(period3, 2) + 1), 2)
        

        qry = "select lc_itemid,lc_itemname,cavity,ct,manpower,bln1,bln2, bln3, bln4,lc_fprodtvty,prod_nomach from (select lc_itemid,lc_itemname,cavity,ct,manpower,sum(bln1) bln1,sum(bln2) bln2,sum(bln3) bln3,sum(bln4) bln4,lc_fprodtvty from " _
        & " (select distinct on (partno) priorit,partno, cavity,ct,manpower,lc_fprodtvty from loadcap_generate_h a  inner join loadcap_proc b on a.lc_itemid=b.partno " _
        & " where fltpp_rev=" & CmbRevision & " and fltpp_doc='" & CmbDocument & "' and priorit=1 " _
        & " order by partno, priorit asc) viat1 " _
        & " Inner Join " _
        & " (select lc_itemid,lc_itemname, (case when fltpp_ym='" & period1 & "' then lc_sisa_pp end) bln1, " _
        & " (case when fltpp_ym='" & period2 & "' then lc_sisa_pp end) bln2, " _
        & " (case when fltpp_ym='" & period3 & "' then lc_sisa_pp end) bln3, " _
        & " (case when fltpp_ym='" & period4 & "' then lc_sisa_pp end) bln4   from loadcap_generate_h " _
        & " where lc_sisa_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and lc_subcont='no') viat2 on viat1.partno=viat2.lc_itemid " _
        & " group by lc_itemid,lc_itemname,cavity,ct,manpower,lc_fprodtvty" _
        & " order by lc_itemid asc) xv1 left join (" _
        & " select distinct on ( v1.partno) v1.partno, prod_nomach from " _
        & " (select partno,max(priorit) priorit from loadcap_proc " _
        & " group by partno) v1 inner join loadcap_proc a on v1.partno=a.partno and v1.priorit=a.priorit ) xv2 on xv1.lc_itemid=xv2.partno order by lc_itemid"
        
        Set RsA = Con.Execute(qry)
        If RsA.RecordCount > 0 Then
            i = 1
            agrid.rows = 1
            angrid2.rows = 1
            agrid.rows = 2
            angrid2.rows = 2
            agrid.rows = RsA.RecordCount + i
            angrid2.rows = RsA.RecordCount + i
            formatHeaderFG
            While Not RsA.EOF
                capday1 = (60 / RsA("ct")) * RsA("cavity") * 7 * 3 * 60 * RsA("lc_fprodtvty")
                needday1 = RsA("bln1") / capday1
                needday2 = RsA("bln2") / capday1
                needday3 = RsA("bln3") / capday1
                needday4 = RsA("bln4") / capday1
                PN_to_FD = PN_to_FD & "'" & Trim(RsA("lc_itemid")) & "',"
                    With agrid
                        .TextMatrix(i, 0) = i
                        .TextMatrix(i, 1) = " " & RsA("lc_itemid")
                        .TextMatrix(i, 2) = RsA("lc_itemname")
                        .TextMatrix(i, 3) = "Pcs"
                        .TextMatrix(i, 4) = FormatNumber(IIf(IsNull(RsA("bln1")), 0, RsA("bln1")), 0)
                            .TextMatrix(i, 5) = FormatNumber(IIf(IsNull(RsA("bln2")), 0, RsA("bln2")), 0)

                            .TextMatrix(i, 6) = FormatNumber(IIf(IsNull(RsA("bln3")), 0, RsA("bln3")), 0)
                            .TextMatrix(i, 7) = FormatNumber(IIf(IsNull(RsA("bln4")), 0, RsA("bln4")), 0)
                            .TextMatrix(i, 8) = IIf(IsNull(RsA("prod_nomach")), "", RsA("prod_nomach"))

                    End With
                    With angrid2
                        .TextMatrix(i, 0) = i
                        .TextMatrix(i, 1) = RsA("lc_itemid")
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
            PN_to_FD = Left(PN_to_FD, Len(PN_to_FD) - 1)
            qry = "SELECT distinct on(partno) partno,prod_nomach,cavity,hour_p_shift,shift_usg,faktor_productivity,tonage_mach,ct FROM " _
            & " (SELECT distinct on(a.partno,prod_nomach) a.partno,prod_nomach,cavity,hour_p_shift,shift_usg,faktor_productivity,tonage_mach,ct FROM loadcap_proc a inner join loadcap_mst_product_r b " _
            & " on a.partno=b.partno inner join loadcap_mst_mach c on a.prod_nomach=c.no_mach where a.partno in (" & PN_to_FD & ")) v1 where " _
            & " tonage_mach=(SELECT min(tonage_mach) FROM loadcap_proc x INNER JOIN loadcap_mst_mach y on x.prod_nomach=y.no_mach " _
            & " where partno=v1.partno)"
            Set rsOVRLD = Con.Execute(qry)
            If rsOVRLD.RecordCount > 0 Then
               
                Dim li As ListItem
                Dim nour As Integer
                Dim neday As Double
                With gridku
                    .rows = 1
                    .rows = 1 + rsOVRLD.RecordCount
                    While Not rsOVRLD.EOF
                        nour = nour + 1
                        c_cap_p_day = ((60 / rsOVRLD("ct")) * rsOVRLD("cavity") * rsOVRLD("hour_p_shift") * rsOVRLD("shift_usg") * 60) * rsOVRLD("faktor_productivity")
                        .TextMatrix(nour, 0) = rsOVRLD("partno")
                        .TextMatrix(nour, 1) = rsOVRLD("prod_nomach")
                        .TextMatrix(nour, 2) = rsOVRLD("tonage_mach")
                        .TextMatrix(nour, 3) = rsOVRLD("cavity")
                        .TextMatrix(nour, 4) = rsOVRLD("hour_p_shift")
                        .TextMatrix(nour, 5) = rsOVRLD("shift_usg")
                        .TextMatrix(nour, 6) = rsOVRLD("faktor_productivity")
                        .TextMatrix(nour, 7) = rsOVRLD("ct")
                        .TextMatrix(nour, 8) = FormatNumber(c_cap_p_day, 0)
                        neday = getSisaPP(Trim(rsOVRLD("partno")), 4) / c_cap_p_day
                        .TextMatrix(nour, 9) = FormatNumber(neday / ovhkw1 * 100)
                        neday = getSisaPP(Trim(rsOVRLD("partno")), 5) / c_cap_p_day
                        .TextMatrix(nour, 10) = FormatNumber(neday / ovhkw2 * 100)
                        neday = getSisaPP(Trim(rsOVRLD("partno")), 6) / c_cap_p_day
                        .TextMatrix(nour, 11) = FormatNumber(neday / ovhkw3 * 100)
                        neday = getSisaPP(Trim(rsOVRLD("partno")), 7) / c_cap_p_day
                        .TextMatrix(nour, 12) = FormatNumber(neday / ovhkw4 * 100)
                       
                        rsOVRLD.MoveNext
                    Wend
                End With
            End If
        End If
        
    End If
End Sub


Private Function isAddedTon(inTon As String) As Boolean
    Dim b As Byte
    Dim idreturn As Boolean
    idreturn = False
    With gridku2
        For b = 1 To .rows - 1
            If .TextMatrix(b, 0) = inTon Then
                idreturn = True
                Exit For
            End If
        Next
    End With
    isAddedTon = idreturn
End Function

Private Sub ckShow_Click()
    Dim u As Byte
    Dim u2 As Byte
    If ckShow.Value = vbChecked Then
        picrecapTonage.Visible = True
        gridku2.rows = 2
        Dim lit As ListItem
        Dim ke As Byte
        ke = 1
        With gridku
            If .rows <= 2 Then Exit Sub
            copyProcessPlan
            For u = 1 To .rows - 1
                 If .TextMatrix(u, 9) * 1 > 0 Then
                    If isAddedTon(.TextMatrix(u, 2)) = False Then
                        ke = 1 + ke
                        gridku2.rows = gridku2.rows + 1
                        gridku2.TextMatrix(ke, 0) = .TextMatrix(u, 2)
                        gridku2.TextMatrix(ke, 2) = .TextMatrix(u, 9)

                    Else
                        For u2 = 1 To gridku2.rows - 1
                            If gridku2.TextMatrix(u2, 0) = .TextMatrix(u, 2) Then
                                gridku2.TextMatrix(u2, 2) = gridku2.TextMatrix(u2, 2) * 1 + .TextMatrix(u, 9) * 1
                            End If
                        Next
                    End If
                 End If
            Next
            
            
            For u = 1 To .rows - 1
                 If .TextMatrix(u, 10) * 1 > 0 Then
                    If isAddedTon(.TextMatrix(u, 2)) = False Then
                        ke = 1 + ke
                        gridku2.rows = gridku2.rows + 1
                        gridku2.TextMatrix(ke, 0) = .TextMatrix(u, 2)
                        gridku2.TextMatrix(ke, 3) = .TextMatrix(u, 10)

                    Else
                        For u2 = 1 To gridku2.rows - 1
                            If gridku2.TextMatrix(u2, 0) = .TextMatrix(u, 2) Then
                                If IsNumeric(gridku2.TextMatrix(u2, 3)) Then
                                    gridku2.TextMatrix(u2, 3) = gridku2.TextMatrix(u2, 3) * 1 + .TextMatrix(u, 10) * 1
                                Else
                                    gridku2.TextMatrix(u2, 3) = .TextMatrix(u, 10) * 1
                                End If
                            End If
                        Next
                    End If
                 End If
            Next
            
            For u = 1 To .rows - 1
                 If .TextMatrix(u, 11) * 1 > 0 Then
                    If isAddedTon(.TextMatrix(u, 2)) = False Then
                        ke = 1 + ke
                        gridku2.rows = gridku2.rows + 1
                        gridku2.TextMatrix(ke, 0) = .TextMatrix(u, 2)
                        gridku2.TextMatrix(ke, 4) = .TextMatrix(u, 11)

                    Else
                        For u2 = 1 To gridku2.rows - 1
                            If gridku2.TextMatrix(u2, 0) = .TextMatrix(u, 2) Then
                                If IsNumeric(gridku2.TextMatrix(u2, 4)) Then
                                    gridku2.TextMatrix(u2, 4) = gridku2.TextMatrix(u2, 4) * 1 + .TextMatrix(u, 11) * 1
                                Else
                                    gridku2.TextMatrix(u2, 4) = .TextMatrix(u, 11) * 1
                                End If
                            End If
                        Next
                    End If
                 End If
            Next
            
            For u = 1 To .rows - 1
                 If .TextMatrix(u, 12) * 1 > 0 Then
                    If isAddedTon(.TextMatrix(u, 2)) = False Then
                        ke = 1 + ke
                        gridku2.rows = gridku2.rows + 1
                        gridku2.TextMatrix(ke, 0) = .TextMatrix(u, 2)
                        gridku2.TextMatrix(ke, 5) = .TextMatrix(u, 12)

                    Else
                        For u2 = 1 To gridku2.rows - 1
                            If gridku2.TextMatrix(u2, 0) = .TextMatrix(u, 2) Then
                                If IsNumeric(gridku2.TextMatrix(u2, 5)) Then
                                    gridku2.TextMatrix(u2, 5) = gridku2.TextMatrix(u2, 5) * 1 + .TextMatrix(u, 12) * 1
                                Else
                                    gridku2.TextMatrix(u2, 5) = .TextMatrix(u, 12) * 1
                                End If
                            End If
                        Next
                    End If
                 End If
            Next
        End With
        
        With gridku3
            For u = 2 To .rows - 1
                For u2 = 2 To .Cols - 5
                    .TextMatrix(u, u2) = .TextMatrix(u, u2) * 1 + getovrVal(.TextMatrix(u, 0), u2)
                Next
            Next
            
            For u = 2 To .rows - 1
                For u2 = 6 To .Cols - 1
                    .TextMatrix(u, u2) = Round(.TextMatrix(u, u2 - 4) / 100)
                Next
            Next
            
            '=======FORMATING
            For u = 2 To .rows - 1
                For u2 = 2 To .Cols - 5
                    .TextMatrix(u, u2) = .TextMatrix(u, u2) & "%"
                Next
            Next
            
            
        End With
        With gridku2
            For u = 2 To .rows - 1
                For u2 = 2 To .Cols - 1
                    .TextMatrix(u, u2) = .TextMatrix(u, u2) & "%"
                Next
            Next
        End With
       
    Else
        picrecapTonage.Visible = False
    End If
End Sub

Function getovrVal(ton As String, kol1 As Byte) As Double
    Dim ce As Byte
    With gridku2
        For ce = 2 To .rows - 1
            If .TextMatrix(ce, 0) = ton Then
                If IsNumeric(.TextMatrix(ce, kol1)) Then
                    getovrVal = .TextMatrix(ce, kol1)
                Else
                    getovrVal = 0
                End If
                Exit For
            End If
        Next
    End With
End Function

Private Function getSisaPP(pn As String, indexmonth As Byte) As Double
    Dim bb As Integer
    With agrid
        For bb = 1 To .rows - 1
            If pn = Trim(.TextMatrix(bb, 1)) Then
                getSisaPP = .TextMatrix(bb, indexmonth) * 1
                Exit For
            End If
        Next
    End With
End Function

Private Sub CmbRevision_DropDown()
    qry = "select distinct on (fltpp_rev) fltpp_rev from loadcap_generate_h where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "' and fltpp_doc='" & CmbDocument & "' and lc_sisa_pp>0"
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
        oSheet.Cells(4, 8) = Label2.Caption
        With oSheet
            
            .Range(.Cells(1, 1), .Cells(5, 14)).Font.Bold = True
            .Columns(1).NumberFormat = "@"
            .Columns(8).NumberFormat = "@"
        End With
        Dim baris As Integer, k As Integer
        baris = 5
        With agrid
            For i = 0 To .rows - 1
                oSheet.Cells(baris, 1) = LTrim(.TextMatrix(i, 1))
                oSheet.Cells(baris, 2) = .TextMatrix(i, 2)
                oSheet.Cells(baris, 3) = .TextMatrix(i, 3)
                If i = 0 Then
                    oSheet.Cells(baris, 4) = DTPicker1.Value
                    oSheet.Cells(baris, 5) = DateAdd("m", 1, DTPicker1.Value)
                    oSheet.Cells(baris, 6) = DateAdd("m", 2, DTPicker1.Value)
                    oSheet.Cells(baris, 7) = DateAdd("m", 3, DTPicker1.Value)
                    For k = 4 To 7
                        oSheet.Cells(baris, k).NumberFormat = "mmm-yy"
                    Next
                Else
                    oSheet.Cells(baris, 4) = .TextMatrix(i, 4)
                    oSheet.Cells(baris, 5) = .TextMatrix(i, 5)
                    oSheet.Cells(baris, 6) = .TextMatrix(i, 6)
                    oSheet.Cells(baris, 7) = .TextMatrix(i, 7)
                End If
                baris = baris + 1
            Next
        End With
        baris = 5
        With angrid2
            For i = 0 To .rows - 1
                oSheet.Cells(baris, 8) = .TextMatrix(i, 1)
                oSheet.Cells(baris, 9) = .TextMatrix(i, 2)
                oSheet.Cells(baris, 10) = .TextMatrix(i, 3)
                If i = 0 Then
                    oSheet.Cells(baris, 11) = DTPicker1.Value
                    oSheet.Cells(baris, 12) = DateAdd("m", 1, DTPicker1.Value)
                    oSheet.Cells(baris, 13) = DateAdd("m", 2, DTPicker1.Value)
                    oSheet.Cells(baris, 14) = DateAdd("m", 3, DTPicker1.Value)
                    For k = 11 To 14
                        oSheet.Cells(baris, k).NumberFormat = "mmm-yy"
                    Next
                Else
                    oSheet.Cells(baris, 11) = .TextMatrix(i, 4)
                    oSheet.Cells(baris, 12) = .TextMatrix(i, 5)
                    oSheet.Cells(baris, 13) = .TextMatrix(i, 6)
                    oSheet.Cells(baris, 14) = .TextMatrix(i, 7)
                End If
                baris = baris + 1
            Next
        End With
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, -4143 'xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub Command1_Click()
    Clipboard.Clear
    With gridku
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = .rows - 1
        Clipboard.SetText .Clip
    End With
    If gridku.rows < 2 Then MsgBox "nothing to be exported": Exit Sub
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
        
        Clipboard.Clear
        With gridku2
            .Col = 0
            .Row = 0
            .ColSel = .Cols - 1
            .RowSel = .rows - 1
            Clipboard.SetText .Clip
        End With
        With oExcel.ActiveWorkbook.ActiveSheet
            .Range("A" & gridku.rows + 5).Select 'Select Cell A1 (will paste from here, to different cells)
            .Paste              'Paste clipboard contents
        End With
        
        Clipboard.Clear
        With gridku3
            .Col = 0
            .Row = 0
            .ColSel = .Cols - 1
            .RowSel = .rows - 1
            Clipboard.SetText .Clip
        End With
        With oExcel.ActiveWorkbook.ActiveSheet
            .Range("A" & gridku2.rows + 5 + gridku.rows + 5).Select 'Select Cell A1 (will paste from here, to different cells)
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

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub settingLV()
    With agrid
        .Cols = 9: .ColWidth(0) = 700: .ColWidth(1) = 2500: .ColWidth(2) = 2500
        .rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "No"
        
        i = 1
        .TextMatrix(0, i) = "Assy No"
        
        i = 2
        .TextMatrix(0, i) = "Assy Name"
        
        i = 3
        .TextMatrix(0, i) = "Unit"
        i = 8
        .TextMatrix(0, i) = "Machine"
        
        .MergeRow(0) = True
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
    With gridku
        .Cols = 13
        .TextMatrix(0, 0) = "Partno"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 3000
        .TextMatrix(0, 1) = "Machine"
        .TextMatrix(0, 2) = "Tonage"
        .TextMatrix(0, 3) = "Cavity"
        .TextMatrix(0, 4) = "Hour Per Shift"
        .TextMatrix(0, 5) = "Total Shift"
        .TextMatrix(0, 6) = "x Faktor (%)"
        .TextMatrix(0, 7) = "CT"
        .TextMatrix(0, 8) = "Cap per Day"
        .TextMatrix(0, 9) = "LC1"
        .TextMatrix(0, 10) = "LC2"
        .TextMatrix(0, 11) = "LC3"
        .TextMatrix(0, 12) = "LC4"
    End With
    
    With gridku2
        .Cols = 6: .ColWidth(0) = 700: .ColWidth(1) = 780: .ColWidth(2) = 900
        .ColWidth(3) = 900: .ColWidth(4) = 900: .ColWidth(5) = 900:
        .rows = 5
        .FixedRows = 2
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 1
            .CellBackColor = RGB(255, 255, 74)
            .Col = i
            .Row = 0
            .CellBackColor = RGB(255, 255, 74)
            .CellAlignment = flexAlignCenterCenter
        Next
        
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "Tonage":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 1
        .TextMatrix(0, i) = "Need Machine":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 0
        
        i = 2
        .TextMatrix(0, i) = "Total overloading"
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
            
        .MergeRow(0) = True
    End With
    
    With gridku3
        .Cols = 10: .ColWidth(0) = 700: .ColWidth(1) = 780: .ColWidth(2) = 900
        .ColWidth(3) = 900: .ColWidth(4) = 900: .ColWidth(5) = 900:
        .rows = 5
        .FixedRows = 2
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 1
            .CellBackColor = RGB(255, 255, 74)
            .Col = i
            .Row = 0
            .CellBackColor = RGB(255, 255, 74)
            .CellAlignment = flexAlignCenterCenter
        Next
        
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "Tonage":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 1
        .TextMatrix(0, i) = "Total Machine":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        
        i = 2
        .TextMatrix(0, i) = "Total Planned + Unplanned"
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        
        i = 6
        .TextMatrix(0, i) = "Need Machine"
        .MergeCol(i) = True
        
        i = 7
        .TextMatrix(0, i) = .TextMatrix(0, 6)
        .MergeCol(i) = True
        
        i = 8
        .TextMatrix(0, i) = .TextMatrix(0, 6)
        .MergeCol(i) = True
        
        i = 9
        .TextMatrix(0, i) = .TextMatrix(0, 6)
        .MergeCol(i) = True
            
        .MergeRow(0) = True
    End With
    
    With lvex
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "MC No"
        .ColumnHeaders.Add , , "Tonage", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "Remark"
        .ColumnHeaders.Add , , "Mach State", 0
    End With
End Sub


Sub copyProcessPlan()
        Dim xx As ListItem
        period1 = Format(DTPicker1.Value, "yyyyMM")
        period2 = Format(DateAdd("m", 1, DTPicker1.Value), "yyyyMM") 'Left(period1, 4) & Right("00" & Val(Right(period1, 2) + 1), 2)
        period3 = Format(DateAdd("m", 2, DTPicker1.Value), "yyyyMM") 'Left(period2, 4) & Right("00" & Val(Right(period2, 2) + 1), 2)
        period4 = Format(DateAdd("m", 3, DTPicker1.Value), "yyyyMM") 'Left(period3, 4) & Right("00" & Val(Right(period3, 2) + 1), 2)
    
        qry = "select b.no_mach,max(tonage_mach) tonage,sum(case when a.fltpp_ym='" & period1 & "' then lcvsmach end) lcvsmach," _
            & "sum(case when a.fltpp_ym='" & period2 & "' then lcvsmach end) lcvsmach2, " _
            & "sum(case when a.fltpp_ym='" & period3 & "' then lcvsmach end) lcvsmach3," _
            & "sum(case when a.fltpp_ym='" & period4 & "' then lcvsmach end) lcvsmach4,max(remark) remark," _
            & "sum(case when a.fltpp_ym='" & period1 & "' then lcneed_mp end) nmp1," _
            & "sum(case when a.fltpp_ym='" & period2 & "' then lcneed_mp end) nmp2," _
            & "sum(case when a.fltpp_ym='" & period3 & "' then lcneed_mp end) nmp3," _
            & "sum(case when a.fltpp_ym='" & period4 & "' then lcneed_mp end) nmp4,coalesce(rstate_mach,state_mach,rstate_mach) stsmsn " _
            & " from loadcap_generate_d a " _
             & " right join loadcap_mst_mach b on a.no_mach=b.no_mach and a.fltpp_rev=" & CmbRevision & "  and a.fltpp_doc='" & CmbDocument & "' " _
            & " left join v_mc_mat c on b.no_mach=c.no_mach " _
            & " " _
            & " group by b.no_mach,rstate_mach,state_mach " _
            & " order by 1"
        Set RsA = Con.Execute(qry)
        lvex.ListItems.Clear
       
        If RsA.RecordCount > 0 Then
            While Not RsA.EOF
                Set xx = lvex.ListItems.Add(, , RsA("no_mach"))
                xx.SubItems(1) = IIf(IsNull(RsA("tonage")), 0, RsA("tonage")) & "T"
                xx.SubItems(2) = IIf(IsNull(RsA("lcvsmach")), 0, RsA("lcvsmach"))
                xx.SubItems(3) = IIf(IsNull(RsA("lcvsmach2")), 0, RsA("lcvsmach2"))
                xx.SubItems(4) = IIf(IsNull(RsA("lcvsmach3")), 0, RsA("lcvsmach3"))
                xx.SubItems(5) = IIf(IsNull(RsA("lcvsmach4")), 0, RsA("lcvsmach4"))
                xx.SubItems(6) = IIf(IsNull(RsA("remark")), "", RsA("remark"))
                xx.SubItems(7) = IIf(IsNull(RsA("stsmsn")), "", RsA("stsmsn"))
                
                
                If RsA("stsmsn") = 1 Then
                    ttlMesin = ttlMesin + 1
                End If
                RsA.MoveNext
            Wend
           
        End If
        qry = "select tonage_mach, COUNT(distinct(b.no_mach)) ttlmesin,sum(case when a.fltpp_ym='" & period1 & "' then lcvsmach end) avglc1, " _
            & " sum(case when a.fltpp_ym='" & period2 & "' then lcvsmach end) avglc2," _
            & " sum(case when a.fltpp_ym='" & period3 & "' then lcvsmach end) avglc3," _
            & " sum(case when a.fltpp_ym='" & period4 & "' then lcvsmach end) avglc4" _
            & " from loadcap_generate_d a " _
            & " right join loadcap_mst_mach b on a.no_mach=b.no_mach and a.fltpp_rev=" & CmbRevision & " and a.fltpp_doc='" & CmbDocument & "' " _
            & " left join v_mc_mat c on b.no_mach=c.no_mach " _
            & " group by tonage_mach " _
            & " order by 1"
        Set RsA = Con.Execute(qry)
        
        If RsA.RecordCount > 0 Then
            i = 2
            gridku3.rows = RsA.RecordCount + i
            While Not RsA.EOF
                With gridku3
                    .TextMatrix(i, 0) = RsA("tonage_mach")
                    .TextMatrix(i, 1) = RsA("ttlmesin")
                    
                    .TextMatrix(i, 2) = FormatNumber(IIf(IsNull(RsA("avglc1")), 0, RsA("avglc1")), 2)
                    .TextMatrix(i, 3) = FormatNumber(IIf(IsNull(RsA("avglc2")), 0, RsA("avglc2")), 2)
                    .TextMatrix(i, 4) = FormatNumber(IIf(IsNull(RsA("avglc3")), 0, RsA("avglc3")), 2)
                    .TextMatrix(i, 5) = FormatNumber(IIf(IsNull(RsA("avglc4")), 0, RsA("avglc4")), 2)
                End With
                i = i + 1
                RsA.MoveNext
            Wend
            
            
            
            Dim k As Byte
            With gridku3
                For i = 2 To gridku3.rows - 1
                   
                   
                    If i Mod 2 = 0 Then
                        For j = 0 To .Cols - 1
                            .Col = j
                            .Row = i
                            .CellBackColor = RGB(255, 255, 149)
                        Next
                    End If
                    .Col = 2
                    .Row = i
                    .CellAlignment = flexAlignRightCenter
                Next
            End With
        End If
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
    cmbFiletype.Top = CmbRevision.Top
    cmbFiletype.Width = cmdExportLC.Width
    cmbFiletype.Left = cmdExportLC.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub Label1_Click()
    MsgBox agrid.rows
End Sub
