VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_ReportofNeedMM 
   Caption         =   "Need for Mold and Machine"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
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
   ScaleHeight     =   7260
   ScaleWidth      =   12330
   Begin VB.OptionButton Option3 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      TabIndex        =   26
      Top             =   1200
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3120
      TabIndex        =   25
      Top             =   1200
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "All"
      Height          =   270
      Left            =   2160
      TabIndex        =   24
      Top             =   1200
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details of Need machine"
      Height          =   1335
      Left            =   6240
      TabIndex        =   16
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton CmdDetails 
         Caption         =   "Details"
         Height          =   855
         Left            =   4440
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbnPeriod 
         Height          =   390
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "F_ReportofNeedMM.frx":0000
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "F_ReportofNeedMM.frx":0060
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel skinHKW 
         Height          =   495
         Left            =   1080
         OleObjectBlob   =   "F_ReportofNeedMM.frx":00C0
         TabIndex        =   20
         Top             =   720
         Width           =   3255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid agrid4 
      Height          =   4815
      Left            =   6240
      TabIndex        =   15
      Top             =   2160
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      _Version        =   393216
      BackColorBkg    =   12648447
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
   Begin MSFlexGridLib.MSFlexGrid agrid3 
      Height          =   4815
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      _Version        =   393216
      BackColorBkg    =   12648447
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
   Begin VB.ListBox List1 
      Height          =   1140
      Left            =   9240
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6240
      ScaleHeight     =   435
      ScaleWidth      =   5955
      TabIndex        =   8
      Top             =   1560
      Width           =   6015
      Begin VB.ComboBox cmbFiletype2 
         Height          =   390
         ItemData        =   "F_ReportofNeedMM.frx":0122
         Left            =   3240
         List            =   "F_ReportofNeedMM.frx":012C
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdNeedMach 
         Caption         =   "Export"
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Need Machine"
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
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   1560
      Width           =   6015
      Begin VB.ComboBox cmbFiletype 
         Height          =   390
         ItemData        =   "F_ReportofNeedMM.frx":0145
         Left            =   3240
         List            =   "F_ReportofNeedMM.frx":014F
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdNeedMold 
         Caption         =   "Export"
         Height          =   375
         Left            =   5160
         TabIndex        =   22
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Need Mold"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.ComboBox CmbRevision 
      Height          =   390
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox CmbDocument 
      Height          =   390
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportofNeedMM.frx":0168
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1275
      TabIndex        =   0
      Top             =   120
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
      Format          =   152764419
      CurrentDate     =   42544
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportofNeedMM.frx":01CE
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportofNeedMM.frx":0230
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   5160
      OleObjectBlob   =   "F_ReportofNeedMM.frx":0296
      Top             =   480
   End
   Begin ACTIVESKINLibCtl.SkinLabel slRow 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportofNeedMM.frx":04CA
      TabIndex        =   10
      Top             =   6960
      Width           =   5895
   End
   Begin ACTIVESKINLibCtl.SkinLabel slRow2 
      Height          =   255
      Left            =   6240
      OleObjectBlob   =   "F_ReportofNeedMM.frx":0522
      TabIndex        =   12
      Top             =   6960
      Width           =   5895
   End
   Begin MSFlexGridLib.MSFlexGrid agrid2 
      Height          =   4815
      Left            =   6240
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      _Version        =   393216
      BackColorBkg    =   12648447
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
Attribute VB_Name = "F_ReportofNeedMM"
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
Private rsB As ADODB.Recordset, rsQ As ADODB.Recordset
Dim qry As String
Dim nmbulan() As String
Dim arrMold() As String
Dim period1 As String
Dim period2 As String
Dim period3 As String
Dim period4 As String
Dim hkw1 As Variant
Dim hkw2 As Variant
Dim hkw3 As Variant
Dim hkw4 As Variant
Private oExcel      As Object
Private oBook       As Object
Private oSheet      As Object
Dim i As Integer, j As Integer
Private totalprodPlan As Variant
Dim totalCapMonth As Variant
Dim ListTonaseToSearch As String
Dim ListPeriodToSearch As String
Dim m_SortColumn As Integer
Dim m_SortOrder As Variant
Dim tempDist As Variant
Dim hkwSelected As Variant
Dim stateFilter As Integer

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

Private Sub SortByColumn(ByVal sort_column As Integer, pgrid As MSFlexGrid)
    ' Hide the FlexGrid.
    pgrid.Visible = False
    pgrid.Refresh

    ' Sort using the clicked column.
    pgrid.Col = sort_column
    pgrid.ColSel = sort_column
    pgrid.Row = 0
    pgrid.RowSel = 0

    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If m_SortColumn <> sort_column Then
        m_SortOrder = flexSortGenericAscending
    ElseIf m_SortOrder = flexSortGenericAscending Then
        m_SortOrder = flexSortGenericDescending
    Else
        m_SortOrder = flexSortGenericAscending
    End If
    pgrid.Sort = m_SortOrder

    ' Restore the previous sort column's name.
    If m_SortColumn >= 0 Then
        pgrid.TextMatrix(0, m_SortColumn) = _
            Mid$(pgrid.TextMatrix(0, m_SortColumn), 3)
    End If

    ' Display the new sort column's name.
    m_SortColumn = sort_column
    If m_SortOrder = flexSortGenericAscending Then
        pgrid.TextMatrix(0, m_SortColumn) = "> " & _
            pgrid.TextMatrix(0, m_SortColumn)
    Else
        pgrid.TextMatrix(0, m_SortColumn) = "< " & _
            pgrid.TextMatrix(0, m_SortColumn)
    End If

    ' Display the FlexGrid.
    pgrid.Visible = True
End Sub

Private Sub agrid2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If agrid2.MouseRow <> 0 Then Exit Sub
    SortByColumn agrid2.MouseCol, agrid2
End Sub

Private Sub reNeedMach()
    If Len(CmbDocument) > 0 And Len(CmbRevision) > 0 Then
        Screen.MousePointer = 11
        qry = "select tonase,sum(neday1) neday1,sum(neday2) neday2,sum(neday3) neday3,sum(neday4) neday4,max(hkw1) hkw1," _
        & " max(hkw2) hkw2,max(hkw3) hkw3,max(hkw4) hkw4 from " _
        & " (select tonase,sum(case when fltpp_ym ='" & period1 & "' then neday else 0 end) neday1,sum(case when fltpp_ym ='" & period2 & "' then neday else 0 end) neday2," _
        & " sum(case when fltpp_ym ='" & period3 & "' then neday else 0 end) neday3,sum(case when fltpp_ym ='" & period4 & "' then neday else 0 end) neday4," _
        & " max(case when fltpp_ym ='" & period1 & "' then hkw else 0 end) hkw1,max(case when fltpp_ym ='" & period2 & "' then hkw else 0 end) hkw2," _
        & " max(case when fltpp_ym ='" & period3 & "' then hkw else 0 end) hkw3,max(case when fltpp_ym ='" & period4 & "' then hkw else 0 end) hkw4," _
        & " fltpp_ym From " _
        & " (select lc_itemid,max(lc_itemname) itemname,max(lc_sisa_pp) sisa_pp,max(fltpp_hkw) hkw,max(lc_fprodtvty) prodtvty,max(no_mach) mach,min(ton_mach) tonase,max(cav) cav,max(ct) ct," _
        & " max(cap_p_day) cap_p_day,(case when max(cap_p_day)>0 then max(lc_sisa_pp)/max(cap_p_day) else 0 end) neday,every(rstate_mach) rstate_mach, fltpp_ym from " _
             & " (select lc_itemid,lc_itemname,lc_sisa_pp,fltpp_hkw,lc_fprodtvty,fltpp_ym from loadcap_generate_h " _
             & " where  fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & ") AS v1 " _
             & " Inner Join " _
             & " (select lcd_itemdid,no_mach,ton_mach,cav,ct,cap_p_day,rstate_mach from loadcap_generate_d " _
             & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " order by ton_mach asc) as v2 on v1.lc_itemid=v2.lcd_itemdid " _
             & " group by lc_itemid,fltpp_ym) as vv1 " _
        & " group by tonase,fltpp_ym) as vvv1 " _
        & " group by tonase " _
        & " order by tonase asc"
        Set RsA = Con.Execute(qry)
        agrid4.rows = 2
        If RsA.RecordCount > 0 Then
            i = 2
            hkw1 = RsA("hkw1")
            hkw2 = RsA("hkw2")
            hkw3 = RsA("hkw3")
            hkw4 = RsA("hkw4")
            agrid4.rows = RsA.RecordCount + 2
            While Not RsA.EOF
                With agrid4
                    .TextMatrix(i, 0) = RsA("tonase")
                    .TextMatrix(i, 1) = FormatNumber(RsA("neday1"), 2)
                    .TextMatrix(i, 2) = FormatNumber(RsA("neday2"), 2)
                    .TextMatrix(i, 3) = FormatNumber(RsA("neday3"), 2)
                    .TextMatrix(i, 4) = FormatNumber(RsA("neday4"), 2)
                End With
                i = i + 1
                RsA.MoveNext
            Wend
            '-----------------DAPATKAN SISA HARI MESIN----------------------
            getTotalBal period1, 5, hkw1
            getTotalBal period2, 6, hkw2
            getTotalBal period3, 7, hkw3
            getTotalBal period4, 8, hkw4
'            '--------------DAPATKAN NEED MACHINE
'            For i = 1 To agrid2.rows - 1
'                With agrid2
'                    .TextMatrix(i, 4) = .TextMatrix(i, 2) * 1 - .TextMatrix(i, 3) * 1
'                    .TextMatrix(i, 5) = Application.WorksheetFunction.RoundUp(.TextMatrix(i, 4) / -hkwSelected, 0)
'
'                End With
'            Next
'            '---------------DAPATKAN MESIN OFF
'            If Len(ListTonaseToSearch) > 0 Then
'                ListTonaseToSearch = Left(ListTonaseToSearch, Len(ListTonaseToSearch) - 1)
''                MsgBox ListTonaseToSearch
'                qry = "select ton_mach,count(case rstate_mach when FALSE then 1 end) from" _
'                    & "(SELECT distinct on(no_mach) no_mach,ton_mach,rstate_mach from loadcap_generate_d " _
'                    & "where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "' and ton_mach in (" & ListTonaseToSearch & ")) v1 " _
'                    & "group by ton_mach"
'                Set rsA = Con.Execute(qry)
'                If rsA.RecordCount > 0 Then
'                    While Not rsA.EOF
'                        With agrid2
'                        For i = 1 To .rows - 1
'                            If rsA("ton_mach") = .TextMatrix(i, 0) Then
'                                .TextMatrix(i, 1) = rsA(1)
'                            End If
'                        Next
'                        End With
'                        rsA.MoveNext
'                    Wend
'                End If
'            End If
'            '----------FINISHING: FORMAT NUMBER------------
'            For j = 1 To agrid2.rows - 1
'                With agrid2
'                    .TextMatrix(j, 3) = FormatNumber(.TextMatrix(j, 3), 2)
'                    .TextMatrix(j, 4) = FormatNumber(.TextMatrix(j, 4), 2)
'                End With
'            Next
        End If
        Screen.MousePointer = 0
    End If
End Sub

Private Sub getTotalBal(pperiod As String, pINDEX As Integer, phkw As Variant)
    'ListTonaseToSearch = ""
    For i = 2 To agrid4.rows - 1
        With agrid4
            qry = "select ton_mach,sum(neday) neday,count(distinct no_mach) ttlmach from loadcap_generate_d " _
            & " where fltpp_doc='" & CmbDocument.Text & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & pperiod & "' and ton_mach=" & .TextMatrix(i, 0) _
            & " group by ton_mach"
            Set RsA = Con.Execute(qry)
            While Not RsA.EOF
                .TextMatrix(i, pINDEX) = RsA("ttlmach") * phkw - RsA("neday")
                If Left(.TextMatrix(i, pINDEX), 1) = "-" Then
                    .TextMatrix(i, pINDEX) = 0
                End If
                .TextMatrix(i, pINDEX + 4) = .TextMatrix(i, pINDEX) * 1 - .TextMatrix(i, pINDEX - 4) * 1
                .TextMatrix(i, pINDEX + 4 + 4) = Round(.TextMatrix(i, pINDEX + 4) / phkw, 0) 'Application.WorksheetFunction.RoundUp(.TextMatrix(i, pINDEX + 4) / phkw, 0)
                RsA.MoveNext
            Wend
            'ListTonaseToSearch = ListTonaseToSearch & .TextMatrix(i, 0) & ","
        End With
    Next
End Sub

Private Sub cmbnPeriod_Click()
    If Len(CmbDocument) > 0 And Len(CmbRevision) > 0 And Len(cmbnPeriod) > 0 Then

        qry = "select lc_itemid,lc_itemname,lc_pp,fltpp_hkw,lc_subcont " _
        & " from loadcap_generate_h " _
        & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "'" _
        & " order by lc_itemid asc"
        Set RsA = Con.Execute(qry)

        If RsA.RecordCount > 0 Then
            hkwSelected = RsA("fltpp_hkw")
            skinHKW.Caption = "HKW : " & hkwSelected

        End If
''        loadDetil_v2
'
        '-------------------==============------------------------
        qry = "select lc_itemid,max(lc_itemname) itemname,max(lc_sisa_pp) sisa_pp,max(fltpp_hkw) hkw,max(lc_fprodtvty) prodtvty,max(no_mach) mach,min(ton_mach) tonase,max(cav) cav,max(ct) ct, " _
            & " max(cap_p_day) cap_p_day,(case when max(cap_p_day)>0 then max(lc_sisa_pp)/max(cap_p_day) else 0 end) neday,every(rstate_mach) rstate_mach from " _
            & " (select lc_itemid,lc_itemname,lc_sisa_pp,fltpp_hkw,lc_fprodtvty from loadcap_generate_h " _
            & " where  fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "') AS v1 " _
            & " Inner Join " _
            & " (select lcd_itemdid,no_mach,ton_mach,cav,ct,cap_p_day,rstate_mach from loadcap_generate_d " _
            & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "' order by ton_mach asc) as v2 on v1.lc_itemid=v2.lcd_itemdid " _
            & " group by lc_itemid"

        Set RsA = Con.Execute(qry)
        agrid2.rows = 1
        If RsA.RecordCount > 0 Then
            i = 1
            agrid2.rows = 2
            While Not RsA.EOF
                With agrid2
                    tempDist = myDistinctWay(RsA("tonase"))
                    If tempDist > 0 Then
'                        .TextMatrix(tempDist, 0) = rsA("tonase")
'                        If IsNumeric(.TextMatrix(tempDist, 3)) = False Then
                        .TextMatrix(tempDist, 3) = RsA("neday") + .TextMatrix(tempDist, 3) * 1
'                        Else
'                            .TextMatrix(i, 3) = rsA("neday") + .TextMatrix(i, 3) * 1
'                        End If
                    Else
                        .TextMatrix(i, 0) = RsA("tonase")
                        If IsNumeric(.TextMatrix(i, 3)) = False Then
                            .TextMatrix(i, 3) = RsA("neday")
                        Else
                            .TextMatrix(i, 3) = RsA("neday") + .TextMatrix(i, 3) * 1
                        End If
                        i = i + 1
                        agrid2.rows = agrid2.rows + 1
                    End If
                End With
                RsA.MoveNext
            Wend
            agrid2.rows = agrid2.rows - 1
'            slRow2.Caption = agrid2.rows - 1 & " row(s) found"
'
'            '-----------------DAPATKAN SISA HARI MESIN----------------------
            ListTonaseToSearch = ""
            For i = 1 To agrid2.rows - 1
                With agrid2
                    qry = "select ton_mach,sum(neday) neday,count(distinct no_mach) ttlmach from loadcap_generate_d " _
                    & " where fltpp_doc='" & CmbDocument.Text & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "' and ton_mach=" & .TextMatrix(i, 0) _
                    & " group by ton_mach"
                    Set RsA = Con.Execute(qry)
                    While Not RsA.EOF
                        .TextMatrix(i, 2) = RsA("ttlmach") * hkwSelected - RsA("neday")
                        If Left(.TextMatrix(i, 2), 1) = "-" Then
                            .TextMatrix(i, 2) = 0
                        End If
                        RsA.MoveNext
                    Wend
                    ListTonaseToSearch = ListTonaseToSearch & .TextMatrix(i, 0) & ","
                End With
            Next
'            '--------------DAPATKAN NEED MACHINE
            For i = 1 To agrid2.rows - 1
                With agrid2
                    .TextMatrix(i, 4) = .TextMatrix(i, 2) * 1 - .TextMatrix(i, 3) * 1
                    .TextMatrix(i, 5) = Round(.TextMatrix(i, 4) / -hkwSelected, 0) 'Application.WorksheetFunction.RoundUp(.TextMatrix(i, 4) / -hkwSelected, 0)

                End With
            Next
'            '---------------DAPATKAN MESIN OFF
'            If Len(ListTonaseToSearch) > 0 Then
'                ListTonaseToSearch = Left(ListTonaseToSearch, Len(ListTonaseToSearch) - 1)
''                MsgBox ListTonaseToSearch
'                qry = "select ton_mach,count(case rstate_mach when FALSE then 1 end) from" _
'                    & "(SELECT distinct on(no_mach) no_mach,ton_mach,rstate_mach from loadcap_generate_d " _
'                    & "where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "' and ton_mach in (" & ListTonaseToSearch & ")) v1 " _
'                    & "group by ton_mach"
'                Set rsA = Con.Execute(qry)
'                If rsA.RecordCount > 0 Then
'                    While Not rsA.EOF
'                        With agrid2
'                        For i = 1 To .rows - 1
'                            If rsA("ton_mach") = .TextMatrix(i, 0) Then
'                                .TextMatrix(i, 1) = rsA(1)
'                            End If
'                        Next
'                        End With
'                        rsA.MoveNext
'                    Wend
'                End If
'            End If
'            '----------FINISHING: FORMAT NUMBER------------
'            For j = 1 To agrid2.rows - 1
'                With agrid2
'                    .TextMatrix(j, 3) = FormatNumber(.TextMatrix(j, 3), 2)
'                    .TextMatrix(j, 4) = FormatNumber(.TextMatrix(j, 4), 2)
'                End With
'            Next
        End If
        Screen.MousePointer = 0
    End If
End Sub

Private Function myDistinctWay(ptonas As String) As Variant
    For j = 1 To agrid2.rows - 1
        With agrid2
            If ptonas = .TextMatrix(j, 0) Then
                myDistinctWay = j
                Exit For
            Else
                myDistinctWay = 0
            End If
        End With
    Next
End Function

Private Sub cmbnPeriod_DropDown()
    If Len(CmbDocument) > 0 And Len(CmbRevision) > 0 Then
        qry = "select distinct on (fltpp_ym) fltpp_ym from loadcap_generate_h where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "' and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision
        Set RsA = Con.Execute(qry)
        cmbnPeriod.Clear
        If RsA.RecordCount > 0 Then
            While Not RsA.EOF
                cmbnPeriod.AddItem RsA(0)
                RsA.MoveNext
            Wend
        End If
    Else
        cmbnPeriod.Clear
    End If
End Sub

Private Sub CmbRevision_Click()
On Error GoTo Ex
    If CmbDocument.Text <> "" Then
        Screen.MousePointer = 11
        Dim whereS As String
        If stateFilter = 1 Then
            whereS = " where (scapbln1-bln1)<0"
        ElseIf stateFilter = 2 Then
            whereS = " where (scapbln1-bln1)>0"
        Else
            whereS = ""
        End If
'        qry = "select lc_itemid,coalesce(sum(bln1),0) bln1,coalesce(sum(bln2),0) bln2,coalesce(sum(bln3),0) bln3,coalesce(sum(bln4),0) bln4,max(subcont) subcont from " _
'        & " (select lc_itemid, " _
'        & " (case when fltpp_ym='" & period1 & "' then sum(lc_pp) end) bln1, " _
'        & " (case when fltpp_ym='" & period2 & "' then sum(lc_pp) end) bln2, " _
'        & " (case when fltpp_ym='" & period3 & "' then sum(lc_pp) end) bln3, " _
'        & " (case when fltpp_ym='" & period4 & "' then sum(lc_pp) end) bln4,max(lc_subcont) subcont " _
'        & " From loadcap_generate_h " _
'        & " where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision _
'        & " group by lc_itemid,fltpp_ym " _
'        & " order by lc_itemid asc ) as v1 " _
'        & " group by lc_itemid"
'        qry = "select lcd_itemdid,reg_mold, cav, scpday, tfin.subcont, cavity_std, cav_std," _
'        & " capbln1,capbln2, capbln3, capbln4,bln1,bln2,bln3,bln4   from " _
'    & " (select lcd_itemdid,reg_mold, max(cav) cav,max(scpday) scpday,max(subcont) subcont,max(cavity_std) cavity_std, max(cav_std) cav_std, " _
'    & " coalesce(sum(capbln1),0) capbln1, coalesce(sum(capbln2),0) capbln2, " _
'    & " coalesce(sum(capbln3),0) capbln3,coalesce(sum(capbln4),0) capbln4 from " _
'    & " (select lcd_itemdid,reg_mold,max(cav) cav,max(cap_p_day) scpday,subcont,cavity_std,max(a.cav_std) cav_std,max(fltpp_hkw) fltpp_hkw,a.fltpp_ym, " _
'    & " (case when a.fltpp_ym='" & period1 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln1, " _
'    & " (case when a.fltpp_ym='" & period2 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln2, " _
'& " (case when a.fltpp_ym='" & period3 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln3, " _
'& " (case when a.fltpp_ym='" & period3 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln4 " _
'& " from loadcap_generate_d a " _
'& " inner join loadcap_proc b on a.lcd_itemdid=b.partno and a.no_mach=b.prod_nomach and a.reg_mold=b.mold_no " _
'& " inner join loadcap_generate_h c on a.lcd_itemdid=c.lc_itemid and a.fltpp_doc=c.fltpp_doc and a.fltpp_ym=c.fltpp_ym and a.fltpp_rev=c.fltpp_rev " _
'& " where a.fltpp_doc='" & CmbDocument & "' and a.fltpp_rev=" & CmbRevision _
'& " group by a.fltpp_ym,lcd_itemdid,reg_mold,cavity_std,subcont " _
'& " order by lcd_itemdid asc, a.fltpp_ym asc) tlast  group by lcd_itemdid, reg_mold " _
'& " order by lcd_itemdid asc) tfin inner join (select lc_itemid,coalesce(sum(bln1),0) bln1,coalesce(sum(bln2),0) bln2,coalesce(sum(bln3),0) bln3,coalesce(sum(bln4),0) bln4,max(subcont) subcont from " _
'& "         (select lc_itemid,  (case when fltpp_ym='" & period1 & "' then sum(lc_pp) end) bln1, " _
'         & " (case when fltpp_ym='" & period2 & "' then sum(lc_pp) end) bln2, " _
'         & " (case when fltpp_ym='" & period3 & "' then sum(lc_pp) end) bln3, " _
'         & " (case when fltpp_ym='" & period4 & "' then sum(lc_pp) end) bln4,max(lc_subcont) subcont " _
'         & " From loadcap_generate_h  where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision _
'         & " group by lc_itemid,fltpp_ym order by lc_itemid asc ) as v1 " _
'         & " group by lc_itemid) tfin2 on tfin.lcd_itemdid=tfin2.lc_itemid"
    
    qry = "select tfixa.lcd_itemdid,reg_mold, cav, scpday, subcont, coalesce(cavity_std,cav) cavity_std, cav_std,capbln1,capbln2, capbln3, capbln4,bln1,bln2,bln3,bln4, scapbln1, scapbln2,  scapbln3,  scapbln4 from (select lcd_itemdid,reg_mold, cav, scpday, tfin.subcont, cavity_std, cav_std,capbln1,capbln2, capbln3, capbln4,bln1,bln2,bln3,bln4  from " _
    & " (select lcd_itemdid,reg_mold, max(cav) cav,max(scpday) scpday,max(subcont) subcont,max(cavity_std) cavity_std, max(cav_std) cav_std,coalesce(sum(capbln1),0) capbln1,coalesce(sum(capbln2),0) capbln2,coalesce(sum(capbln3),0) capbln3,coalesce(sum(capbln4),0) capbln4 from " _
    & " (select lcd_itemdid,reg_mold,max(cav) cav,max(cap_p_day) scpday,subcont,cavity_std,max(a.cav_std) cav_std,max(fltpp_hkw) fltpp_hkw,a.fltpp_ym, (case when a.fltpp_ym='" & period1 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln1,(case when a.fltpp_ym='" & period2 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln2, " _
    & " (case when a.fltpp_ym='" & period3 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln3, (case when a.fltpp_ym='" & period4 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln4 " _
    & " from loadcap_generate_d a inner join loadcap_proc b on a.lcd_itemdid=b.partno and a.no_mach=b.prod_nomach and a.reg_mold=b.mold_no inner join loadcap_generate_h c on a.lcd_itemdid=c.lc_itemid and a.fltpp_doc=c.fltpp_doc and a.fltpp_ym=c.fltpp_ym and a.fltpp_rev=c.fltpp_rev " _
    & " where a.fltpp_doc='" & CmbDocument & "' and a.fltpp_rev=" & CmbRevision & " group by a.fltpp_ym,lcd_itemdid,reg_mold,cavity_std,subcont order by lcd_itemdid asc, a.fltpp_ym asc) tlast " _
    & " group by lcd_itemdid, reg_mold order by lcd_itemdid asc) tfin inner join (select lc_itemid,coalesce(sum(bln1),0) bln1,coalesce(sum(bln2),0) bln2,coalesce(sum(bln3),0) bln3,coalesce(sum(bln4),0) bln4,max(subcont) subcont from " _
         & " (select lc_itemid, (case when fltpp_ym='" & period1 & "' then sum(lc_pp) end) bln1,  (case when fltpp_ym='" & period2 & "' then sum(lc_pp) end) bln2, (case when fltpp_ym='" & period3 & "' then sum(lc_pp) end) bln3, " _
         & " (case when fltpp_ym='" & period4 & "' then sum(lc_pp) end) bln4,max(lc_subcont) subcont  From loadcap_generate_h where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision _
         & " group by lc_itemid,fltpp_ym order by lc_itemid asc ) as v1 group by lc_itemid) tfin2 on tfin.lcd_itemdid=tfin2.lc_itemid) tfixa " _
        & " inner join ( select lcd_itemdid, sum(capbln1) scapbln1,sum(capbln2) scapbln2, sum(capbln3) scapbln3, sum(capbln4) scapbln4 from " _
    & " (select lcd_itemdid, capbln1,capbln2, capbln3, capbln4  from (select lcd_itemdid,reg_mold,coalesce(sum(capbln1),0) capbln1,coalesce(sum(capbln2),0) capbln2,coalesce(sum(capbln3),0) capbln3,coalesce(sum(capbln4),0) capbln4 from " _
    & " (select lcd_itemdid,reg_mold,max(cav) cav,max(cap_p_day) scpday,subcont,cavity_std,max(a.cav_std) cav_std,max(fltpp_hkw) fltpp_hkw,a.fltpp_ym, " _
    & " (case when a.fltpp_ym='" & period1 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln1, " _
    & " (case when a.fltpp_ym='" & period2 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln2, " _
    & " (case when a.fltpp_ym='" & period3 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln3, " _
    & " (case when a.fltpp_ym='" & period4 & "' then max(cap_p_day)*max(fltpp_hkw) end) capbln4 " _
    & " from loadcap_generate_d a inner join loadcap_proc b on a.lcd_itemdid=b.partno and a.no_mach=b.prod_nomach and a.reg_mold=b.mold_no " _
    & " inner join loadcap_generate_h c on a.lcd_itemdid=c.lc_itemid and a.fltpp_doc=c.fltpp_doc and a.fltpp_ym=c.fltpp_ym and a.fltpp_rev=c.fltpp_rev " _
    & " where a.fltpp_doc='" & CmbDocument & "' and a.fltpp_rev=" & CmbRevision & " group by a.fltpp_ym,lcd_itemdid,reg_mold,cavity_std,subcont " _
    & " order by lcd_itemdid asc, a.fltpp_ym asc) tlast group by lcd_itemdid, reg_mold " _
    & " order by lcd_itemdid asc) tfin inner join (select lc_itemid,coalesce(sum(bln1),0) bln1,coalesce(sum(bln2),0) bln2,coalesce(sum(bln3),0) bln3,coalesce(sum(bln4),0) bln4,max(subcont) subcont from " _
         & " (select lc_itemid, (case when fltpp_ym='" & period1 & "' then sum(lc_pp) end) bln1, (case when fltpp_ym='" & period2 & "' then sum(lc_pp) end) bln2, (case when fltpp_ym='" & period3 & "' then sum(lc_pp) end) bln3, " _
         & " (case when fltpp_ym='" & period4 & "' then sum(lc_pp) end) bln4,max(lc_subcont) subcont From loadcap_generate_h where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision _
         & " group by lc_itemid,fltpp_ym  order by lc_itemid asc ) as v1  group by lc_itemid) tfin2 on tfin.lcd_itemdid=tfin2.lc_itemid) tfix group by lcd_itemdid ) tfixb on tfixa.lcd_itemdid=tfixb.lcd_itemdid " & whereS
'        Clipboard.Clear
'        Clipboard.SetText qry
        Set RsA = Con.Execute(qry)
        agrid3.rows = 2
        If RsA.RecordCount > 0 Then
            With agrid3
                .rows = 2
                i = .rows
                .rows = 3
                While Not RsA.EOF
                    If (RsA("scapbln2") - RsA("bln2")) < 0 Then
                        .Row = i
                        .Col = 0
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                        .Col = 15
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                    End If
                    If (RsA("scapbln1") - RsA("bln1")) < 0 Then
                        .Row = i
                        .Col = 0
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                        .Col = 14
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                    End If
                    If (RsA("scapbln3") - RsA("bln3")) < 0 Then
                        .Row = i
                        .Col = 0
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                        .Col = 16
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                    End If
                    If (RsA("scapbln4") - RsA("bln4")) < 0 Then
                        .Row = i
                        .Col = 0
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                        .Col = 17
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                    End If
                    .TextMatrix(i, 0) = RsA("lcd_itemdid")
                    .TextMatrix(i, 6) = FormatNumber(RsA("capbln1"), 0)
                    .TextMatrix(i, 7) = FormatNumber(RsA("capbln2"), 0)
                    .TextMatrix(i, 8) = FormatNumber(RsA("capbln3"), 0)
                    .TextMatrix(i, 9) = FormatNumber(RsA("capbln4"), 0)
                    .TextMatrix(i, 10) = FormatNumber(RsA("bln1"), 0)
                    .TextMatrix(i, 11) = FormatNumber(RsA("bln2"), 0)
                    .TextMatrix(i, 12) = FormatNumber(RsA("bln3"), 0)
                    .TextMatrix(i, 13) = FormatNumber(RsA("bln4"), 0)
                    .TextMatrix(i, 14) = FormatNumber(RsA("scapbln1") - RsA("bln1"), 0)
                    .TextMatrix(i, 15) = FormatNumber(RsA("scapbln2") - RsA("bln2"), 0)
                    .TextMatrix(i, 16) = FormatNumber(RsA("scapbln3") - RsA("bln3"), 0)
                    .TextMatrix(i, 17) = FormatNumber(RsA("scapbln4") - RsA("bln4"), 0)
                    .TextMatrix(i, 1) = RsA("reg_mold")
                    .TextMatrix(i, 3) = RsA("cavity_std")
                    .TextMatrix(i, 4) = RsA("cav")
                    .TextMatrix(i, 5) = FormatNumber(RsA("cav") / RsA("cavity_std") * 100, 2)
                    .TextMatrix(i, 2) = RsA("subcont")
                    .Col = 0
                    .Row = i
                    .CellAlignment = flexAlignLeftCenter
                    i = 1 + i
                    .rows = 1 + .rows
                    RsA.MoveNext
                Wend
                .rows = .rows - 1
            End With
        End If

'        P4bulan period1, 2, 10, 6
'        P4bulan period2, 3, 11, 7
'        P4bulan period3, 4, 12, 8
'        P4bulan period4, 5, 13, 9
        Screen.MousePointer = 0
    End If
    reNeedMach
'    If CmbDocument.Text <> "" Then
'        Screen.MousePointer = 11
'        qry = "select lc_itemid,coalesce(sum(bln1),0) bln1,coalesce(sum(bln2),0) bln2,coalesce(sum(bln3),0) bln3,coalesce(sum(bln4),0) bln4,max(subcont) subcont from " _
'        & " (select lc_itemid, " _
'        & " (case when fltpp_ym='" & period1 & "' then sum(lc_pp) end) bln1, " _
'        & " (case when fltpp_ym='" & period2 & "' then sum(lc_pp) end) bln2, " _
'        & " (case when fltpp_ym='" & period3 & "' then sum(lc_pp) end) bln3, " _
'        & " (case when fltpp_ym='" & period4 & "' then sum(lc_pp) end) bln4,max(lc_subcont) subcont " _
'        & " From loadcap_generate_h " _
'        & " where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision _
'        & " group by lc_itemid,fltpp_ym " _
'        & " order by lc_itemid asc ) as v1 " _
'        & " group by lc_itemid"
'        Clipboard.Clear
'        Clipboard.SetText qry
'        Set rsA = Con.Execute(qry)
'        If rsA.RecordCount > 0 Then
'            With agrid3
'                .rows = 2
'                i = .rows
'                .rows = 3
'                While Not rsA.EOF
'                    .TextMatrix(i, 0) = rsA("lc_itemid")
'                    .TextMatrix(i, 6) = FormatNumber(rsA("bln1"), 0)
'                    .TextMatrix(i, 7) = FormatNumber(rsA("bln2"), 0)
'                    .TextMatrix(i, 8) = FormatNumber(rsA("bln3"), 0)
'                    .TextMatrix(i, 9) = FormatNumber(rsA("bln4"), 0)
'                    .TextMatrix(i, 1) = rsA("subcont")
'                    .Col = 0
'                    .Row = i
'                    .CellAlignment = flexAlignLeftCenter
'                    i = 1 + i
'                    .rows = 1 + .rows
'                    rsA.MoveNext
'                Wend
'                .rows = .rows - 1
'            End With
'        End If
'
'        P4bulan period1, 2, 10, 6
'        P4bulan period2, 3, 11, 7
'        P4bulan period3, 4, 12, 8
'        P4bulan period4, 5, 13, 9
'        Screen.MousePointer = 0
'    End If
'    reNeedMach
Exit Sub
Ex:
    MsgBox Err.Description, vbCritical, "Sorry"
    Screen.MousePointer = 0
End Sub

Private Sub P4bulan(pperiod As String, pkol As Integer, pkolNM As Integer, pkolPp As Integer)
    Dim b As Integer, B2 As Integer
    With agrid3
        If .rows > 1 Then
            Me.Refresh
            .Refresh
            For B2 = 2 To .rows - 1
                qry = "select * from " _
                    & " (select lc_itemid,lc_pp,lc_fprodtvty,fltpp_hkw " _
                    & " From loadcap_generate_h " _
                    & " where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & "  and fltpp_ym='" & pperiod & "' and lc_itemid='" & .TextMatrix(B2, 0) & "' " _
                    & " order by lc_itemid asc) x " _
                    & " Inner Join " _
                    & " (select lcd_itemdid,cav,ct,cap_p_day,reg_mold,no_mach " _
                    & " From loadcap_generate_d " _
                    & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & pperiod & "' " _
                    & " order by lcd_itemdid asc) y on x.lc_itemid=y.lcd_itemdid " _
                    & " order by cap_p_day asc"
                 Set rsQ = Con.Execute(qry)
'                 MsgBox qry
                 List1.Clear
                 ReDim arrMold(1 To 1)
                 b = 1
'                 MsgBox rsQ.RecordCount, vbInformation, "wooow"
                If rsQ.RecordCount > 0 Then
                    While Not rsQ.EOF
                        If checkArrUni(rsQ("reg_mold")) = False Then
                            If UBound(arrMold) = 1 And arrMold(1) = "" Then
                                arrMold(UBound(arrMold)) = rsQ("reg_mold")
                            Else
                                ReDim Preserve arrMold(1 To UBound(arrMold) + 1)
                                arrMold(UBound(arrMold)) = rsQ("reg_mold")
                            End If
                        End If
                        List1.AddItem rsQ("cap_p_day") * rsQ("fltpp_hkw")
                        rsQ.MoveNext
                    Wend
                Else
                    List1.AddItem 0
                End If
                totalCapMonth = 0
                For b = 1 To UBound(arrMold)
                    totalCapMonth = totalCapMonth * 1 + List1.List(b - 1) * 1
                Next
                .TextMatrix(B2, pkol) = FormatNumber(totalCapMonth, 0)
                .TextMatrix(B2, pkolNM) = totalCapMonth - .TextMatrix(B2, pkolPp) * 1
                .TextMatrix(B2, pkolNM) = FormatNumber(.TextMatrix(B2, pkolNM), 0)
                If Left(.TextMatrix(B2, pkolNM), 1) = "-" Then
'                    For j = 0 To .Cols - 1
                        .Row = B2
                        .Col = 0
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
                        .Col = pkolNM
                        .CellBackColor = RGB(255, 0, 0)
                        .CellForeColor = RGB(255, 255, 255)
'                    Next
                End If
            Next
        End If
    End With
End Sub

Private Sub CmdDetails_Click()
    If agrid2.rows > 1 And agrid2.TextMatrix(1, 0) <> "" Then
        PopUp_DetailsNM.Show
    End If
End Sub

Private Sub cmdNeedMach_Click()
    If agrid3.rows < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Dim spreasheet     As String
        If cmbFiletype.ListIndex = 0 Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        Dim k As Integer
        oExcel.DisplayAlerts = False
        With agrid4
            For i = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If i = 1 Then
                        If k + 1 > 13 Then
                            If k = 1 Or k = 13 Then
                                oSheet.Cells(i + 1, (k + 1) - 8) = DTPicker1.Value
                            ElseIf k = 2 Or k = 14 Then
                                oSheet.Cells(i + 1, (k + 1) - 8) = DateAdd("m", 1, DTPicker1.Value)
                            ElseIf k = 3 Or k = 15 Then
                                oSheet.Cells(i + 1, (k + 1) - 8) = DateAdd("m", 2, DTPicker1.Value)
                            ElseIf k = 4 Or k = 16 Then
                                oSheet.Cells(i + 1, (k + 1) - 8) = DateAdd("m", 3, DTPicker1.Value)
                            End If
                            oSheet.Cells(i + 1, (k + 1) - 8).NumberFormat = "mmm-yy"
                        Else
                            If k = 1 Or k = 13 Then
                                oSheet.Cells(i + 1, k + 1) = DTPicker1.Value
                            ElseIf k = 2 Or k = 14 Then
                                oSheet.Cells(i + 1, k + 1) = DateAdd("m", 1, DTPicker1.Value)
                            ElseIf k = 3 Or k = 15 Then
                                oSheet.Cells(i + 1, k + 1) = DateAdd("m", 2, DTPicker1.Value)
                            ElseIf k = 4 Or k = 16 Then
                                oSheet.Cells(i + 1, k + 1) = DateAdd("m", 3, DTPicker1.Value)
                            End If
                            oSheet.Cells(i + 1, k + 1).NumberFormat = "mmm-yy"
                        End If
                    Else
                        If k + 1 > 5 And k + 1 < 14 Then
                            
                        Else
                            If k + 1 > 13 Then
                                oSheet.Cells(i + 1, (k + 1) - 8) = .TextMatrix(i, k)
                            Else
                                oSheet.Cells(i + 1, k + 1) = .TextMatrix(i, k)
                            End If
                        End If
                    End If
                Next
            Next
            With oSheet
                .Range(.Cells(1, 2), .Cells(1, 5)).Merge
                .Cells(1, 2).HorizontalAlignment = xlCenter
                .Range(.Cells(1, 6), .Cells(1, 9)).Merge
                .Cells(1, 6).HorizontalAlignment = xlCenter
            End With
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

Private Sub cmdNeedMold_Click()
    If agrid3.rows < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Dim spreasheet     As String
        If cmbFiletype.ListIndex = 0 Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        Dim k As Integer
        With agrid3
            For i = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If i = 1 Then
                        If k = 2 Or k = 6 Or k = 10 Then
                            oSheet.Cells(i + 1, k + 1) = DTPicker1.Value
                        ElseIf k = 3 Or k = 7 Or k = 11 Then
                            oSheet.Cells(i + 1, k + 1) = DateAdd("m", 1, DTPicker1.Value)
                        ElseIf k = 4 Or k = 8 Or k = 12 Then
                            oSheet.Cells(i + 1, k + 1) = DateAdd("m", 2, DTPicker1.Value)
                        ElseIf k = 5 Or k = 9 Or k = 13 Then
                            oSheet.Cells(i + 1, k + 1) = DateAdd("m", 3, DTPicker1.Value)
                        End If
                        oSheet.Cells(i + 1, k + 1).NumberFormat = "mmm-yy"
                    Else
                        oSheet.Cells(i + 1, k + 1) = .TextMatrix(i, k)
                    End If
                Next
            Next
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

Private Sub CmbRevision_DropDown()
    If Len(CmbDocument) > 0 Then
        qry = "select distinct on (fltpp_rev) fltpp_rev from loadcap_generate_h where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "' and fltpp_doc='" & CmbDocument & "'"
        Set RsA = Con.Execute(qry)
        CmbRevision.Clear
        If RsA.RecordCount > 0 Then
            period1 = Format(DTPicker1.Value, "yyyyMM")
            period2 = Format(DateAdd("m", 1, DTPicker1.Value), "yyyyMM") 'Left(period1, 4) & Right("00" & Val(Right(period1, 2) + 1), 2)
            period3 = Format(DateAdd("m", 2, DTPicker1.Value), "yyyyMM") 'Left(period2, 4) & Right("00" & Val(Right(period2, 2) + 1), 2)
            period4 = Format(DateAdd("m", 3, DTPicker1.Value), "yyyyMM") 'Left(period3, 4) & Right("00" & Val(Right(period3, 2) + 1), 2)
'            ListPeriodToSearch = "'" & period1 & "','" & period2 & "','" & period3 & "','" & period4 & "'"
'            MsgBox ListPeriodToSearch
            With agrid3
                For i = 1 To .Cols - 1
                    If i = 6 Or i = 10 Or i = 14 Then
                        .TextMatrix(1, i) = Format(DTPicker1.Value, "mmm-yy")   'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
                    ElseIf i = 7 Or i = 11 Or i = 15 Then
                        .TextMatrix(1, i) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
                    ElseIf i = 8 Or i = 12 Or i = 16 Then
                        .TextMatrix(1, i) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
                    ElseIf i = 9 Or i = 13 Or i = 17 Then
                        .TextMatrix(1, i) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
                    End If
                Next
            End With
            With agrid4
                .TextMatrix(1, 1) = Format(DTPicker1.Value, "mmm-yy") ' nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 2) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 3) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 4) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 5) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy") & "[balc]"
                .TextMatrix(1, 6) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy") & "[balc]"
                .TextMatrix(1, 7) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy") & "[balc]"
                .TextMatrix(1, 8) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy") & "[balc]"
                .TextMatrix(1, 9) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 10) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 11) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 12) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 13) = Format(DTPicker1.Value, "mmm-yy") 'nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 14) = Format(DateAdd("m", 1, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 15) = Format(DateAdd("m", 2, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DTPicker1, "yy")
                .TextMatrix(1, 16) = Format(DateAdd("m", 3, DTPicker1.Value), "mmm-yy") 'nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DTPicker1, "yy")
            End With
            While Not RsA.EOF
                CmbRevision.AddItem RsA(0)
                RsA.MoveNext
            Wend
        End If
    Else
        CmbRevision.Clear
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

Private Sub settingLV()
'    With agrid
'        .Cols = 5: .FixedCols = 0
'        .TextMatrix(0, 0) = "Item ID"
'        .TextMatrix(0, 1) = "Prod Plan"
'        .TextMatrix(0, 2) = "Total Cap/Month"
'        .TextMatrix(0, 3) = "Result"
'        .TextMatrix(0, 4) = "Subcont"
'        .ColWidth(0) = 2500
'        .ColWidth(1) = 1900
'        .ColWidth(2) = 1900
'        .ColAlignment(0) = flexAlignLeftCenter
'    End With
    With agrid2
        .Cols = 6: .FixedCols = 0
        .TextMatrix(0, 0) = "Tonage"
        .TextMatrix(0, 1) = "Machine Off"
        .TextMatrix(0, 2) = "Balance Cap (day)"
        .TextMatrix(0, 3) = "Total unprocessed (day)"
        .TextMatrix(0, 4) = "Need (day)"
        .TextMatrix(0, 5) = "Need (machine)"
        .ColWidth(1) = 1100
        .ColWidth(2) = 1600
        .ColWidth(3) = 2090
        .ColWidth(4) = 1000
        .ColWidth(5) = 1490
        .ColAlignment(2) = flexAlignLeftCenter
    End With
    With agrid4
        .Cols = 17: .FixedCols = 0
        .rows = 3
        .FixedRows = 2
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
        .MergeCol(0) = True
        .Col = 0
        .Row = 0
        .CellBackColor = RGB(5, 122, 239)
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .Row = 1
        .CellBackColor = RGB(5, 122, 239)
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 0) = "Tonage"
        .TextMatrix(1, 0) = .TextMatrix(0, 0)
        .TextMatrix(0, 1) = "Total unprocessed (day)"
        .TextMatrix(0, 2) = .TextMatrix(0, 1)
        .TextMatrix(0, 3) = .TextMatrix(0, 1)
        .TextMatrix(0, 4) = .TextMatrix(0, 1)
        .TextMatrix(0, 5) = ""
        .TextMatrix(0, 6) = .TextMatrix(0, 5)
        .TextMatrix(0, 7) = .TextMatrix(0, 5)
        .TextMatrix(0, 8) = .TextMatrix(0, 5)
        .TextMatrix(0, 9) = ""
        .TextMatrix(0, 10) = .TextMatrix(0, 9)
        .TextMatrix(0, 11) = .TextMatrix(0, 9)
        .TextMatrix(0, 12) = .TextMatrix(0, 9)
        .TextMatrix(0, 13) = "Need (machine)"
        .TextMatrix(0, 14) = .TextMatrix(0, 13)
        .TextMatrix(0, 15) = .TextMatrix(0, 13)
        .TextMatrix(0, 16) = .TextMatrix(0, 13)
        For i = 0 To .Cols - 1
            .Row = 0
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            If i < 13 And i > 4 Then
                .ColWidth(i) = 0
            Else
                If i > 0 Then
                    .Col = i
                    .Row = 0
                    .CellBackColor = RGB(82, 167, 252)
                    .CellFontBold = True
                    .Row = 1
                    .CellBackColor = RGB(82, 167, 252)
                    .CellFontBold = True
                    .CellAlignment = flexAlignCenterCenter
                End If
                If i > 4 Then
                    .Col = i
                    .Row = 0
                    .CellBackColor = RGB(147, 206, 255)
                    .CellFontBold = True
                    .Row = 1
                    .CellBackColor = RGB(147, 206, 255)
                    .CellFontBold = True
                    .CellAlignment = flexAlignCenterCenter
                End If
            End If
        Next
    End With
    With agrid3
        .Cols = 18: .FixedCols = 0
        .rows = 3
        .FixedRows = 2
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
        .ColWidth(0) = 2500
        .WordWrap = True
        .Row = 0
        .Col = 0
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .Col = 0
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .Col = 1
        .Row = 0
        .CellAlignment = flexAlignCenterCenter
        .Col = 1
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .TextMatrix(0, 0) = "Item ID"
        .TextMatrix(1, 0) = .TextMatrix(0, 0)
        .MergeCol(0) = True
        .TextMatrix(0, 1) = "No Mold"
        .TextMatrix(1, 1) = .TextMatrix(0, 1)
        .MergeCol(1) = True
        .TextMatrix(0, 2) = "Subcont"
        .TextMatrix(1, 2) = .TextMatrix(0, 2)
        .MergeCol(2) = True
        .ColWidth(2) = 700
        .TextMatrix(0, 3) = "Cav Std"
        .TextMatrix(1, 3) = .TextMatrix(0, 3)
        .MergeCol(3) = True
        .ColWidth(3) = 600
        .TextMatrix(0, 4) = "Cav Act"
        .TextMatrix(1, 4) = .TextMatrix(0, 4)
        .MergeCol(4) = True
        .ColWidth(4) = 600
        .TextMatrix(0, 5) = "% Cav"
        .TextMatrix(1, 5) = .TextMatrix(0, 5)
        .MergeCol(5) = True
        .ColWidth(5) = 700
        .TextMatrix(0, 6) = "Cap/Month"
        .Row = 0
        .Col = 2
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .Col = 3
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .Col = 4
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .Col = 5
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .Col = 6
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        .TextMatrix(0, 7) = .TextMatrix(0, 6)
        .TextMatrix(0, 8) = .TextMatrix(0, 6)
        .TextMatrix(0, 9) = .TextMatrix(0, 6)
        .TextMatrix(0, 10) = "Prod Plan"
        .Row = 0
        .Col = 10
        .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 11) = .TextMatrix(0, 10)
        .TextMatrix(0, 12) = .TextMatrix(0, 10)
        .TextMatrix(0, 13) = .TextMatrix(0, 10)
        .Row = 0
        .Col = 14
        .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 14) = "Need Mold"
        .TextMatrix(0, 15) = .TextMatrix(0, 14)
        .TextMatrix(0, 16) = .TextMatrix(0, 14)
        .TextMatrix(0, 17) = .TextMatrix(0, 14)
        

'        .TextMatrix(0, 15) = "Cav Std"
'        .TextMatrix(1, 15) = "Cav Std"
'        .TextMatrix(0, 16) = "Cav Act"
'        .TextMatrix(1, 16) = .TextMatrix(0, 16)
'        .TextMatrix(0, 17) = "% Cav"
'        .TextMatrix(1, 17) = .TextMatrix(0, 17)
        For i = 2 To .Cols - 1
            .Row = 1
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
        Next
        For i = 0 To 1
            .Col = i
            .Row = 0
            .CellBackColor = RGB(53, 183, 30)
            .Col = i
            .Row = 1
            .CellBackColor = RGB(53, 183, 30)
        Next
        For i = 2 To 5
            .Row = 0
            .Col = i
            .CellBackColor = RGB(60, 206, 34)
            .CellFontBold = True
            .Row = 1
            .Col = i
            .CellBackColor = RGB(60, 206, 34)
            .CellFontBold = True
        Next
        For i = 6 To 9
            .Row = 0
            .Col = i
            .CellBackColor = RGB(102, 226, 80)
            .CellFontBold = True
            .Row = 1
            .Col = i
            .CellBackColor = RGB(102, 226, 80)
            .CellFontBold = True
        Next
        For i = 10 To .Cols - 1
            .Row = 0
            .Col = i
            .CellBackColor = RGB(211, 254, 205)
            .CellFontBold = True
            .Row = 1
            .Col = i
            .CellBackColor = RGB(211, 254, 205)
            .CellFontBold = True
        Next
        .Refresh
    End With
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call settingLV
    Call activeTheme(skinFD, Me)
    DTPicker1.Value = Now
    Me.Height = 7830
    Me.Width = 12600
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
    cmbFiletype.ListIndex = 0
    cmbFiletype2.ListIndex = 0
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
    With cmbnPeriod
        .Left = SkinLabel5.Left: .Top = SkinLabel4.Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Function checkArrUni(pmould As String) As Boolean
    Dim it As Integer
    For it = 1 To UBound(arrMold)
'        MsgBox "If" & arrMold(it) & "=" & pmould & "Then "
        If arrMold(it) = pmould Then
'            MsgBox "wuii"
            checkArrUni = True
            Exit For
        Else
            checkArrUni = False
        End If
    Next
End Function

Private Sub loadDetil_v2()
'    Dim rsQ As ADODB.Recordset, b As Integer, b2 As Integer
'    If agrid.rows > 1 Then
'        Me.Refresh
'        agrid.Refresh
'        For b2 = 1 To agrid.rows - 1
'            qry = "select * from " _
'                & " (select lc_itemid,lc_pp,lc_fprodtvty,fltpp_hkw " _
'                & " From loadcap_generate_h " _
'                & " where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & "  and fltpp_ym='" & cmbnPeriod & "' and lc_itemid='" & agrid.TextMatrix(b2, 0) & "' " _
'                & " order by lc_itemid asc) x " _
'                & " Inner Join " _
'                & " (select lcd_itemdid,cav,ct,cap_p_day,reg_mold,no_mach " _
'                & " From loadcap_generate_d " _
'                & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "' " _
'                & " order by lcd_itemdid asc) y on x.lc_itemid=y.lcd_itemdid " _
'                & " order by cap_p_day asc"
'             Set rsQ = Con.Execute(qry)
''             MsgBox qry
'             List1.Clear
'             ReDim arrMold(1 To 1)
'             b = 1
''             MsgBox rsQ.RecordCount, vbInformation, "wooow"
'             While Not rsQ.EOF
'                If checkArrUni(rsQ("reg_mold")) = False Then
'                    If UBound(arrMold) = 1 And arrMold(1) = "" Then
'                        arrMold(UBound(arrMold)) = rsQ("reg_mold")
'                    Else
'                        ReDim Preserve arrMold(1 To UBound(arrMold) + 1)
'                        arrMold(UBound(arrMold)) = rsQ("reg_mold")
'                    End If
'                End If
'                List1.AddItem rsQ("cap_p_day") * rsQ("fltpp_hkw") ' & "|" & rsQ("reg_mold") ' * rsA("fltpp_hkw")
'                rsQ.MoveNext
'            Wend
'            totalprodPlan = 0
'            For b = 1 To UBound(arrMold)
'                totalprodPlan = totalprodPlan * 1 + List1.List(b - 1) * 1
'            Next
''            MsgBox totalprodPlan
'
'            With agrid
'                .TextMatrix(b2, 2) = FormatNumber(totalprodPlan, 0)
'                .TextMatrix(b2, 3) = totalprodPlan - .TextMatrix(b2, 1) * 1
'                .TextMatrix(b2, 3) = FormatNumber(.TextMatrix(b2, 3), 0)
'                If Left(.TextMatrix(b2, 3), 1) = "-" Then
''                    .TextMatrix(b2, 4) = "Need"
'                    For j = 0 To .Cols - 1
'                        .Row = b2
'                        .Col = j
'                        .CellBackColor = RGB(255, 0, 0)
'                        .CellForeColor = RGB(255, 255, 255)
'                    Next
'                Else
''                    .TextMatrix(b2, 4) = "No"
'                End If
'
'            End With
'        Next
'    End If
End Sub

Private Sub loadDetil()
'    If lv1.ListItems.Count > 0 Then
'        Dim rsQ As ADODB.Recordset, b As Integer, b2 As Integer
'            qry = "select * from " _
'                & " (select lc_itemid,lc_pp,lc_fprodtvty,fltpp_hkw " _
'                & " From loadcap_generate_h " _
'                & " where lc_pp>0 and fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & "  and fltpp_ym='" & cmbnPeriod & "' and lc_itemid='" & lv1.SelectedItem.Text & "' " _
'                & " order by lc_itemid asc) x " _
'                & " Inner Join " _
'                & " (select lcd_itemdid,cav,ct,cap_p_day,reg_mold,no_mach " _
'                & " From loadcap_generate_d " _
'                & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & CmbRevision & " and fltpp_ym='" & cmbnPeriod & "' " _
'                & " order by lcd_itemdid asc) y on x.lc_itemid=y.lcd_itemdid " _
'                & " order by cap_p_day asc"
'             Set rsQ = Con.Execute(qry)
'             List1.Clear
'             ReDim arrMold(1 To 1)
'             b = 1
'             While Not rsQ.EOF
'    '            MsgBox rsQ("cap_p_day").value & rsQ("no_mach")
'                If checkArrUni(rsQ("reg_mold")) = False Then
'                    If UBound(arrMold) = 1 And arrMold(1) = "" Then
'                        arrMold(UBound(arrMold)) = rsQ("reg_mold")
'                    Else
'                        ReDim Preserve arrMold(1 To UBound(arrMold) + 1)
'                        arrMold(UBound(arrMold)) = rsQ("reg_mold")
'                    End If
'                End If
'                List1.AddItem rsQ("cap_p_day") * rsQ("fltpp_hkw") ' & "|" & rsQ("reg_mold") ' * rsA("fltpp_hkw")
'                rsQ.MoveNext
'             Wend
'    '         MsgBox UBound(arrMold)
'            totalprodPlan = 0
'            For b = 1 To UBound(arrMold)
'                totalprodPlan = totalprodPlan * 1 + List1.List(b - 1) * 1
'            Next
'            MsgBox totalprodPlan
'            'lv1.ListItems(3).SubItems(2) = totalprodPlan
'    End If
End Sub

Private Sub Option1_Click()
    stateFilter = 0
End Sub

Private Sub Option2_Click()
    stateFilter = 1
End Sub

Private Sub Option3_Click()
    stateFilter = 2
End Sub
