VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_RC_menuloading 
   Caption         =   "Report of Menu Loading (Chart)"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   10410
   Begin VB.CommandButton cmdlu_docno 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4090
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox CmbDocument 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox txtRevision 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox cmbPeriod 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   5880
      OleObjectBlob   =   "Form_RC_menuloading2.frx":0000
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Form_RC_menuloading2.frx":0062
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Form_RC_menuloading2.frx":00C8
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5880
      OleObjectBlob   =   "Form_RC_menuloading2.frx":012E
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   1440
      OleObjectBlob   =   "Form_RC_menuloading2.frx":018A
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel lbl_hkw 
      Height          =   255
      Left            =   7080
      OleObjectBlob   =   "Form_RC_menuloading2.frx":01F4
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_RC_menuloading2.frx":025E
      Top             =   0
   End
End
Attribute VB_Name = "Form_RC_menuloading"
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
Dim HKWs As Variant
Private oExcel      As Object
Private oBook       As Object
Private oSheet      As Object
Private rsGrap As New ADODB.Recordset
Private ttlMch As Byte
Private mesin2 As String
Private presen As String


Private Sub cmbPeriod_Click()
'    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
'    Screen.MousePointer = 11
'    Const kolomtOs As String = "no_mach, ton_mach ,fltpp_hkw, lcd_itemdid,lc_itemname , cap_p_day, fltpp_ym, lcvsmach,neday,lcvsmach,lc_subcont,neqty "
'    qry = "select " & kolomtOs & " from mpp_gen_d where fltpp_doc='" & CmbDocument & "'" _
'        & " and fltpp_rev='" & txtRevision & "' and fltpp_ym='" & cmbPeriod & "' " _
'    & " order by no_mach asc, lc_customer asc, lcd_itemdid asc"
'
'    Set RsGet = Con.Execute(qry)
'
'    If RsGet.RecordCount > 0 Then
'
'    End If
'    Screen.MousePointer = 0
        
End Sub

Private Sub cmbPeriod_DropDown()
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    qry = "select distinct on (fltpp_ym) fltpp_ym from mpp_gen_d  where fltpp_doc='" & CmbDocument & "'" _
    & " and fltpp_rev='" & txtRevision & "'"
    Set RsGet = Con.Execute(qry)
    cmbPeriod.Clear
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            cmbPeriod.AddItem RsGet(0)
            RsGet.MoveNext
        Wend
    End If
End Sub

Private Sub cmdExport_Click()
    qry = "select no_mach, max(ton_mach) tonase ,max(fltpp_hkw) hkw, max(cap_p_day) cappday, sum(lcvsmach) plc,sum(neday) neday,sum(neqty) neqty  from mpp_gen_d " _
    & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev='" & txtRevision & "' and fltpp_ym='" & cmbPeriod & "'  and lc_subcont='no' and fltpp_ym='" & cmbPeriod & "' " _
    & " group by no_mach order by no_mach"
    Set rsGrap = Con.Execute(qry)
    If rsGrap.RecordCount > 0 Then
        lbl_hkw = rsGrap("hkw")
        ttlMch = rsGrap.RecordCount
        mesin2 = ""
        presen = ""
        For i = 1 To rsGrap.RecordCount
            rsGrap.AbsolutePosition = i
            mesin2 = mesin2 & rsGrap("no_mach") & "*"
            presen = presen & rsGrap("plc") & "*"
        Next
        mesin2 = Left(mesin2, Len(mesin2) - 1)
        presen = Left(presen, Len(presen) - 1)
        DoTheChart
        
    End If
    
End Sub

Private Sub cmdlu_docno_Click()
    PopUp_MLDOC.Show 1
    CmbDocument.Text = PopUp_MLDOC.lu_nodoc
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub CmbDocument_DropDown()
    qry = " select * from " _
        & "(select distinct on (fltpp_doc) fltpp_doc  from mpp_gen_d ) v1 " _
        & " order by right(fltpp_doc,4) asc,substring(fltpp_doc from 17 for 2)"
    Set RsGet = Con.Execute(qry)
    CmbDocument.Clear
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            CmbDocument.AddItem RsGet(0)
            RsGet.MoveNext
        Wend
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
    Call activeTheme(skinFD, Me)
    Me.Height = 7740
    Me.Width = 10650
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    CmbDocument.Left = SkinLabel5.Left
'    CmbDocument.Top = SkinLabel3.Top
    txtRevision.Top = SkinLabel2.Top
    txtRevision.Left = SkinLabel5.Left
    cmbPeriod.Left = lbl_hkw.Left
    cmbPeriod.Top = SkinLabel1.Top
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

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub txtRevision_DropDown()
    If Len(CmbDocument) < 2 Then CmbDocument.SetFocus: Exit Sub
    qry = "select distinct on (fltpp_rev) fltpp_rev from mpp_gen_d  where fltpp_doc='" & CmbDocument & "'"
    Set RsGet = Con.Execute(qry)
    txtRevision.Clear
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            txtRevision.AddItem RsGet(0)
            RsGet.MoveNext
        Wend
    End If
End Sub

Sub DoTheChart()
    Dim nRetval As Long
 
'    With RMChartX1
'        .Reset
'        .RMCBackColor = Linen
'        '.SetProperties 500, 400, Aquamarine, RMC_CTRLSTYLE3D, "", "Tahoma", 0, 0, Default, Default
'        '************** Add Region 1 *****************************
'        .AddRegion
'        With .Region(1)
'            .SetProperties 5, 5, -5, -5, Format(Now, "dd MMMM yyyy")
'            '************** Add caption to region 1 *******************
'            .AddCaption
'            With .Caption
'                .SetProperties "Menu Loading", Linen, Black, 10, True
'            End With 'Caption
'            '************** Add grid to region 1 *****************************
'            .AddGrid
'            With .grid
'                .SetProperties Honeydew, False, 0, 0, 0, 0, RMC_BICOLOR_NONE
'            End With 'Grid
'            '************** Add data axis to region 1 *****************************
'            .AddDataAxis
'            With .DataAxis(1)
'                .SetProperties RMC_DATAAXISLEFT, 0, 105, 8, 8, Default, Default, RMC_LINESTYLESOLID, 0, "%", "", "", RMC_TEXTCENTER
'            End With 'DataAxis(1)
'            '************** Add label axis to region 1 *****************************
'            .AddLabelAxis
'            With .LabelAxis
'                .SetProperties 1, ttlMch, RMC_LABELAXISBOTTOM, 8, Default, RMC_TEXTCENTER, Default, RMC_LINESTYLESOLID, ""
'                .LabelString = mesin2
'            End With 'LabelAxis
'            '************** Add legend to region 1 *******************************
'            .AddLegend
'            With .Legend
'                .SetProperties RMC_LEGEND_UL, Default, RMC_LEGENDRECT, Default, 8, False
'                .LegendString = ""
'            End With 'Legend
'            '************** Add Series 1 to region 1 *******************************
'            .AddBarSeries
'            With .BarSeries(1)
'                .SetProperties RMC_BARSINGLE, RMC_BAR_FLAT, False, DeepSkyBlue, False, 1, RMC_VLABEL_DEFAULT_NOZERO, 1, RMC_HATCHBRUSH_OFF
'                '****** Set data values ******
'                .DataString = presen
'            End With 'BarSeries(1)
'        End With 'Region(1)
'        nRetval = .Draw(True)
'    End With 'RMChartX1
End Sub


