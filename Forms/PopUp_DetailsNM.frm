VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PopUp_DetailsNM 
   Caption         =   "Detail of selected data"
   ClientHeight    =   4785
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8415
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid agrid 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7435
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
      Left            =   240
      OleObjectBlob   =   "PopUp_DetailsNM.frx":0000
      Top             =   360
   End
End
Attribute VB_Name = "PopUp_DetailsNM"
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
Dim qry As String
Dim i As Integer
Dim oExcel      As Object
Dim oBook       As Object
Dim oSheet      As Object

Private Sub Command1_Click()
    If agrid.rows < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject("Ket.Application")
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        Dim k As Integer
        With agrid
            For i = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    oSheet.Cells(i + 1, k + 1) = .TextMatrix(i, k)
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

Private Sub settingLV()
    With agrid
        .Cols = 7: .FixedCols = 0
        .TextMatrix(0, 0) = "Part No"
        .TextMatrix(0, 1) = "Mold"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Cap/day (Min)"
        .TextMatrix(0, 5) = "Tonage"
        .TextMatrix(0, 6) = "Need Day"
        .TextMatrix(0, 2) = "Subcont"
        .ColWidth(0) = 2500
        .ColWidth(2) = 2300
        .ColAlignment(0) = flexAlignLeftCenter
    End With
End Sub

Private Sub loadData()
'    qry = "select lc_itemid,max(lc_itemname) itemname,max(lc_sisa_pp) sisa_pp,max(fltpp_hkw) hkw,max(lc_fprodtvty) prodtvty,max(no_mach) mach,min(ton_mach) tonase,max(cav) cav,max(ct) ct, " _
'            & " max(cap_p_day) cap_p_day,max(lc_sisa_pp)/max(cap_p_day) neday,every(rstate_mach) rstate_mach,max(lc_subcont) subcont from " _
'            & " (select lc_itemid,lc_itemname,lc_sisa_pp,fltpp_hkw,lc_fprodtvty,lc_subcont from loadcap_generate_h " _
'            & " where lc_sisa_pp>0 and fltpp_doc='" & F_ReportofNeedMM.CmbDocument & "' and fltpp_rev=" & F_ReportofNeedMM.CmbRevision & " and fltpp_ym='" & F_ReportofNeedMM.cmbnPeriod & "') AS v1 " _
'            & " Inner Join " _
'            & " (select lcd_itemdid,no_mach,ton_mach,cav,ct,cap_p_day,rstate_mach from loadcap_generate_d " _
'            & " where fltpp_doc='" & F_ReportofNeedMM.CmbDocument & "' and fltpp_rev=" & F_ReportofNeedMM.CmbRevision & " and fltpp_ym='" & F_ReportofNeedMM.cmbnPeriod & "' order by ton_mach asc) as v2 on v1.lc_itemid=v2.lcd_itemdid " _
'            & " group by lc_itemid"
    qry = "select lc_itemid,max(lc_itemname) itemname,max(lc_sisa_pp) sisa_pp,max(fltpp_hkw) hkw,max(lc_fprodtvty) prodtvty,max(no_mach) mach,min(ton_mach) tonase,max(cav) cav,max(ct) ct, " _
            & " max(cap_p_day) cap_p_day,case when max(cap_p_day)=0 then 0 else  max(lc_sisa_pp)/max(cap_p_day) end neday,every(rstate_mach) rstate_mach,max(lc_subcont) subcont,max(reg_mold) reg_mold from " _
            & " (select lc_itemid,lc_itemname,lc_sisa_pp,fltpp_hkw,lc_fprodtvty,lc_subcont from loadcap_generate_h " _
            & " where lc_sisa_pp>0 and fltpp_doc='" & F_ReportofNeedMM.CmbDocument & "' and fltpp_rev=" & F_ReportofNeedMM.CmbRevision & " and fltpp_ym='" & F_ReportofNeedMM.cmbnPeriod & "' and fltpp_period='" & Format(F_ReportofNeedMM.DTPicker1, "yyyyMM") & "') AS v1 " _
            & " Inner Join " _
            & " (select lcd_itemdid,no_mach,ton_mach,cav,ct,cap_p_day,rstate_mach,reg_mold from loadcap_generate_d " _
            & " where fltpp_doc='" & F_ReportofNeedMM.CmbDocument & "' and fltpp_rev=" & F_ReportofNeedMM.CmbRevision & " and fltpp_ym='" & F_ReportofNeedMM.cmbnPeriod & "' order by cap_p_day desc) as v2 on v1.lc_itemid=v2.lcd_itemdid " _
            & " group by lc_itemid"
'
'    Clipboard.Clear
'    Clipboard.SetText qry
    Set RsA = Con.Execute(qry)
    
    agrid.rows = 1
    If RsA.RecordCount > 0 Then
        i = 1
        agrid.rows = RsA.RecordCount + 1
        While Not RsA.EOF
            With agrid
                .TextMatrix(i, 0) = RsA(0) ' ITEM ID
                .TextMatrix(i, 1) = RsA("reg_mold")
                .TextMatrix(i, 3) = FormatNumber(RsA(2), 0) ' QTY
                .TextMatrix(i, 4) = FormatNumber(RsA("cap_p_day"), 0) ' CAP P DAY
                .TextMatrix(i, 5) = RsA("tonase")
                .TextMatrix(i, 6) = FormatNumber(RsA("neday"), 2)
                .TextMatrix(i, 2) = RsA("subcont")
            End With
            i = i + 1
            RsA.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
'On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call settingLV
    Call activeTheme(skinFD, Me)
    loadData
Exit Sub
'errLoad:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
'    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub
