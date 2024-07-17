VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_RekapMPPMCH 
   Caption         =   "Data Load Vs Cap Machine"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
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
   ScaleHeight     =   7485
   ScaleWidth      =   11265
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   5280
   End
   Begin VB.PictureBox PicEm 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   7440
      ScaleHeight     =   915
      ScaleWidth      =   3480
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   3533
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Enter value"
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
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.PictureBox PicListMPP 
      BackColor       =   &H00C0FFC0&
      Height          =   4335
      Left            =   480
      ScaleHeight     =   4275
      ScaleWidth      =   10155
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox txtFind 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid fgmpp 
         Height          =   3195
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "double click or press Enter to load saved data"
         Top             =   960
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   5636
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9720
         TabIndex        =   8
         Top             =   0
         Width           =   430
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "MPP Data List"
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   9735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3375
      Left            =   45
      TabIndex        =   1
      Top             =   960
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5953
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter"
      Height          =   855
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.ComboBox cmbFiletype 
         Height          =   390
         ItemData        =   "Form_RekapMPPMCH.frx":0000
         Left            =   2640
         List            =   "Form_RekapMPPMCH.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   2655
      End
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   135
         Left            =   8760
         TabIndex        =   16
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3600
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblHKW 
         Height          =   255
         Left            =   8760
         OleObjectBlob   =   "Form_RekapMPPMCH.frx":0023
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   135
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_RekapMPPMCH.frx":0081
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   1095
      Left            =   45
      TabIndex        =   11
      Top             =   6360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1931
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grid3 
      Height          =   1815
      Left            =   45
      TabIndex        =   12
      Top             =   4440
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "Form_RekapMPPMCH"
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
Dim NoDocMPS As String
Dim rev_MPS As String
Dim ltppdoc As String
Dim ltpprev As String
Dim period As String
Dim tonase() As String
Dim aktif_grid As String
Dim aktif_grid_x As Byte
Dim aktif_grid_y As Byte
Private oExcel      As Object
Private oBook       As Object
Private oSheet      As Object

Private Declare Function GetCursorPos Lib "user32" (lpPoint As _
   POINTAPI) As Long

Private Type POINTAPI
        x As Long
        Y As Long
End Type
Dim a As POINTAPI
Dim b As Long
Dim c As Long

Private Sub mousepos()
Dim Ret As Long
Ret = GetCursorPos(a)
b = a.x * Screen.TwipsPerPixelX
c = a.Y * Screen.TwipsPerPixelY - (15 / 100 * (a.Y * Screen.TwipsPerPixelY))

End Sub

Private Sub loadData(pmps As String, pmps_rev As String, pltppdoc As String, pperiod As String, pltpprev As String)
    Dim l As Byte
    Dim c As Byte
    Dim tempton As String
    qry = "select v1.no_mach,max(ton_mach) tonase,sum(lcvsmach) lcvsmach,rstate_mach,max(partname) pn,max(ml_hkw) hkw from " _
    & " (select distinct on (no_mach,lcd_itemdid,reg_mold) no_mach,lcvsmach, ton_mach,ml_hkw from mpp_gen " _
    & " where mpp_doc_no='" & pmps & "' and mpp_revisi='" & pmps_rev & "' and ml_subcont='no') as v1 " _
    & " inner join (select no_mach,string_agg(distinct(partname),',') partname from mpp_gen " _
    & " where mpp_doc_no='" & pmps & "' and mpp_revisi='" & pmps_rev & "' and ml_subcont='no' " _
    & " group by no_mach) as v2 on v1.no_mach=v2.no_mach " _
    & " inner join (select no_mach ,rstate_mach, sum(lcvsmach) lc from mpp_gen_d " _
    & " where fltpp_doc='" & pltppdoc & "' " _
    & " and fltpp_rev='" & pltpprev & "' and fltpp_ym='" & pperiod & "' " _
    & " group by no_mach, rstate_mach " _
    & " order by no_mach asc) as v3 on v1.no_mach=v3.no_mach " _
    & " group by v1.no_mach ,rstate_mach " _
    & " order by v1.no_mach"
    Set RsGet = Con.Execute(qry)
    If RsGet.RecordCount > 0 Then
        lblHKW.Caption = RsGet("hkw") & " Hari kerja"
        With grid1
            .rows = 2
            .rows = 2 + RsGet.RecordCount
            l = 2
            While Not RsGet.EOF
                If RsGet("rstate_mach") = False Then
                    .Row = l
                    .Col = 0
                    .CellBackColor = vbRed
                End If
                .TextMatrix(l, 0) = RsGet("no_mach") & IIf(RsGet("rstate_mach") = False, " (OFF)", "")
                .TextMatrix(l, 1) = RsGet("pn")
                l = l + 1
                RsGet.MoveNext
            Wend
        End With
        RsGet.MoveFirst
        RsGet.Sort = "tonase asc"
        Erase tonase
        For l = 1 To RsGet.RecordCount
            RsGet.AbsolutePosition = l
            If RsGet("tonase") <> tempton Then
                If (Not tonase) <> -1 Then
                    ReDim Preserve tonase(1 To UBound(tonase) + 1) As String
                    tonase(UBound(tonase)) = RsGet("tonase")
                Else
                    ReDim tonase(1 To 1) As String
                    tonase(1) = RsGet("tonase")
                End If
            End If
            tempton = RsGet("tonase")
        Next
        grid1.Cols = UBound(tonase) + 2
        grid1.FixedCols = 2
        grid2.Cols = UBound(tonase) + 1
        grid3.Cols = UBound(tonase) + 1
        For c = 1 To UBound(tonase)
            grid1.TextMatrix(1, c + 1) = tonase(c) & "T"
            grid1.Col = c + 1
            grid1.Row = 1
            grid1.CellAlignment = flexAlignCenterCenter
            grid1.CellFontBold = True
            grid1.TextMatrix(0, c + 1) = "Tonage"
            grid1.Col = c + 1
            grid1.Row = 0
            grid1.CellAlignment = flexAlignCenterCenter
            grid1.CellFontBold = True
            
            '----grid 2
            grid2.TextMatrix(0, c) = tonase(c) & "T"
            grid2.Col = c
            grid2.Row = 0
            grid2.CellAlignment = flexAlignCenterCenter
            grid2.CellFontBold = True
            grid2.TextMatrix(2, c) = 0
            
            '----grid 3
            grid3.TextMatrix(0, c) = tonase(c) & "T"
            grid3.Col = c
            grid3.Row = 0
            grid3.CellAlignment = flexAlignCenterCenter
            grid3.CellFontBold = True
            grid3.TextMatrix(2, c) = 0
        Next
        RsGet.Sort = ""
        RsGet.Sort = "no_mach asc"
        With grid1
            For l = 1 To RsGet.RecordCount
                RsGet.AbsolutePosition = l
                For c = 2 To .Cols - 1
                    If RsGet("tonase") & "T" = .TextMatrix(1, c) Then
                        .TextMatrix(l + 1, c) = RsGet("lcvsmach")
                        Exit For
                    End If
                Next
                If RsGet("no_mach") <> .TextMatrix(l + 1, 0) Then
                    For c = 2 To .Cols - 1
                        If IsNumeric(.TextMatrix(l + 1, c)) Then
                            If IsNumeric(grid3.TextMatrix(3, c - 1)) Then
                                grid3.TextMatrix(3, c - 1) = grid3.TextMatrix(3, c - 1) * 1 + 1
                            Else
                                grid3.TextMatrix(3, c - 1) = 1
                            End If
                        End If
                    Next
                End If
            Next
        End With
        hitunghitung
    End If
End Sub



Private Sub hitunghitung()
    Dim kol As Byte
    
    'reset presentase
    With grid2
         For kol = 1 To .Cols - 1
            .TextMatrix(2, kol) = 0
         Next
    End With
    
    With grid1
        For i = 2 To .rows - 1
            For kol = 2 To .Cols - 1
                If IsNumeric(.TextMatrix(i, kol)) Then
                    If IsNumeric(grid2.TextMatrix(2, kol - 1)) Then
                        grid2.TextMatrix(2, kol - 1) = grid2.TextMatrix(2, kol - 1) * 1 + .TextMatrix(i, kol) * 1
                    Else
                        grid2.TextMatrix(2, kol - 1) = .TextMatrix(i, kol)
                    End If
                End If
            Next
        Next
    End With
    With grid2
        For kol = 1 To .Cols - 1
            If IsNumeric(.TextMatrix(2, kol)) And IsNumeric(.TextMatrix(1, kol)) Then
                If .TextMatrix(1, kol) > 0 Then
                    .TextMatrix(2, kol) = FormatNumber(.TextMatrix(2, kol) / .TextMatrix(1, kol), 0) & "%"
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdExport_Click()
    If grid1.TextMatrix(2, 0) = "" Then Exit Sub
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
        oExcel.DisplayAlerts = False
        Dim k As Byte
        Dim ttl As Integer
        Dim il As Integer
        Dim iaw As Integer
        Screen.MousePointer = 11
        prog1.Visible = True
        With grid1
            ttl = .rows - 1
            For i = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If i = 0 Then
                        oSheet.Cells(i + 2, k + 1).Font.Bold = True
                        oSheet.Cells(i + 2, k + 1) = .TextMatrix(i, k)
                    Else
                        oSheet.Cells(i + 2, k + 1) = .TextMatrix(i, k)
                    End If
                Next
                prog1.Value = (i * 100) / ttl
            Next
        End With
        With oSheet
            .Cells(1, 1) = "Export Date : " & Now
            .Range("A2:A3").Merge
            .Range("B2:B3").Merge
            .Range(.Cells(2, 3), .Cells(2, grid1.Cols)).Merge
            .Range("A2").HorizontalAlignment = xlCenter
            .Range("A2").VerticalAlignment = xlCenter
            .Range("A2").WrapText = True
            .Range("B2").HorizontalAlignment = xlCenter
            .Range("B2").VerticalAlignment = xlCenter
            .Range("B2").WrapText = True
            .Range("C2").HorizontalAlignment = xlCenter
            .Range("C2").VerticalAlignment = xlCenter
            .Range("C2").WrapText = True
            .Range(.Cells(2, 1), .Cells(ttl + 2, grid1.Cols)).Borders.LineStyle = xlContinuous
'            .Range("A2:K" & ttl + 2).Borders.LineStyle = xlContinuous
        End With
        i = i + 4
        iaw = i
        With grid3
            ttl = .rows - 1
            For il = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If k = 0 Then
                        oSheet.Range("A" & i & ":B" & i).Merge
                        If il = 0 Then
                            oSheet.Cells(i, k + 1).Font.Bold = True
                            oSheet.Cells(i, k + 1) = .TextMatrix(il, k)
                        Else
                            oSheet.Cells(i, k + 1) = .TextMatrix(il, k)
                        End If
                    Else
                        If il = 0 Then
                            oSheet.Cells(i, k + 2).Font.Bold = True
                            oSheet.Cells(i, k + 2) = .TextMatrix(il, k)
                        Else
                            oSheet.Cells(i, k + 2) = .TextMatrix(il, k)
                        End If
                    End If
                Next
                prog1.Value = (il * 100) / ttl
                i = i + 1
            Next
        End With
        With oSheet
            .Range(.Cells(iaw, 1), .Cells(i - 1, grid1.Cols)).Borders.LineStyle = xlContinuous
        End With
        
        i = i + 1
        iaw = i
        With grid2
            ttl = .rows - 1
            For il = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If k = 0 Then
                        oSheet.Range("A" & i & ":B" & i).Merge
                        If il = 0 Then
                            oSheet.Cells(i, k + 1).Font.Bold = True
                            oSheet.Cells(i, k + 1) = .TextMatrix(il, k)
                        Else
                            oSheet.Cells(i, k + 1) = .TextMatrix(il, k)
                        End If
                    Else
                        If il = 0 Then
                            oSheet.Cells(i, k + 2).Font.Bold = True
                            oSheet.Cells(i, k + 2) = .TextMatrix(il, k)
                        Else
                            oSheet.Cells(i, k + 2) = .TextMatrix(il, k)
                        End If
                    End If
                Next
                prog1.Value = (il * 100) / ttl
                i = i + 1
            Next
        End With
        With oSheet
            .Range(.Cells(iaw, 1), .Cells(i - 1, grid1.Cols)).Borders.LineStyle = xlContinuous
        End With
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
        prog1.Visible = False
        Screen.MousePointer = 0
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub cmdfind_Click()
    PicListMPP.Visible = True
    txtfind.SetFocus
End Sub

Private Sub fgmpp_DblClick()
    fgmpp_KeyPress 13
End Sub

Private Sub fgmpp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText fgmpp.Clip
    End If
End Sub

Private Sub fgmpp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With fgmpp
            Screen.MousePointer = 11
            NoDocMPS = .TextMatrix(.Row, 1)
            rev_MPS = .TextMatrix(.Row, 2)
            ltppdoc = .TextMatrix(.Row, 5)
            ltpprev = .TextMatrix(.Row, 4)
            period = .TextMatrix(.Row, 3)
            loadData NoDocMPS, rev_MPS, ltppdoc, period, ltpprev
            Screen.MousePointer = 0
        End With
        PicListMPP.Visible = False
    End If
    
End Sub

Private Sub Form_Activate()
    FocusTab Me
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
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingFGrid
    Call WheelHook(Me.hwnd)
    prog1.Value = 0
    Me.Height = 8055
    Me.Width = 11505
End Sub

Private Sub settingFGrid()
    With grid1
        .rows = 3
        .Cols = 2
        .FixedRows = 2
        .FixedCols = 0
        .MergeCells = flexMergeFree
        .TextMatrix(0, 0) = "Machine"
        .TextMatrix(1, 0) = .TextMatrix(0, 0)
        .ColWidth(0) = 1500
        .ColWidth(1) = 2500
        .MergeRow(0) = True
        .MergeCol(0) = True
        .TextMatrix(0, 1) = "Product"
        .TextMatrix(1, 1) = .TextMatrix(0, 1)
        .MergeRow(1) = True
        .MergeCol(1) = True
    End With
    With grid2
        .rows = 3
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .MergeCells = flexMergeFree
        .TextMatrix(0, 0) = "Tonage"
        .TextMatrix(1, 0) = "Total (unit)"
        .TextMatrix(2, 0) = "Load VS Cap"
        .ColWidth(0) = 4000
    End With
    With grid3
        .rows = 5
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .MergeCells = flexMergeFree
        .TextMatrix(0, 0) = "Mesin"
        .TextMatrix(1, 0) = "Mesin Lama"
        .TextMatrix(2, 0) = "Mesin Tambahan"
        .TextMatrix(3, 0) = "Mesin Mati"
        .TextMatrix(4, 0) = "Total Mesin Injection"
        .ColWidth(0) = 4000
    End With
    With fgmpp
        .Cols = 6
        .FixedCols = 1
        .TextMatrix(0, 0) = "No"
        .ColWidth(0) = 500
        .TextMatrix(0, 1) = "Doc No"
        .ColWidth(1) = 3000
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Rev"
        .ColWidth(2) = 500
        .TextMatrix(0, 3) = "Period"
'        .ColWidth(3) = 0
        .TextMatrix(0, 4) = "Revisi LTPP"
'        .ColWidth(4) = 0
        .TextMatrix(0, 5) = "No LTPP"
        .ColWidth(5) = 3000
        .ColAlignment(5) = flexAlignLeftCenter
    End With
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbFiletype.Left = Label3.Left
    cmbFiletype.Top = cmdExport.Top
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
    If Cancel = 0 Then
        DelTab Me
    End If
End Sub

Private Sub grid1_Click()
    PicEm.Visible = False
End Sub

Private Sub grid2_Click()
    PicEm.Visible = False
End Sub

Private Sub grid3_Click()
    With grid3
        If .Row = 1 Or .Row = 2 And .Col > 0 Then
            PicEm.Top = c
            PicEm.Left = b
            aktif_grid_x = .Col
            aktif_grid_y = .Row
            PicEm.Visible = True
            txtEdit.Text = .Text
            txtEdit.SetFocus
            If txtEdit <> "" Then
                txtEdit.SelStart = 0
                txtEdit.SelLength = Len(txtEdit)
            End If
            aktif_grid = .Name
        Else
            PicEm.Visible = False
        End If
    End With
End Sub

Private Sub Label2_Click()
    PicListMPP.Visible = False
End Sub

Private Sub Timer1_Timer()
    mousepos
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(txtEdit) = False Then PicEm.Visible = False
        If aktif_grid = "grid3" Then
            Dim ioldmch As Integer
            Dim iaddmch As Integer
            Dim ioffmch As Integer
            
            With grid3
                .TextMatrix(aktif_grid_y, aktif_grid_x) = txtEdit
                If IsNumeric(.TextMatrix(1, aktif_grid_x)) Then
                    ioldmch = .TextMatrix(1, aktif_grid_x)
                Else
                    ioldmch = 0
                End If
                If IsNumeric(.TextMatrix(2, aktif_grid_x)) Then
                    iaddmch = .TextMatrix(2, aktif_grid_x)
                Else
                    iaddmch = 0
                End If
                If IsNumeric(.TextMatrix(3, aktif_grid_x)) Then
                    ioffmch = .TextMatrix(3, aktif_grid_x)
                Else
                    ioffmch = 0
                End If
                .TextMatrix(4, aktif_grid_x) = ioldmch + iaddmch - ioffmch
                grid2.TextMatrix(1, aktif_grid_x) = .TextMatrix(4, aktif_grid_x)
                
            End With
            hitunghitung
        End If
        PicEm.Visible = False
    ElseIf KeyAscii = vbKeyEscape Then
        PicEm.Visible = False
    End If
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        qry = "select * from (select distinct on (mpp_doc_no,mpp_revisi,ml_doc,ml_rev) mpp_doc_no,mpp_revisi,ml_ym,ml_rev,ml_doc  from mpp_gen where mpp_doc_no like '%" & txtfind & "%' ) v1 order by ml_ym desc, mpp_revisi desc limit 3"
        Set RsBantu = Con.Execute(qry)
        fgmpp.rows = 1
        If RsBantu.RecordCount > 0 Then
            RsBantu.Sort = "ml_ym desc"
            RsBantu.Fields("mpp_doc_no").Properties("Optimize") = True
            RsBantu.Fields("mpp_revisi").Properties("Optimize") = True
            With fgmpp
                If Len(Trim(txtfind)) > 0 Then
                    RsBantu.Filter = adFilterNone
                    RsBantu.Filter = "mpp_doc_no LIKE '*" & txtfind & "*'"
                    If RsBantu.RecordCount > 0 Then
                        .rows = RsBantu.RecordCount + 1
                        For i = 1 To RsBantu.RecordCount
                            RsBantu.AbsolutePosition = i
                            .TextMatrix(i, 0) = i
                            .TextMatrix(i, 1) = RsBantu("mpp_doc_no")
                            .TextMatrix(i, 2) = RsBantu("mpp_revisi")
                            .TextMatrix(i, 3) = RsBantu("ml_ym")
                            .TextMatrix(i, 4) = RsBantu("ml_rev")
                            .TextMatrix(i, 5) = RsBantu("ml_doc")
                        Next
                    Else
                        RsBantu.Filter = adFilterNone
                        RsBantu.Filter = "mpp_revisi LIKE '*" & txtfind & "*'"
                        If RsBantu.RecordCount > 0 Then
                            
                            .rows = RsBantu.RecordCount + 1
                            For i = 1 To RsBantu.RecordCount
                                RsBantu.AbsolutePosition = i
                                .TextMatrix(i, 0) = i
                                .TextMatrix(i, 1) = RsBantu("mpp_doc_no")
                                .TextMatrix(i, 2) = RsBantu("mpp_revisi")
                                .TextMatrix(i, 3) = RsBantu("ml_ym")
                                .TextMatrix(i, 4) = RsBantu("ml_rev")
                                .TextMatrix(i, 5) = RsBantu("ml_doc")
                              
                            Next
                        Else
                            .rows = 1
                        End If
                    End If
                Else
                    .rows = RsBantu.RecordCount + 1
                  
                    For i = 1 To RsBantu.RecordCount
                        RsBantu.AbsolutePosition = i
                        .TextMatrix(i, 0) = i
                        .TextMatrix(i, 1) = RsBantu("mpp_doc_no")
                        .TextMatrix(i, 2) = RsBantu("mpp_revisi")
                        .TextMatrix(i, 3) = RsBantu("ml_ym")
                        .TextMatrix(i, 4) = RsBantu("ml_rev")
                        .TextMatrix(i, 5) = RsBantu("ml_doc")
                      
                    Next
                End If
            End With
        End If
    End If
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

