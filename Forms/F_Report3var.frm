VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_Report3var 
   Caption         =   "Document Comparison"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14340
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
   ScaleHeight     =   7740
   ScaleWidth      =   14340
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   14115
      TabIndex        =   15
      Top             =   7080
      Width           =   14175
      Begin VB.TextBox txtTotalWO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtTotalMps 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtTotalLTPP 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Total WO"
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
         Left            =   9360
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "Total MPP Plan"
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
         Left            =   4440
         TabIndex        =   18
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Total LTPP Plan"
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
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   4680
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   120
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   0
         Width           =   4215
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblperc 
      Height          =   375
      Left            =   8760
      OleObjectBlob   =   "F_Report3var.frx":0000
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   495
      Left            =   11400
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Document"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox PicListMPP 
      BackColor       =   &H00C0FFC0&
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   14115
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   14175
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
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid fgmpp 
         Height          =   4215
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "double click or press Enter to load the data"
         Top             =   840
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   7435
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
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MPS document no"
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
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "FIND DOCUMENT"
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
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   13695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
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
         Height          =   255
         Left            =   13680
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11033
      _Version        =   393216
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "F_Report3var.frx":0058
      Top             =   0
   End
End
Attribute VB_Name = "F_Report3var"
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
Dim s_mps_doc As String
Dim s_mps_doc_rev As String
Dim s_ltpp_doc As String
Dim s_ltpp_doc_rev As String
Dim s_period As String
Dim i As Integer
Dim qry As String
Private posisisFind As Double
Dim counter As Long

Private Sub cmdfind_Click()
    If PicListMPP.Visible Then
        PicListMPP.Visible = False
    Else
        qry = "select * from (select distinct on (mpp_doc_no,mpp_revisi,ml_doc,ml_rev) mpp_doc_no,mpp_revisi,ml_ym,ml_rev,ml_doc  from mpp_gen ) v1 order by ml_ym desc, mpp_revisi desc limit 1"
        Set RsBantu = Con.Execute(qry)
        fgmpp.rows = 1
        If RsBantu.RecordCount > 0 Then
            fgmpp.rows = 1 + RsBantu.RecordCount
            fgmpp.TextMatrix(1, 0) = 1
            fgmpp.TextMatrix(1, 1) = RsBantu("mpp_doc_no")
            fgmpp.TextMatrix(1, 2) = RsBantu("mpp_revisi")
            fgmpp.TextMatrix(1, 3) = RsBantu("ml_ym")
            fgmpp.TextMatrix(1, 4) = RsBantu("ml_rev")
            fgmpp.TextMatrix(1, 5) = RsBantu("ml_doc")
        End If
        PicListMPP.Visible = True
        txtfind.SetFocus
    End If
    counter = 0
End Sub

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

Private Sub fgmpp_Click()
    With fgmpp
        s_mps_doc = .TextMatrix(.Row, 1)
        s_mps_doc_rev = .TextMatrix(.Row, 2)
        s_period = .TextMatrix(.Row, 3)
        s_ltpp_doc = .TextMatrix(.Row, 5)
    End With
    
End Sub

Private Sub fgmpp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText fgmpp.Clip
    End If
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText grid1.Clip
    ElseIf KeyCode = 70 And Shift = 2 Then
        PicFIND.Visible = True
        txtFindNext.SetFocus
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



Private Sub txtFindNext_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    ElseIf KeyAscii = vbKeyEscape Then
        PicFIND.Visible = False
    ElseIf KeyAscii = 1 Then
        txtFindNext.SelStart = 0
        txtFindNext.SelLength = Len(txtFindNext.Text)
    End If
End Sub

Private Sub fgmpp_DblClick()
    PicListMPP.Visible = False
    loadDokumen
    
End Sub

Private Sub fgmpp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fgmpp_DblClick
    End If
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
End Sub

Private Sub Form_Load()
On Error GoTo Ex
    AddTab Me
    BukaKoneksi
    Call activeTheme(skinFD, Me)
    settingFG
    Height = 8310
    Width = 14580
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

Private Sub loadDokumen()
    Dim nom As Long
    Dim ttlRow As Long
    Dim ttlMPP As Long
    Dim ttlLTPP As Long
    Dim ttlWO As Long
    Dim mppPeriod As String
    DoEvents
    lblperc.Caption = "Please Wait"
    mppPeriod = Right$(s_period, 2) & "/" & Mid(s_period, 3, 2)
    
    qry = "select lcd_itemdid,assy_no,item_name,mpsqty,cust_name, prod_plan_1, coalesce(ttlwo,0) ttlwo from " _
    & " (select assy_no,a.item_name,prod_plan_1,ltpp_doc,cust_name from ltpp_generate a inner join mst_item b on a.assy_no=b.item_id inner join r_customer c on b.cust_id=c.cust_id where ltpp_doc='" & s_ltpp_doc & "') v3 left join " _
        & " (select lcd_itemdid,ml_doc,sum(planqty) mpsqty from mpp_gen a " _
    & " where mpp_doc_no='" & s_mps_doc & "' and mpp_revisi='" & s_mps_doc_rev & "' and ml_doc='" & s_ltpp_doc & "' " _
    & " group by lcd_itemdid,ml_doc) v1 on v3.assy_no=v1.lcd_itemdid and v3.ltpp_doc=v1.ml_doc left join ( " _
    & " select partno,sum(coalesce(qty,0)) ttlwo from worko " _
    & " where substring(mpp_doc from 12 for 5)='" & mppPeriod & "' " _
    & " group by partno) v2 on v1.lcd_itemdid=v2.partno " _
    & " order by cust_name asc, assy_no asc"
    
    Set RsBantu = Con.Execute(qry)
    nom = 1
    ttlRow = RsBantu.RecordCount
    If ttlRow > 0 Then
        With grid1
            .rows = 1
            .rows = 1 + ttlRow
            While Not RsBantu.EOF
                .TextMatrix(nom, 0) = nom
                .TextMatrix(nom, 1) = RsBantu("cust_name")
                .TextMatrix(nom, 2) = Trim(RsBantu("assy_no"))
                If IsNull(RsBantu("lcd_itemdid")) Then
                    .TextMatrix(nom, 3) = ""
                    .Col = 3
                    .Row = nom
                    .CellBackColor = RGB(255, 85, 42)
                Else
                    .TextMatrix(nom, 3) = RsBantu("lcd_itemdid")
                End If
                .TextMatrix(nom, 4) = RsBantu("item_name")
                .TextMatrix(nom, 5) = FormatNumber(RsBantu("prod_plan_1"), 0)
                .TextMatrix(nom, 6) = FormatNumber(RsBantu("mpsqty"), 0)
                .TextMatrix(nom, 7) = FormatNumber(RsBantu("ttlwo"), 0)
                ttlLTPP = RsBantu("prod_plan_1") + ttlLTPP
                ttlMPP = IIf(IsNull(RsBantu("mpsqty")), 0, RsBantu("mpsqty")) + ttlMPP
                ttlWO = RsBantu("ttlwo") + ttlWO
                pb1.Value = (nom / ttlRow) * 100
                lblperc.Caption = pb1.Value & "%"
                nom = nom + 1
                RsBantu.MoveNext
            Wend
        End With
        pb1.Value = 0
        lblperc.Caption = ""
        txtTotalLTPP = FormatNumber(ttlLTPP, 0)
        txtTotalMps = FormatNumber(ttlMPP, 0)
        txtTotalWO = FormatNumber(ttlWO, 0)
    End If
    Set RsBantu = Nothing
End Sub

Private Sub settingFG()
    With grid1
        .Cols = 8
        .rows = 2
        .FixedRows = 1
        .RowHeight(0) = 500
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter

        .MergeCells = flexMergeFree

        i = 0
        .TextMatrix(0, i) = "No. "
        .ColWidth(i) = 700
        .ColAlignment(i) = flexAlignLeftCenter

        i = 1
        .TextMatrix(0, i) = "Customer"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800
        .MergeCol(1) = True


        i = 2
        .TextMatrix(0, i) = "Part No LTPP"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2500
        
        i = 3
        .TextMatrix(0, i) = "Part No MPS"
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2500

        i = 4
        .TextMatrix(0, i) = "Part Name"
        .ColWidth(i) = 4500

        i = 5
        .TextMatrix(0, i) = "LTPP Plan"
        .ColWidth(i) = 1200

        i = 6
        .TextMatrix(0, i) = "MPS Plan"
        .ColWidth(i) = 1200
        
        i = 7
        .TextMatrix(0, i) = "WO Qty"
        .ColWidth(i) = 1200

    End With
    With fgmpp
        .Cols = 6
        .FixedCols = 1
        .TextMatrix(0, 0) = "No"
        .ColWidth(0) = 500
        .TextMatrix(0, 1) = "MPS Doc No"
        .ColWidth(1) = 3000
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Rev"
        .ColWidth(2) = 500
        .TextMatrix(0, 3) = "Period"
        .TextMatrix(0, 4) = "Revisi LTPP"
        .ColWidth(4) = 0
        .TextMatrix(0, 5) = "LTPP Doc No"
        .ColWidth(5) = 3000
        .ColAlignment(5) = flexAlignLeftCenter
    End With
End Sub

Private Sub Label11_Click()
    PicListMPP.Visible = False
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim ttlData As Long
        txtfind = FilterIn(txtfind)
        qry = "select * from (select distinct on (mpp_doc_no,mpp_revisi,ml_doc,ml_rev) mpp_doc_no,mpp_revisi,ml_ym,ml_rev,ml_doc  from mpp_gen where mpp_doc_no like '%" & txtfind & "%' ) v1 order by ml_ym desc, mpp_revisi desc limit 3"
        Set RsBantu = Con.Execute(qry)

        fgmpp.rows = 1
        If RsBantu.RecordCount > 0 Then
            
            RsBantu.Fields("mpp_doc_no").Properties("Optimize") = True
            RsBantu.Fields("mpp_revisi").Properties("Optimize") = True
            With fgmpp
                If Len(Trim(txtfind)) > 0 Then
                    RsBantu.Filter = adFilterNone
                    RsBantu.Filter = "mpp_doc_no LIKE '*" & txtfind & "*'"
                    If RsBantu.RecordCount > 0 Then
                        .rows = RsBantu.RecordCount + 1
                        ttlData = RsBantu.RecordCount
                        For i = 1 To ttlData
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
                            ttlData = RsBantu.RecordCount
                            For i = 1 To ttlData
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
                    
                    ttlData = RsBantu.RecordCount
                    For i = 1 To ttlData
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


