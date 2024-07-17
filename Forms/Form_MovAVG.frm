VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_MovAVG 
   Caption         =   "Report Of Forecast Moving AVG"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
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
   ScaleHeight     =   6660
   ScaleWidth      =   9465
   Begin VB.PictureBox PicDet 
      BackColor       =   &H00C0FFC0&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   8955
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2990
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Tx History"
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   9015
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7858
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.CheckBox ckAll 
         Caption         =   "All"
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox CB1 
         Caption         =   "Show tx Before"
         Height          =   375
         Left            =   7320
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmbView 
         Caption         =   "View"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cmbBulan2 
         Height          =   390
         ItemData        =   "Form_MovAVG.frx":0000
         Left            =   4680
         List            =   "Form_MovAVG.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmbBulan1 
         Height          =   390
         ItemData        =   "Form_MovAVG.frx":008E
         Left            =   2880
         List            =   "Form_MovAVG.frx":00B6
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmbTahun1 
         Height          =   390
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "..."
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCustID 
         BackColor       =   &H00C0E0FF&
         Height          =   390
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Form_MovAVG.frx":011C
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Form_MovAVG.frx":0184
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "Form_MovAVG.frx":01E8
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Label1"
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6480
      OleObjectBlob   =   "Form_MovAVG.frx":0244
      Top             =   1200
   End
End
Attribute VB_Name = "Form_MovAVG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type CtrlProportion
    Heightproportion    As Single
    WidthProportion     As Single
    TopProportion       As Single
    LeftProportion      As Single
End Type
Dim proportionArray()   As CtrlProportion
Dim oExcel              As Object
Dim oBook               As Object
Dim oSheet              As Object
Public custID           As String
Dim ar_nmBulan(1 To 12) As String

Sub settingFGrid()
    With grid1
        .rows = 2
        .Cols = 2
        .FixedCols = 0
        .TextMatrix(0, 0) = "Part No"
        .ColWidth(0) = 2900
        .ColAlignment(0) = flexAlignLeftCenter
        .TextMatrix(0, 1) = "Part Name"
        .ColWidth(1) = 3400
        .ColAlignment(1) = flexAlignLeftCenter
    End With
    With grid2
        .rows = 2
        .Cols = 2
        .FixedCols = 0
        .TextMatrix(0, 0) = "Part No"
        .ColWidth(0) = 2900
        .ColAlignment(0) = flexAlignLeftCenter
        .TextMatrix(0, 1) = "Part Name"
        .ColWidth(1) = 3400
        .ColAlignment(1) = flexAlignLeftCenter
    End With
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

Private Sub CB1_Click()
    If CB1.Value Then
        PicDet.Visible = True
    Else
        PicDet.Visible = False
    End If
End Sub

Private Sub ckAll_Click()
    If ckAll.Value = vbChecked Then
        txtCustID = ""
        custID = ""
    End If
End Sub

Private Sub cmbBulan1_Click()
    cmbBulan2.ListIndex = cmbBulan1.ListIndex
End Sub

Private Sub cmbBulan2_Click()
    If cmbBulan2.ListIndex < cmbBulan1.ListIndex Then
        cmbBulan2.ListIndex = cmbBulan1.ListIndex
    End If
End Sub

Private Function getNumbMonth(prName As String) As Byte
    Dim u As Byte
    For u = 1 To UBound(ar_nmBulan)
        If ar_nmBulan(u) = prName Then
            getNumbMonth = u
            Exit For
        End If
    Next
End Function



Private Sub cmbView_Click()
    Dim qry As String
    Dim a As Long
    Dim jmlKol As Byte
    Dim k As Byte
    Dim k2 As Byte
    Dim kplus As Byte
    Dim tahunbef As Date
    Dim tahunsel As Date
    Dim rsrep As ADODB.Recordset
    Dim bln_a As Byte
    Dim thn_a As Integer
    Dim bagi As Byte
    Dim totalR As Single
    Const AVGREQ As Byte = 7
    Dim whereCust As String
    
    If txtCustID <> "" Then
        whereCust = " and a.cust_id='" & Trim$(custID) & "'"
    Else
        whereCust = ""
    End If
    
    If cmbTahun1.Text = "" Then cmbTahun1.SetFocus: Exit Sub
    tahunsel = DateSerial(cmbTahun1, cmbBulan1.ListIndex + 1, 1)
    tahunbef = DateAdd("m", -7, tahunsel)
    tahunsel = DateSerial(cmbTahun1, cmbBulan2.ListIndex + 1, 1)
    tahunsel = DateAdd("m", -1, tahunsel)
    tahunsel = dhLastDayInMonth(tahunsel)
      
    
    qry = "select distinct a.item_id,item_name from sod a inner join mst_item b on a.item_id=b.item_id " _
    & " where a.item_id not like '%TEST%' " & whereCust
    Set RsBantu = Con.Execute(qry)
    
    If RsBantu.RecordCount > 0 Then
        qry = "select extract(year from inv_date) tahun,extract(month from inv_date) bulan,item_id, sum(coalesce(sod_scanqty,0)) qty from sod a " _
        & " where inv_date between '" & Format(tahunbef, "yyyy-MM-dd") & "' and '" & Format(tahunsel, "yyyy-MM-dd") & "' " & whereCust _
        & " group by item_id,extract(year from inv_date),extract(month from inv_date) " _
        & " order by 1 asc, 2 asc, 3 asc"
        
        Set rsrep = Con.Execute(qry)
        a = 1
        With grid1
            .rows = a
            .rows = a + RsBantu.RecordCount
            .Cols = 2
            jmlKol = ((cmbBulan2.ListIndex + 1) - (cmbBulan1.ListIndex + 1)) + 1
            .Cols = 2 + jmlKol
            jmlKol = cmbBulan1.ListIndex + 1
            .FixedCols = 2
            For k = 2 To .Cols - 1
                .TextMatrix(0, k) = ar_nmBulan(jmlKol)
                jmlKol = jmlKol + 1
            Next
            RsBantu.Sort = "item_id asc"
            While Not RsBantu.EOF
                .TextMatrix(a, 0) = Trim$(RsBantu("item_id"))
                .TextMatrix(a, 1) = RsBantu("item_name")
                a = a + 1
                RsBantu.MoveNext
            Wend
        End With
        RsBantu.MoveFirst
        a = 1
        With grid2
            .rows = a
            .rows = a + RsBantu.RecordCount
            .Cols = 2
            .Cols = 2 + DateDiff("m", tahunbef, tahunsel) + 1 '7  '+ (cmbBulan2.ListIndex - cmbBulan1.ListIndex)
            jmlKol = Val(Format(tahunbef, "MM"))
            .FixedCols = 2
            For k = 2 To .Cols - 1
                If jmlKol > 12 Then jmlKol = 1
                .TextMatrix(0, k) = ar_nmBulan(jmlKol)
                jmlKol = jmlKol + 1
            Next
            While Not RsBantu.EOF
                .TextMatrix(a, 0) = Trim$(RsBantu("item_id"))
                .TextMatrix(a, 1) = RsBantu("item_name")
                a = a + 1
                RsBantu.MoveNext
            Wend
            bln_a = Val(Format(tahunbef, "MM"))
            thn_a = Val(Format(tahunbef, "yyyy"))
            For a = 1 To .rows - 1
                thn_a = Val(Format(tahunbef, "yyyy"))
                For k = 2 To .Cols - 1
                    If bln_a > 12 Then
                        bln_a = 1
                        thn_a = thn_a + 1
                    End If
                    qry = "item_id='" & .TextMatrix(a, 0) & "' and tahun=" & thn_a & " and bulan=" & getNumbMonth(.TextMatrix(0, k))
                    rsrep.Filter = qry
                    
                    'MsgBox rsrep.RecordCount & " all " & qry
                    If rsrep.RecordCount > 0 Then
                        .TextMatrix(a, k) = rsrep("qty")
                    End If
                    bln_a = bln_a + 1
                Next
            Next
        End With
        With grid1
            For a = 1 To .rows - 1
                bagi = 0
                totalR = 0
                kplus = 0
                For k = 2 To .Cols - 1
'                    If k = 2 Then
                        For k2 = k To k + 6 'grid2.Cols - 1
'                            MsgBox k2 & "[" & k & "]"
                            If IsNumeric(grid2.TextMatrix(a, k2)) Then
                                totalR = totalR + grid2.TextMatrix(a, k2) * 1
                                bagi = bagi + 1
                            End If
                        Next

                        If totalR > 0 And bagi > 0 Then
                            .TextMatrix(a, k) = totalR / bagi
                        End If
                        bagi = 0
                        totalR = 0

                Next
            Next
        End With
    End If
End Sub

Private Sub cmdfind_Click()
    GetForm = Me.Name
    popUp_Customer.Show 1
    cmbTahun1.SetFocus
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

Private Sub Form_Load()
    settingFGrid
    AddTab Me
    Width = 9705
    Height = 7230
    activeTheme Skin1, Me
    
    '------------+YEAR
    Dim tahun As Integer
    Dim dtbefore As Date
    Dim u As Integer
    dtbefore = DateAdd("yyyy", -1, Now)
    tahun = Format(dtbefore, "yyyy")
    For u = 1 To 3
        cmbTahun1.AddItem tahun
        tahun = tahun + 1
    Next
    
    ar_nmBulan(1) = "January"
    ar_nmBulan(2) = "February"
    ar_nmBulan(3) = "March"
    ar_nmBulan(4) = "April"
    ar_nmBulan(5) = "May"
    ar_nmBulan(6) = "June"
    ar_nmBulan(7) = "July"
    ar_nmBulan(8) = "August"
    ar_nmBulan(9) = "September"
    ar_nmBulan(10) = "October"
    ar_nmBulan(11) = "November"
    ar_nmBulan(12) = "December"
    
    cmbBulan1.ListIndex = 0
    cmbBulan2.ListIndex = 0
    
    If Val(Format(Now, "MM")) >= 12 Then
        cmbBulan2.ListIndex = Val(Format(Now, "MM")) - 1
    Else
        cmbBulan2.ListIndex = Val(Format(Now, "MM"))
    End If
    
    Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbTahun1.Left = Label1.Left
    cmbTahun1.Width = Label1.Width
    cmbTahun1.Top = SkinLabel2.Top
    cmbBulan1.Left = Label2.Left
    cmbBulan1.Width = Label2.Width
    cmbBulan1.Top = SkinLabel2.Top
    cmbBulan2.Left = Label4.Left
    cmbBulan2.Width = Label4.Width
    cmbBulan2.Top = SkinLabel2.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        
    ElseIf KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText LTrim$(grid1.Clip)
        
    End If
End Sub

Private Sub grid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        
    ElseIf KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText LTrim$(grid2.Clip)
    End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 15
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lX As Integer, lY As Single
    If Button = vbLeftButton Then
        PicDet.Left = PicDet.Left + (x / 15 - lX)
        PicDet.Top = PicDet.Top + (Y / 15 - lY)
    Else
        lX = x / 15: lY = Y / 15
    End If
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 0
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
