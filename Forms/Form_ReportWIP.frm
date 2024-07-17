VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_ReportWIP 
   Caption         =   "Report of WIP"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   9330
   Begin VB.PictureBox PicLookLot 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   3360
      ScaleHeight     =   2895
      ScaleWidth      =   2505
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   2565
      Begin VB.ListBox ListLot 
         Height          =   2340
         Left            =   45
         TabIndex        =   8
         Top             =   480
         Width           =   2465
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Lot Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbFiletype 
      Height          =   405
      ItemData        =   "Form_ReportWIP.frx":0000
      Left            =   7440
      List            =   "Form_ReportWIP.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "View"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10186
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skinfd 
      Left            =   0
      OleObjectBlob   =   "Form_ReportWIP.frx":0023
      Top             =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Form_ReportWIP"
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
Dim i As Integer
Dim proportionArray() As CtrlProportion
Dim oExcel As Object
Dim oBook  As Object
Dim oSheet As Object
Dim rsLookup As Recordset


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

Private Sub settingCFlex()
    With grid1
        .row = 1
        .Rows = 2
        .Cols = 3
        .FixedCols = 0
        .TextMatrix(0, 0) = "No"
        .ColWidth(0) = 800
        .ColAlignment(0) = flexAlignLeftCenter
        .TextMatrix(0, 1) = "Part No"
        .ColWidth(1) = 2600
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Part Name"
        .ColWidth(2) = 2600
    End With
End Sub

Private Sub cmdExport_Click()
    Dim spreasheet      As String
    If cmbFiletype.ListIndex = 0 Then
        spreasheet = "Excel.Application"
    Else
        spreasheet = "et.Application"
    End If
    With grid1
        .Col = 0
        .row = 0
        .ColSel = .Cols - 1
        .RowSel = .Rows - 1
        Clipboard.SetText .Clip
        If .Rows > 1 Then
            CommonDialog1.Filter = ""
            CommonDialog1.ShowSave
            If CommonDialog1.FileName = "" Then Exit Sub
            If Err.Number = &H7FF3 Then MsgBox "Canceled": Exit Sub
            Set oExcel = CreateObject(spreasheet)
            Set oBook = oExcel.Workbooks.Add
            Set oSheet = oBook.Sheets.Item(1)
            oSheet.Cells(1, 1) = "Time Export : " & Now
            With oExcel.ActiveWorkbook.ActiveSheet
                .range("A2").Select 'Select Cell A1 (will paste from here, to different cells)
                .Paste              'Paste clipboard contents
            End With
            oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName ', xlWorkbookNormal
            MsgBox "saved !", vbInformation, "Creating Template"
            oExcel.Quit
            Set oSheet = Nothing
            Set oBook = Nothing
            Set oExcel = Nothing
        End If
    End With
End Sub

Private Sub cmdGen_Click()
    Dim qry As String
    Dim R As Long
    Dim ttlqty As Long
    qry = "select distinct partno,item_name from wip a inner join mst_item b on a.partno=b.item_id " _
        & " where coalesce(statuss,'')!='RFG' and coalesce(statuss,'')!='NG' group by partno, item_name"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        R = 1
        With grid1
            .Rows = 1
            .Rows = 1 + RsBantu.RecordCount
            RsBantu.Sort = "partno asc"
            While Not RsBantu.EOF
                .TextMatrix(R, 0) = R
                .TextMatrix(R, 1) = RsBantu("partno")
                .TextMatrix(R, 2) = RsBantu("item_name")
                R = R + 1
                RsBantu.MoveNext
            Wend
        End With
    End If
    
    qry = "select distinct locationn from wip "
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        R = 3
        With grid1
            .Cols = 3
            .Cols = 3 + RsBantu.RecordCount
            While Not RsBantu.EOF
                .TextMatrix(0, R) = RsBantu(0)
                R = R + 1
                RsBantu.MoveNext
            Wend
            .Cols = .Cols + 1
            .TextMatrix(0, .Cols - 1) = "Total"
        End With
    End If
    
    qry = "select partno,sum(qty) qty,locationn from wip a inner join mst_item b on a.partno=b.item_id " _
        & " where coalesce(statuss,'')!='RFG' and coalesce(statuss,'')!='NG' group by partno, locationn"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        With grid1
            For R = 1 To .Rows - 1
                For i = 3 To .Cols - 2
                    RsBantu.Filter = "partno='" & .TextMatrix(R, 1) & "' and locationn='" & .TextMatrix(0, i) & "'"
                    If RsBantu.RecordCount > 0 Then
                        .TextMatrix(R, i) = FormatNumber(RsBantu("qty"), 0)
                    Else
                        .TextMatrix(R, i) = 0
                    End If
                Next
                RsBantu.Filter = "partno='" & .TextMatrix(R, 1) & "'"
                ttlqty = 0
                If RsBantu.RecordCount > 0 Then
                    ttlqty = 0
                    While Not RsBantu.EOF
                        ttlqty = RsBantu("qty") + ttlqty
                        RsBantu.MoveNext
                    Wend
                    .TextMatrix(R, .Cols - 1) = FormatNumber(ttlqty, 0)
                Else
                    .TextMatrix(R, .Cols - 1) = FormatNumber(ttlqty, 0)
                End If
            Next
        End With
    End If
    qry = "select partno,sum(qty) qty,locationn,lotno from wip a inner join mst_item b on a.partno=b.item_id " _
        & " where coalesce(statuss,'')!='RFG' group by partno, locationn,lotno"
    Set rsLookup = New ADODB.Recordset
    Set rsLookup = Con.Execute(qry)
    Set RsBantu = Nothing
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
    Call TemaAktif(SkinFD, Me)
    settingCFlex
    AddTab Me
    Height = 7035
    Width = 9570
    cmbFiletype.ListIndex = 0
    grid1.Rows = 1
    Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbFiletype.Top = cmdExport.Top
    cmbFiletype.Left = Label1.Left
    cmbFiletype.Width = Label1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    WheelUnHook (Me.hWnd)
End Sub

Private Sub grid1_DblClick()
    With grid1
        If .Col > 2 And .Col < .Cols - 1 Then
            If .text * 1 > 0 Then
                rsLookup.Filter = "partno='" & .TextMatrix(.row, 1) & "' and locationn='" & .TextMatrix(0, .Col) & "'"
                ListLot.Clear
                If rsLookup.RecordCount > 0 Then
                    rsLookup.Sort = "lotno asc"
                    PicLookLot.Visible = True
                    While Not rsLookup.EOF
                        ListLot.AddItem rsLookup("lotno")
                        rsLookup.MoveNext
                    Wend
                End If
            End If
        End If
    End With
End Sub

Private Sub Label3_Click()
    PicLookLot.Visible = False
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos, 2
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
End Sub

