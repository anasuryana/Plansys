VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_STE 
   Caption         =   "ENG Trial Schedule"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
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
   ScaleHeight     =   5865
   ScaleWidth      =   8715
   Begin VB.ComboBox cmbFiletype 
      Height          =   390
      ItemData        =   "Form_STE.frx":0000
      Left            =   3960
      List            =   "Form_STE.frx":000A
      TabIndex        =   19
      Top             =   1920
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pg 
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Import"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdFindMCH 
      Caption         =   "..."
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find"
      Height          =   735
      Left            =   5640
      TabIndex        =   12
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtFind 
         Height          =   390
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTstart 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy HH:mm"
      Format          =   287047683
      CurrentDate     =   42727
   End
   Begin VB.TextBox txtMachine 
      Height          =   390
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton CMDfind 
      Caption         =   "..."
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtPartNo 
      Height          =   390
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_STE.frx":0023
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid msflx 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double click to edit"
      Top             =   2400
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5953
      _Version        =   393216
      ForeColorSel    =   -2147483635
      TextStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_STE.frx":0089
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_STE.frx":02BD
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_STE.frx":0323
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTfinish 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy HH:mm"
      Format          =   287047683
      CurrentDate     =   42727
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   3960
      OleObjectBlob   =   "Form_STE.frx":0385
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblPg 
      Height          =   255
      Left            =   6960
      OleObjectBlob   =   "Form_STE.frx":03E1
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   3960
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Form_STE"
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
Dim i As Double
Dim qry As String
Dim idT As String
Private oExcel      As Object
Private oBook       As Object
Private oSheet      As Object

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

Private Sub cmdDel_Click()
    If Len(idT) > 0 Then
        qry = "delete from mpp_ste where idste=" & idT
        Con.Execute qry
        MsgBox "Deleted successfully"
        txtPartNo = ""
        txtMachine = ""
        Call LoadDataSQL
        Call LoadDatanya_V2
    End If
End Sub

Private Sub cmdExport_Click()
On Error GoTo duhH
    If msflx.rows < 2 Then MsgBox "nothing to be exported": Exit Sub
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
        Dim k As Integer
        With oSheet
            .Cells(1, 1) = "ENG Trial Schedule"
            .Cells(2, 1) = "Date : " & Now
            .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
            .Columns(5).NumberFormat = "@"
            .Columns(6).NumberFormat = "@"
            
        End With
        pg.Visible = True
        pg.Value = 0
        lblPg.Visible = True
        With msflx
            For i = 0 To .rows - 1
                For k = 0 To .Cols - 1
                    If i = 0 Then
                        oSheet.Cells(i + 3, k + 1).Font.Bold = True
                        oSheet.Cells(i + 3, k + 1) = .TextMatrix(i, k)
                    Else
                        oSheet.Cells(i + 3, k + 1) = .TextMatrix(i, k)
                    End If
                    If k = 0 Then
                        If i = 0 Then
                            oSheet.Cells(i + 3, k + 1) = "No"
                        Else
                            oSheet.Cells(i + 3, k + 1) = i
                        End If
                    End If
                Next
                pg.Value = FormatNumber(((i + 1) * 100) / .rows, 0)
                lblPg.Caption = pg.Value & "%"
            Next
        End With
        oSheet.Columns("B:F").AutoFit
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
        pg.Visible = False
        lblPg.Visible = False
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
    Exit Sub
duhH:
    oExcel.Quit
    Set oSheet = Nothing
    Set oBook = Nothing
    Set oExcel = Nothing
End Sub

Private Sub cmdfind_Click()
    GetForm = Me.Name
    PopUp_Item_Sup.Show 1
End Sub

Private Sub cmdFindMCH_Click()
    GetForm = Me.Name
    PopUp_machine.Show 1
End Sub

Private Sub cmdSave_Click()
On Error GoTo Duh
    If Len(txtPartNo) = 0 Then txtPartNo.SetFocus: Exit Sub
    If Len(txtMachine) = 0 Then txtMachine.SetFocus: Exit Sub
    If cmdSave.Tag = "s" Then
        qry = "insert into mpp_ste (idste,part_no,mch,date_trial,date_trialf) values " _
            & "(DEFAULT,'" & txtPartNo & "','" & txtMachine & "','" & Format(DTstart, "yyyy-MM-dd HH:mm") & "','" & Format(DTfinish, "yyyy-MM-dd HH:mm") & "')"
        Con.Execute qry
        MsgBox "Saved successfully", vbInformation, "Good"
    Else
        qry = "update mpp_ste set part_no='" & txtPartNo & "',mch='" & txtMachine & "',date_trial='" & Format(DTstart, "yyyy-MM-dd HH:mm") & "',date_trialf='" & Format(DTstart, "yyyy-MM-dd HH:mm") & "'" _
            & " where idste=" & idT
        Con.Execute qry
        MsgBox "Updated successfully", vbInformation, "Good"
        cmdSave.Tag = "s"
    End If
    LoadDataSQL
    LoadDatanya_V2
    Exit Sub
Duh:
    MsgBox Err.Description, vbCritical, "Sorry [" & Err.Number & "]"
End Sub

Private Sub cmdUpload_Click()
'On Error GoTo Duh
    Dim urlFILE As String, ada As Boolean, barisX As Double
    Const NamaTabel As String = "mpp_ste"
    Const FormatTGLJAM As String = "yyyy-MM-dd HH:mm"
    Const FormatJAM As String = "HH:mm"
    With CommonDialog1
        .Filter = ""
        .ShowOpen
        urlFILE = .FileName
    End With
    If urlFILE <> "" Then
        Set oExcel = New Excel.Application
        oExcel.Workbooks.Open urlFILE
        Set oBook = oExcel.Workbooks(1)
        Set oSheet = oBook.Worksheets(1)
        ada = True
        barisX = 4
        BukaKoneksi
        Screen.MousePointer = 11
        pg.Visible = False
        While ada
            If oSheet.Cells(barisX, 2) <> "" Then
                qry = "INSERT INTO " & NamaTabel & " (idste,part_no,mch,date_trial,date_trialf) values(DEFAULT,'" & oSheet.Cells(barisX, 2) & "'" _
                        & ",'" & oSheet.Cells(barisX, 4) & "'" _
                        & ",'" & Format(oSheet.Cells(barisX, 5), FormatTGLJAM) & "'" _
                        & ",'" & Format(oSheet.Cells(barisX, 6), FormatTGLJAM) & "')"
                Con.Execute qry
                txtPartNo = oSheet.Cells(barisX, 2)
                txtMachine = oSheet.Cells(barisX, 4)
            Else
                ada = False
            End If
            barisX = 1 + barisX
            lblPg.Caption = barisX - 3 & " row(s) saved"
        Wend
        Screen.MousePointer = 0
        MsgBox "Uploaded !", vbInformation, "Upload Status"
        
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    End If
    Exit Sub
'Duh:
'    MsgBox Err.Description, vbInformation, "Sorry [" & Err.Number & "]"
'    oExcel.Quit
'        Set oSheet = Nothing
'        Set oBook = Nothing
'        Set oExcel = Nothing
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

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbFiletype.Left = Label1.Left
    cmbFiletype.Top = cmdDel.Top
    cmbFiletype.Width = Label1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub settingFG()
    With msflx
        .Cols = 6
        .rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        
        i = 0
        .TextMatrix(0, i) = "IDkok"
        .ColWidth(i) = 0
        
        i = 1
        .TextMatrix(0, i) = "Part No"
        .ColWidth(i) = 2600
        
        i = 2
        .TextMatrix(0, i) = "Part Name"
        .ColWidth(i) = 2700
        
        i = 3
        .TextMatrix(0, i) = "Machine"
        
        i = 4
        .TextMatrix(0, i) = "Start Date"
        .ColWidth(i) = 2700
        
        i = 5
        .TextMatrix(0, i) = "Finish Date"
        .ColWidth(i) = 2700
        
        
    End With
End Sub

Private Sub LoadDataSQL()
    qry = "SELECT idste,part_no,partname,mch,date_trial,date_trialf FROM mpp_ste a " _
    & " left join loadcap_mst_product_r b on a.part_no=b.partno "
    Set RsGet = Con.Execute(qry)
End Sub

Private Sub getList()
    msflx.rows = 1 + RsGet.RecordCount
    i = 1
    With msflx
        Do Until RsGet.EOF
            .TextMatrix(i, 0) = RTrim(RsGet!idste)
            .TextMatrix(i, 1) = RTrim(RsGet!part_no)
            .TextMatrix(i, 2) = RTrim(RsGet!partname)
            .TextMatrix(i, 3) = IIf(IsNull(RsGet!mch), 0, RsGet!mch)
            .TextMatrix(i, 4) = Format(RsGet!date_trial, "yyyy-MM-dd HH:mm")
            .TextMatrix(i, 5) = IIf(IsNull(RsGet!date_trialf), 0, Format(RsGet!date_trialf, "yyyy-MM-dd HH:mm"))
            i = 1 + i
            RsGet.MoveNext
        Loop
    End With
End Sub

Private Sub LoadDatanya_V2()
    If Len(Trim(txtfind)) > 0 Then
        RsGet.Fields("part_no").Properties("Optimize") = True
        RsGet.Fields("partname").Properties("Optimize") = True
        RsGet.Filter = "part_no LIKE '*" & txtfind & "*'"
        If RsGet.RecordCount > 0 Then
            Call getList
        Else
            RsGet.Filter = adFilterNone
            RsGet.Filter = "partname LIKE '*" & txtfind & "*'"
            Call getList
        End If
    Else
        RsGet.Filter = adFilterNone
        Call getList
    End If
End Sub

Private Sub lToForm()
    With msflx
        idT = .TextMatrix(.Row, 0)
        txtPartNo = .TextMatrix(.Row, 1)
        txtMachine = .TextMatrix(.Row, 3)
        DTstart = .TextMatrix(.Row, 4)
        DTfinish = .TextMatrix(.Row, 5)
    End With
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingFG
    Call LoadDataSQL
    Call LoadDatanya_V2
    Me.Width = 8955
    Me.Height = 6435
    cmdSave.Tag = "s"
    DTstart = Now
    DTfinish = Now
    cmbFiletype.ListIndex = 0
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If

End Sub

Private Sub msflx_Click()
    Call lToForm
End Sub

Private Sub msflx_DblClick()
    cmdSave.Tag = "u"
End Sub

Private Sub msflx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText msflx.Clip
    End If
End Sub

Private Sub msflx_KeyUp(KeyCode As Integer, Shift As Integer)
    Call lToForm
    If KeyCode = 46 Then
        cmdDel_Click
    End If
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LoadDataSQL
        Call LoadDatanya_V2
    End If
End Sub
