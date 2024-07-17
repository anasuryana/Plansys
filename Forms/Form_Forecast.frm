VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Forecast 
   Caption         =   "Forecast"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8970
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
   ScaleHeight     =   6240
   ScaleWidth      =   8970
   Begin VB.PictureBox PicHistory 
      BackColor       =   &H0080FF80&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   8715
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   8775
      Begin Planning_System.McCalendar kalender 
         Height          =   2535
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4471
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarHeight  =   151
         CalendarBackCol =   8454016
         MonthBackCol    =   12648384
         HeaderBackCol   =   8454016
         YearBackCol     =   12648384
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
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
         Left            =   8280
         TabIndex        =   29
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Revision Date"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   8295
      End
   End
   Begin MSComCtl2.DTPicker dtissue 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "mmmm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   22
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   145096707
      CurrentDate     =   43117
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   855
      Left            =   4440
      TabIndex        =   18
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox txtFind 
         Height          =   405
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "I/E Data"
      Height          =   1335
      Left            =   4440
      TabIndex        =   14
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox cmbFiletype 
         Height          =   405
         ItemData        =   "Form_Forecast.frx":0000
         Left            =   2520
         List            =   "Form_Forecast.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Import"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CustomerId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item Id"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Input FC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Period FC"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1200
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmbFind2 
      Caption         =   "..."
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtItemid 
      BackColor       =   &H00FFFFC0&
      Height          =   405
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "..."
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtCust 
      BackColor       =   &H00FFFFC0&
      Height          =   405
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_Forecast.frx":0023
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4920
      OleObjectBlob   =   "Form_Forecast.frx":008B
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_Forecast.frx":02BF
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblitemName 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "Form_Forecast.frx":0325
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_Forecast.frx":037D
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_Forecast.frx":03DB
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_Forecast.frx":0447
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtperiod 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "mmmm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   144834563
      CurrentDate     =   43117
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Revision"
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Form_Forecast"
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
Dim ar_cust() As String
Dim ar_custCode() As String
Dim ar_nmBulan(1 To 12) As String
Private posisisFind As Double
Dim oExcel As Object
Dim oBook  As Object
Dim oSheet As Object
Public cust_id As String
Dim dtissuev As Date
Dim dtperiodv As Date

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

Private Sub inAddValue(CustNamee As String, custCode As String)
    Dim r As Long
    If (Not ar_cust) <> -1 Then
        For r = 1 To UBound(ar_cust)
            If CustNamee = ar_cust(r) Then
            Exit Sub
            End If
        Next
        ReDim Preserve ar_cust(1 To UBound(ar_cust) + 1) As String
        ReDim Preserve ar_custCode(1 To UBound(ar_cust) + 1) As String
        ar_cust(UBound(ar_cust)) = CustNamee
        ar_custCode(UBound(ar_cust)) = custCode
    Else
        ReDim ar_cust(1 To 1) As String
        ReDim ar_custCode(1 To 1) As String
        ar_cust(UBound(ar_cust)) = CustNamee
        ar_custCode(UBound(ar_cust)) = custCode
    End If
End Sub

Private Sub Check1_Click()
    If Check1 = vbChecked Then
        PicHistory.Visible = True
    Else
        PicHistory.Visible = False
    End If
End Sub

Private Sub cmbFind2_Click()
    GetForm = Me.Name
    PopUp_Item_Sup.Show 1
    txtQty.SetFocus
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo) = vbYes Then
        Dim qry As String
        Dim raff As Long
        Dim period_h As String
        Dim period_fc As String
        period_h = Format(dtissue, "yyyyMM")
        period_fc = Format(dtperiod, "yyyyMM")
        qry = "delete from forecast_mod where item_id='" & txtItemid & "' and period='" & period_fc & "' " _
                & " and cust_id='" & cust_id & "' and period_h='" & period_h & "'"
        Con.Execute qry, raff
        MsgBox "Deleted"
        cmdNew_Click
        loadData
    End If
End Sub

Private Sub cmdfind_Click()
    GetForm = Me.Name
    popUp_Customer.Show 1
    cmbFind2.SetFocus
End Sub


Private Sub cmdNew_Click()
    cust_id = ""
    txtCust = ""
    txtItemid = ""
    lblitemname = ""
    txtQty = ""
    cmdSave.Tag = "s"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Ex
    Dim qry As String
    Dim period_h As String
    Dim uperiod_h As String
    Dim period_fc As String
    Dim uperiod_fc As String
    
    period_h = Format(dtissue, "yyyyMM")
    period_fc = Format(dtperiod, "yyyyMM")
    
    If IsNumeric(txtQty) = False Then txtQty.SetFocus: Exit Sub
    If txtItemid = "" Then txtItemid.SetFocus: Exit Sub
    If txtCust = "" Then txtCust.SetFocus: Exit Sub
    
   
    If cmdSave.Tag = "s" Then
        If MsgBox("Are you sure want to save ? ", vbQuestion + vbYesNo) = vbYes Then
            qry = "insert into forecast_mod values('" & txtItemid & "','" & period_fc & "'," _
            & txtQty * 1 & ",'" & cust_id & "','" & period_h & "')"
            Con.Execute qry
            MsgBox "Saved"
            loadData
        End If
    Else
        uperiod_h = Format(dtissuev, "yyyyMM")
        uperiod_fc = Format(dtperiodv, "yyyyMM")
        If MsgBox("Are you sure want to update ? ", vbQuestion + vbYesNo) = vbYes Then
            qry = "update forecast_mod set period_h='" & period_h & "',period='" & period_fc & "',qty =" & txtQty * 1 _
            & " where item_id='" & txtItemid & "' and period='" & uperiod_fc & "' " _
            & " and cust_id='" & cust_id & "' and period_h='" & uperiod_h & "'"
            Con.Execute qry
            MsgBox "Updated"
            loadData
        End If
    End If
    
    Exit Sub
Ex:
    MsgBox Err.Description
End Sub

Private Sub cmdUpload_Click()

    Dim urlFILE As String, ada As Boolean, fSO As FileSystemObject
    Dim xPart As String
    Dim xQty As Long
    Dim periodYM As String
    Dim RsA As ADODB.Recordset
    Dim kol As Byte
    Dim i As Long
    Dim qry As String
    Dim totalBaris As Long
    Const NamaTabel As String = "forecast_mod"
    With CommonDialog1
        .Filter = ""
        .ShowOpen
        urlFILE = .FileName
    End With
    If urlFILE <> "" Then
        Dim spreasheet      As String
        If cmbFiletype.ListIndex = 0 Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        Set fSO = New FileSystemObject
        If fSO.FileExists(urlFILE) = False Then Exit Sub
        Set oExcel = CreateObject(spreasheet)
        oExcel.Workbooks.Open urlFILE
        Set oBook = oExcel.Workbooks(1)
        Set oSheet = oBook.Worksheets(1)
        With oSheet
            totalBaris = .Range("A" & .rows.Count).End(-4162).Row
        End With
        BukaKoneksi
        i = 5
        ada = True
        While ada
            xPart = oSheet.Cells(i, 1)
            If xPart <> "" Then
                For kol = 2 To 13
                    If IsNumeric(oSheet.Cells(i, kol)) Then
                        xQty = Val(oSheet.Cells(i, kol))
                    Else
                        xQty = 0
                    End If
                    If kol - 1 > 9 Then
                        periodYM = Left(oSheet.Cells(2, 2), 4) & kol - 1
                    Else
                        periodYM = Left(oSheet.Cells(2, 2), 4) & "0" & kol - 1
                    End If
                    qry = "select qty from " & NamaTabel & " " _
                    & " where item_id='" & xPart & "' " _
                    & " and period='" & periodYM & "' " _
                    & " and cust_id='" & oSheet.Cells(1, 2) & "' and period_h='" & oSheet.Cells(2, 2) & "'"
                    Set RsA = New ADODB.Recordset

                    Set RsA = Con.Execute(qry)
                    If RsA.RecordCount > 0 Then
                        qry = "update " & NamaTabel & " set qty=" & xQty _
                        & " where item_id='" & xPart & "' " _
                        & " and period='" & periodYM & "' " _
                        & " and cust_id='" & oSheet.Cells(1, 2) & "' and period_h='" & oSheet.Cells(2, 2) & "'"
                        Con.Execute qry
                    Else
                        qry = "INSERT INTO " & NamaTabel & " values('" & xPart & "', " _
                        & "'" & periodYM & "'," & xQty & ",'" & oSheet.Cells(1, 2) & "','" & oSheet.Cells(2, 2) & "')"
                        Con.Execute qry
                    End If
                Next
                Set RsA = Nothing
            Else
                ada = False
            End If
            pb1.Value = ((i - 5) / (totalBaris - 4)) * 100
            i = i + 1
        Wend
        '====END NEW

        MsgBox "Uploaded !", vbInformation, "Upload Status"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    End If
End Sub



Private Function checkMonth(prNo As String, bln As String) As Boolean
    Dim rsC As Byte
    Dim hasil As Boolean
    hasil = False
    For rsC = 1 To UBound(ar_nmBulan)
        If rsC = Val(prNo) And bln = Left$(ar_nmBulan(rsC), 3) Then
            hasil = True
        End If
    Next
    checkMonth = hasil
End Function

Private Function nmMonthtoNumber(ptext As String) As String
    Dim ra As Byte
    For ra = 1 To UBound(ar_nmBulan)
        If Left$(ar_nmBulan(ra), 3) = ptext Then
            nmMonthtoNumber = CStr(ra)
            Exit For
        End If
    Next
End Function

Private Function nmCustmtoId(ptext As String) As String
    Dim ra As Byte
    For ra = 1 To UBound(ar_cust)
        If ar_cust(ra) = ptext Then
            nmCustmtoId = ar_custCode(ra)
            Exit For
        End If
    Next
End Function



Private Sub Command2_Click()
    txtQty.SetFocus
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub lvToForm()
    With lv1
        cust_id = .SelectedItem.Text
        txtCust.Text = .SelectedItem.SubItems(1)
        txtItemid = .SelectedItem.SubItems(2)
        lblitemname = .SelectedItem.SubItems(3)
        txtQty = .SelectedItem.SubItems(4)
        dtissue = buildDate(.SelectedItem.SubItems(5))
        dtissuev = buildDate(.SelectedItem.SubItems(5))
        
        dtperiod = buildDate(.SelectedItem.SubItems(6))
        dtperiodv = buildDate(.SelectedItem.SubItems(6))
        
    End With
End Sub

Private Function buildDate(thstring As String) As Date
    Dim tgl As Date
    tgl = DateSerial(Left(thstring, 4), CInt(Right(thstring, 2)), 1)
    buildDate = tgl
End Function


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

Private Function getLastdateInput() As Date
    Dim qry As String
    qry = "select max(inputtime)::date inputtime from forecast_mod"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        getLastdateInput = RsBantu(0)
    Else
        getLastdateInput = Now
    End If
End Function

Private Sub getDateRevision()
    Dim qry As String
    Dim tgl As Byte
    Dim bulan As Byte
    Dim tahun As String
    
    qry = "SELECT inputtime,count(*) FROM " _
    & " (SELECT item_id,period,count(*),inputtime::date FROM forecast_mod group by " _
    & " period,item_id,inputtime having count(*)>1) v1 " _
    & " where extract(YEAR from inputtime)=" & Format(Now, "yyyy") _
    & " GROUP BY inputtime order by 1 DESC"
    Set RsBantu = Con.Execute(qry)
    qry = ""
    If RsBantu.RecordCount > 0 Then
        While Not RsBantu.EOF
            tgl = CByte(Format(RsBantu("inputtime"), "dd"))
            bulan = CByte(Format(RsBantu("inputtime"), "MM"))
            tahun = Format(RsBantu("inputtime"), "yyyy")
            qry = qry & tgl & "-" & bulan & ">Revisi pada tahun " & tahun & ","
            RsBantu.MoveNext
        Wend
        qry = Left(qry, Len(qry) - 1)
        kalender.SpecialDays = qry
    Else
        kalender.SpecialDays = ""
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 9210
    Me.Height = 6810
    AddTab Me
    activeTheme Skin1, Me
    getDateRevision
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
    cmbFiletype.ListIndex = 0
    loadData
    cmdSave.Tag = "s"
    dtissue = getLastdateInput
    dtperiod = Now
    kalender.Value = Now
    
End Sub

Private Sub loadData()
    Dim qry As String
    Dim li As ListItem
    Dim i As Long
    qry = "Select a.cust_id,cust_name,a.item_id,item_name,qty " _
    & " ,period,period_h from forecast_mod a inner join r_customer b " _
    & " on a.cust_id=b.cust_id inner join mst_item c on a.item_id=c.item_id " _
    & " where qty>0" _
    & " order by period_h asc,period asc"
    Set RsBantu = Con.Execute(qry)
    lv1.ListItems.Clear
    While Not RsBantu.EOF
        i = i + 1
        Set li = lv1.ListItems.Add(, , RsBantu("cust_id"))
        li.SubItems(1) = RsBantu("cust_name")
        li.SubItems(2) = RsBantu("item_id")
        li.SubItems(3) = RsBantu("item_name")
        li.SubItems(4) = RsBantu("qty")
        li.SubItems(5) = RsBantu("period_h")
        li.SubItems(6) = RsBantu("period")
        RsBantu.MoveNext
    Wend
End Sub

Private Sub Form_Resize()
    ResizeControls
    cmbFiletype.Top = cmdUpload.Top
    cmbFiletype.Width = Label3.Width
    cmbFiletype.Left = Label3.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
    Call WheelUnHook(Me.hwnd)
End Sub

Sub SelectAllText(tb As TextBox)

tb.SelStart = 0
tb.SelLength = Len(tb.Text)

End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 15
End Sub


Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 0
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    PicHistory.Visible = False
    Check1.Value = vbUnchecked
End Sub


Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 0
End Sub

Private Sub lv1_Click()
    lvToForm
End Sub

Private Sub lv1_DblClick()
    cmdSave.Tag = "u"
    txtQty.SetFocus
End Sub

Private Sub lv1_KeyUp(KeyCode As Integer, Shift As Integer)
    lvToForm
End Sub

Private Sub loadKriteria()
    Dim qry As String
    Dim li As ListItem
    Dim i As Long
    qry = "Select a.cust_id,cust_name,a.item_id,item_name,qty " _
    & " ,period,period_h from forecast_mod a inner join r_customer b " _
    & " on a.cust_id=b.cust_id inner join mst_item c on a.item_id=c.item_id " _
    & " where lower(a.item_id) like '%" & LCase(txtfind) & "%'" _
    & " order by period_h asc,period asc"
    Set RsBantu = Con.Execute(qry)
    lv1.ListItems.Clear
    While Not RsBantu.EOF
        i = i + 1
        Set li = lv1.ListItems.Add(, , RsBantu("cust_id"))
        li.SubItems(1) = RsBantu("cust_name")
        li.SubItems(2) = RsBantu("item_id")
        li.SubItems(3) = RsBantu("item_name")
        li.SubItems(4) = RsBantu("qty")
        li.SubItems(5) = RsBantu("period_h")
        li.SubItems(6) = RsBantu("period")
        RsBantu.MoveNext
    Wend
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtfind = FilterIn(txtfind)
        loadKriteria
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


