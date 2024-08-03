VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_ReprintWO 
   Caption         =   "Reprint WO"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   7680
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicExportContainer 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   10875
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   10935
      Begin VB.OptionButton OptExportWPS 
         Caption         =   "WPS"
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton optExportMicrosoft 
         Caption         =   "Microsoft Excel"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4200
         TabIndex        =   22
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Exporter"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   10335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10320
         TabIndex        =   23
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdTriggerModalExport 
      Caption         =   "Export to ..."
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
      Left            =   9840
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   10875
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton Command1 
         Caption         =   "Synchronize"
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
         Left            =   4200
         TabIndex        =   16
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblsynch 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   19
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10320
         TabIndex        =   18
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Synchronization BOM Data"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   10335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
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
      Left            =   8160
      TabIndex        =   14
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdToday 
      Caption         =   "Today's Printed WO"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   720
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picTempRot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdPrintPrev 
      Caption         =   "Print Preview"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
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
      Left            =   7200
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtFind 
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
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   351928323
      CurrentDate     =   42753
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReprintWO.frx":0000
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5370
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid agrid 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7223
      _Version        =   393216
      Appearance      =   0
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
      Left            =   4200
      OleObjectBlob   =   "F_ReprintWO.frx":0060
      Top             =   840
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   350814211
      CurrentDate     =   42753
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   2520
      OleObjectBlob   =   "F_ReprintWO.frx":0294
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReprintWO.frx":02F0
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   7080
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "F_ReprintWO"
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
Private rsBOM As ADODB.Recordset
Private rsWO As ADODB.Recordset
Private rsUpdateCT As ADODB.Recordset
Private timeupdate As String
Private manPower As Single
Private needDay As Single
Private reqQTy As Single
Private reqQTy_purg As Single
Private ttl_reqQty As Single
Private rsAneHelper As ADODB.Recordset
Private oExcel      As Object 'Excel.Application
Private oBook       As Object 'Excel.Workbook
Private oSheet      As Object 'Excel.Worksheet

Private k As Byte
Dim i As Long
Dim r As Byte
Dim qry As String
Dim cvt_act As String
Private spreasheet  As String

Public Sub RotatePicture(fr_pic As PictureBox, to_pic As PictureBox, ByVal angle As Integer)
Dim fr_pixels() As RGBTriplet
Dim to_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim fr_wid As Long
Dim fr_hgt As Long
Dim to_wid As Long
Dim to_hgt As Long
Dim x As Integer
Dim Y As Integer

    ' Get the picture's image.
    GetBitmapPixels fr_pic, fr_pixels, bits_per_pixel

    ' Get the picture's size.
    fr_wid = UBound(fr_pixels, 1) + 1
    fr_hgt = UBound(fr_pixels, 2) + 1
    If angle = 0 Or angle = 180 Then
        to_wid = fr_wid
        to_hgt = fr_hgt
    Else
        to_wid = fr_hgt
        to_hgt = fr_wid
    End If

    ' Size the output picture to fit.
    to_pic.Width = to_pic.Parent.ScaleX(to_wid, vbPixels, to_pic.Parent.ScaleMode) + _
        to_pic.Width - to_pic.ScaleWidth
    to_pic.Height = to_pic.Parent.ScaleY(to_hgt, vbPixels, to_pic.Parent.ScaleMode) + _
        to_pic.Height - to_pic.ScaleHeight

    ' Copy the rotated pixels.
    ReDim to_pixels(0 To to_wid - 1, 0 To to_hgt - 1)
    Select Case angle
        Case 0
            For x = 0 To fr_wid - 1
                For Y = 0 To fr_hgt - 1
                    to_pixels(x, Y) = fr_pixels(x, Y)
                Next Y
            Next x
        Case 90
            For x = 0 To fr_wid - 1
                For Y = 0 To fr_hgt - 1
                    to_pixels(to_wid - Y - 1, x) = fr_pixels(x, Y)
                Next Y
            Next x
        Case 180
            For x = 0 To fr_wid - 1
                For Y = 0 To fr_hgt - 1
                    to_pixels(to_wid - x - 1, to_hgt - Y - 1) = fr_pixels(x, Y)
                Next Y
            Next x
        Case 270
            For x = 0 To fr_wid - 1
                For Y = 0 To fr_hgt - 1
                    to_pixels(Y, to_hgt - x - 1) = fr_pixels(x, Y)
                Next Y
            Next x
        Case Else
            Stop
    End Select

    ' Display the result.
    SetBitmapPixels to_pic, bits_per_pixel, to_pixels

    ' Make the image permanent.
    to_pic.Refresh
    to_pic.Picture = to_pic.Image
End Sub

Private Sub agrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText agrid.Clip
        StatusBar1.Panels(1).Text = "Last Info: Copied"
    End If
End Sub

Private Sub cmdfind_Click()
    LoadDatanya
    If Len(txtfind) > 0 Then
        rsWO.Fields("wo_no").Properties("Optimize") = True
        rsWO.Fields("partno").Properties("Optimize") = True
        rsWO.Fields("mesinno").Properties("Optimize") = True
        rsWO.Fields("partname").Properties("Optimize") = True
        rsWO.Filter = "wo_no like '*" & txtfind & "*'"
        'rsWO.Sort = ""
        If rsWO.RecordCount = 0 Then
            rsWO.Filter = adFilterNone
            rsWO.Filter = "partno like '*" & txtfind & "*'"
            If rsWO.RecordCount = 0 Then
                rsWO.Filter = adFilterNone
                rsWO.Filter = "mesinno like '*" & txtfind & "*'"
                If rsWO.RecordCount = 0 Then
                    rsWO.Filter = "partname like '*" & txtfind & "*'"
                End If
            End If
        End If
        StatusBar1.Panels(1).Text = "Last Info: Filtered"
    Else
        rsWO.Filter = adFilterNone
    End If
    getList
End Sub

Private Sub getList()
    Dim tempS As String
    Dim no As Long
    Dim posRow As Long
    With agrid
        .rows = 1
        If rsWO.RecordCount > 0 Then
            .rows = 1
            For i = 1 To rsWO.RecordCount
                rsWO.AbsolutePosition = i
                If tempS <> rsWO("wo_no") Then
                    tempS = rsWO("wo_no")
                    .rows = .rows + 1
                    posRow = .rows - 1
                    no = no + 1
                    .TextMatrix(posRow, 0) = no
                    .TextMatrix(posRow, 1) = rsWO("wo_no")
                    .TextMatrix(posRow, 2) = rsWO("lotno")
                    .TextMatrix(posRow, 3) = rsWO("mesinno")
                    .TextMatrix(posRow, 4) = rsWO("moldno")
                    .TextMatrix(posRow, 5) = rsWO("partno")
                    .TextMatrix(posRow, 6) = rsWO("partname")
                    .TextMatrix(posRow, 7) = rsWO("qty")
                    .TextMatrix(posRow, 8) = rsWO("mpp_doc")
                    .TextMatrix(posRow, 9) = rsWO("mpprev")
                    .TextMatrix(posRow, 10) = rsWO("issudate")
                    .TextMatrix(posRow, 11) = rsWO("printdate")
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdprint_Click()
    PopUp_PrinterPrint.Show 1
    With agrid
        If .rows > 1 Then
            rsWO.Filter = adFilterNone
            rsWO.Filter = "wo_no='" & .TextMatrix(.Row, .Col) & "'"
            If rsWO.RecordCount > 0 Then
                getWO True
            End If
        End If
    End With
End Sub

Private Sub cmdToday_Click()
    LoadDatanya_V2
    If Len(txtfind) > 0 Then
        rsWO.Fields("wo_no").Properties("Optimize") = True
        rsWO.Fields("partno").Properties("Optimize") = True
        rsWO.Fields("mesinno").Properties("Optimize") = True
        rsWO.Fields("partname").Properties("Optimize") = True
        rsWO.Filter = "wo_no like '*" & txtfind & "*'"
        
        If rsWO.RecordCount = 0 Then
            rsWO.Filter = adFilterNone
            rsWO.Filter = "partno like '*" & txtfind & "*'"
            If rsWO.RecordCount = 0 Then
                rsWO.Filter = adFilterNone
                rsWO.Filter = "mesinno like '*" & txtfind & "*'"
                If rsWO.RecordCount = 0 Then
                    rsWO.Filter = "partname like '*" & txtfind & "*'"
                End If
            End If
        End If
        StatusBar1.Panels(1).Text = "Last Info: Filtered"
    Else
        rsWO.Filter = adFilterNone
    End If
    getList
End Sub



Private Sub cmdTriggerModalExport_Click()
    PicExportContainer.Visible = True
End Sub

Private Sub Command1_Click()
    'TEST FOR TOMOROOW
    Dim iSynchronized As Byte
    qry = "select a.wo_no,partno,b.wo_no womat from worko a left join worko_mat b " _
    & " on a.wo_no=b.wo_no where b.wo_no is null"
    Set RsGet = Con.Execute(qry)
    If RsGet.RecordCount > 0 Then
        While Not RsGet.EOF
            qry = "SELECT bom_com_item,bom_qty_perassy,item_name,pfm_id FROM mst_bom a inner join mst_item b on a.bom_com_item=b.item_id " _
            & " WHERE bom_par_item='" & RsGet("partno") & "'"
            Set RsBantu = Con.Execute(qry)
            If RsBantu.RecordCount > 0 Then
                While Not RsBantu.EOF
                    qry = "INSERT INTO worko_mat VALUES ('" & RsGet("wo_no") & "','" & Trim(RsBantu("bom_com_item")) & "','" & Trim(RsBantu("item_name")) _
                    & "','" & Trim(RsBantu("pfm_id")) & "'," & RsBantu("bom_qty_perassy") & ")"
                    Con.Execute qry
                    iSynchronized = iSynchronized + 1
                    RsBantu.MoveNext
                Wend
            End If
        Wend
    End If
End Sub

Private Sub Command2_Click()
    Picture1.Visible = True
End Sub

Private Sub Command3_Click()
On Error GoTo exCe
    Screen.MousePointer = 11
    qry = "SELECT a.wo_no,status,lotno,issudate,a.partno,partname,moldno,mesinno, " _
    & " a.qty" _
    & " FROM worko a inner join loadcap_mst_product_r b on a.partno=b.partno" _
    & " where (issudate>='" & Format(dt1, "yyyy-MM-dd") & "' and issudate<='" & Format(dt2, "yyyy-MM-dd") & "') and lotno<>'' order by wo_no asc"

   
    Set rsAneHelper = Con.Execute(qry)
    If rsAneHelper.RecordCount < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.CancelError = True
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        If optExportMicrosoft.Enabled = True Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        Set oExcel = CreateObject(spreasheet) 'New Excel.Application
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        oSheet.Cells(1, 1) = "Template Upload Production Schedule"
        oSheet.Cells(2, 1) = "No"
        oSheet.Cells(2, 2) = "NO WO"
        oSheet.Cells(2, 3) = "PERIOD"
        oSheet.Cells(2, 4) = "MACHINE NO"
        oSheet.Cells(2, 5) = "LOTNO"
        oSheet.Cells(2, 6) = "MOLD ID"
        oSheet.Cells(2, 7) = "WP DATE"
        oSheet.Cells(2, 8) = "PRODUCT NO"
        oSheet.Cells(2, 9) = "QTY"
        oSheet.Columns(9).NumberFormat = "@"
        
        Dim i As Integer, baris As Double
        baris = 3
        ProgressBar1.Visible = True
        ProgressBar1.Value = 0
        Dim totalRows As Double
        totalRows = rsAneHelper.RecordCount
        While Not rsAneHelper.EOF
            oSheet.Cells(baris, 1) = baris - 2
            oSheet.Cells(baris, 2) = rsAneHelper("wo_no")
            oSheet.Cells(baris, 3) = "20" & Right(rsAneHelper("wo_no"), 2) & Left(Right(rsAneHelper("wo_no"), 5), 2)
            oSheet.Cells(baris, 4) = IIf(IsNull(rsAneHelper("mesinno")), "-", rsAneHelper("mesinno"))
            oSheet.Cells(baris, 5) = rsAneHelper("lotno")
            oSheet.Cells(baris, 6) = rsAneHelper("moldno")
            oSheet.Cells(baris, 7).NumberFormat = "@"
            oSheet.Cells(baris, 7) = Format(rsAneHelper("issudate"), "yyyy-MM-dd")
            oSheet.Cells(baris, 8) = rsAneHelper("partno")
            oSheet.Cells(baris, 9) = rsAneHelper("qty")
           
            baris = baris + 1
            ProgressBar1.Value = FormatNumber(((baris - 3) * 100) / totalRows, 0)
            rsAneHelper.MoveNext
        Wend
        'xlWorkbookNormal
        oSheet.Columns("B:I").AutoFit
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, -4143
        MsgBox "saved !", vbInformation, "Good"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
        ProgressBar1.Visible = False
        
    Else
        MsgBox "Canceled !", vbInformation, "Sorry..."
    End If
    Screen.MousePointer = 0
    Exit Sub
exCe:
    MsgBox Err.Description, vbInformation, Err.Number
    Screen.MousePointer = 0
    MsgBox "Silahkan coba lagi"
End Sub

Private Sub dt1_Change()
    If dt1.Value > dt2.Value Then
        dt2.Value = dt1.Value
    End If
End Sub

Private Sub dt2_Change()
    If dt2.Value < dt1.Value Then
        dt1.Value = dt2.Value
    End If
End Sub

Private Sub Form_Initialize()
    WindowState = vbNormal
    
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

Private Sub settingFG()
    With agrid
        .Cols = 12
        .rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .ColAlignment(2) = flexAlignLeftCenter
        
        i = 0
        .TextMatrix(0, i) = "No"
        .ColWidth(i) = 700
        .ColAlignment(i) = flexAlignLeftCenter
        
        
        i = 1
        .TextMatrix(0, i) = "WO"
        .ColWidth(i) = 1800
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 2
        .TextMatrix(0, i) = "Lot"
        .ColWidth(i) = 1000
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 3
        .TextMatrix(0, i) = "Machine"
        .ColWidth(i) = 1300
        
        i = 4
        .TextMatrix(0, i) = "Mold"
        .ColWidth(i) = 3000
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 5
        .TextMatrix(0, i) = "Part No"
        .ColWidth(i) = 3000
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 6
        .TextMatrix(0, i) = "Part Name"
        .ColWidth(i) = 3000
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 7
        .TextMatrix(0, i) = "Qty"
        .ColWidth(i) = 3000
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 8
        .TextMatrix(0, i) = "MPP Doc"
        .ColWidth(i) = 2500
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 9
        .TextMatrix(0, i) = "MPP Doc Rev"
        .ColWidth(i) = 1500
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 10
        .TextMatrix(0, i) = "Issue date"
        .ColWidth(i) = 1500
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 11
        .TextMatrix(0, i) = "Print Time"
        .ColWidth(i) = 2500
        .ColAlignment(i) = flexAlignLeftCenter
        
    End With
End Sub

Private Sub getTimeLastUpdate(mppdoc As String, rev As String, partNo As String, mold As String, machine As String)
    'Awalnya prosedur ini dibuat hanya untuk mendapatkan perubahan terakhir pada kolom C/T
    qry = "select sum(planqty) ttlmpp,aa.cap_p_day, timeupdate,mpower,ml_hkw,bb.cav from mpp_gen aa inner join mpp_gen_d bb " _
        & " on aa.lcd_itemdid = bb.lcd_itemdid and aa.no_mach=bb.no_mach and aa.reg_mold=bb.reg_mold and aa.ml_doc=bb.fltpp_doc and aa.ml_rev=bb.fltpp_rev " _
        & " where aa.mpp_doc_no='" & mppdoc & "' and aa.mpp_revisi='" & rev & "' and aa.lcd_itemdid = '" & partNo & "' and aa.no_mach = '" & machine & "' and aa.reg_mold = '" & mold & "' " _
        & " group by aa.lcd_itemdid, aa.no_mach,aa.reg_mold,timeupdate,aa.cap_p_day,mpower,ml_hkw,bb.cav "
        
    Set rsUpdateCT = Con.Execute(qry)
    If rsUpdateCT.RecordCount > 0 Then
        cvt_act = rsUpdateCT("cav")
        timeupdate = Format(rsUpdateCT("timeupdate"), "dd MMM yyyy HH:mm")
        needDay = rsUpdateCT("ttlmpp") / rsUpdateCT("cap_p_day")
        manPower = rsUpdateCT("mpower") 'needDay / rsUpdateCT("ml_hkw") * rsUpdateCT("mpower")
    End If
End Sub

Private Sub cmdPrintPrev_Click()
On Error GoTo Ex
    With agrid
        If .rows > 1 Then
            rsWO.Filter = adFilterNone
            rsWO.Filter = "wo_no='" & .TextMatrix(.Row, .Col) & "'"
            If rsWO.RecordCount > 0 Then
                getTimeLastUpdate .TextMatrix(.Row, 7), .TextMatrix(.Row, 8), .TextMatrix(.Row, 5), .TextMatrix(.Row, 4), .TextMatrix(.Row, 3)

                getWO False
            End If
        End If
    End With
    Exit Sub
Ex:
    MsgBox "error: print preview" & Err.Description
End Sub

Private Sub LoadDatanya()
    qry = "SELECT a.wo_no,status,lotno,issudate,a.partno,partname,moldno,mesinno, " _
    & " a.qty,c.qty::varchar qty_mat ,cavstd,ctscnd,ctmachine,targetpshift,manpower,leadtime,datesupply,a.isno," _
    & " tipelabel,mpp_doc,mpprev,c.item_id,c.item_nm,c.item_type,um_name,coalesce(qty_prg,0) qty_prg, printdate," _
    & " coalesce(colordesc,'-') colordesc" _
    & " FROM worko a inner join loadcap_mst_product_r b on a.partno=b.partno" _
    & " left join worko_mat c on a.wo_no=c.wo_no " _
    & " left join mst_item d on c.item_id=d.item_id" _
    & " left join r_unit_measure e on d.um_id=e.um_id" _
    & " where (issudate>='" & Format(dt1, "yyyy-MM-dd") & "' and issudate<='" & Format(dt2, "yyyy-MM-dd") & "') and coalesce(lotno,'')<>'' order by wo_no asc"

    Set rsWO = Con.Execute(qry)
    StatusBar1.Panels(1).Text = "Last Info: Loaded"
End Sub

Private Sub LoadDatanya_V2()
    qry = "SELECT a.wo_no,status,lotno,issudate,a.partno,partname,moldno,mesinno, " _
    & " a.qty,c.qty::varchar qty_mat ,cavstd,ctscnd,ctmachine,targetpshift,manpower,leadtime,datesupply,a.isno," _
    & " tipelabel,mpp_doc,mpprev,c.item_id,c.item_nm,c.item_type,um_name,coalesce(qty_prg,0) qty_prg, printdate" _
    & ",coalesce(colordesc,'-') colordesc FROM worko a inner join loadcap_mst_product_r b on a.partno=b.partno" _
    & " inner join worko_mat c on a.wo_no=c.wo_no " _
    & " inner join mst_item d on c.item_id=d.item_id" _
    & " inner join r_unit_measure e on d.um_id=e.um_id" _
    & " where printdate::date='" & Format(Now, "yyyy-MM-dd") & "' and lotno<>'' ORDER BY wo_no asc"
    Set rsWO = Con.Execute(qry)
    StatusBar1.Panels(1).Text = "Last Info: Loaded"
End Sub



Private Function GenerateCode128(Str As String, Optional BarWidth As Integer = 1) As Single
    Dim Code128 As New clsCode128
    Dim BarCodeWidth As Long
    Dim angle As Integer
    angle = 90
    
    picTemp.Cls
    picTemp.Width = 1
    picTemp.Picture = LoadPicture()
    BarCodeWidth = Code128.Code128_Print(Str, picTemp, BarWidth, True)
    picTemp.Picture = picTemp.Image
    SavePicture picTemp.Picture, App.Path & "\Templates\com.bmp"
    
    picTemp.Cls
    picTemp.Picture = LoadPicture()
    picTemp.Picture = LoadPicture(App.Path & "\Templates\com.bmp")
    picTemp.Picture = picTemp.Image
    RotatePicture picTemp, picTempRot, angle
    picTempRot.Picture = picTempRot.Image
    SavePicture picTempRot.Picture, App.Path & "\Templates\comr.bmp"

    GenerateCode128 = Me.CurrentY
End Function

Private Sub getWO(pPrintPP As Boolean)
On Error GoTo Exc
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim q As Integer
    Dim QRYane As String
    Dim sqlRep As String
    Dim a As Single
    
    QRYane = "select item_id from mst_item limit 1"
    sqlRep = "SHAPE {" & QRYane & "} as CMD1 "
    
    cn.Open "PROVIDER=MSDataShape;DSN=" & GetINI("SETTING", "odbc", vbNullString) & ";"

    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlRep
        .Execute
    End With

    With rs
        .ActiveConnection = cn
        .CursorLocation = adUseClient
        .Open cmd
    End With
    Printer.PaperSize = vbPRPSA3
    Printer.Orientation = vbPRORLandscape
    With Report_MPS
        Set .DataSource = rs
        .Orientation = rptOrientLandscape
        .DataMember = ""
        'getBOMbyPar rsWO("lcd_itemdid")
        Dim yy As Single
        yy = GenerateCode128(rsWO("wo_no"))
        .Sections("Section4").Controls("lblheaderpart").Caption = rsWO("partno")
        With .Sections("Section2")
            .Controls("lblIssuedDate").Caption = Format(rsWO("issudate"), "dd-MMM-yyyy")
            .Controls("lblwarna").Caption = rsWO("colordesc")
            .Controls("lbl_k_issued").Caption = .Controls("lblIssuedDate").Caption
            .Controls("lblPartNo").Caption = rsWO("partno")
            .Controls("lbl_k_partno").Caption = .Controls("lblPartNo").Caption
            If Len(.Controls("lbl_k_partno").Caption) > 7 Then
                .Controls("lbl_k_partno").Font.Size = 8
            End If
            .Controls("lblPartNM").Caption = rsWO("partname")
            .Controls("lbl_k_partnm").Caption = .Controls("lblPartNM").Caption
            .Controls("lblMoldNo").Caption = rsWO("moldno")
            .Controls("lblMesinNo").Caption = rsWO("mesinno")
            .Controls("lbl_k_mesin").Caption = .Controls("lblMesinNo").Caption
            .Controls("lblLOT").Caption = Left(rsWO("lotno"), 2) & " " & Mid(rsWO("lotno"), 3, 2) & " " & Right(rsWO("lotno"), 2)
            .Controls("lbl_k_lot").Caption = .Controls("lblLOT").Caption
            .Controls("lblQTY").Caption = rsWO("qty")
            .Controls("lblcavitystd").Caption = rsWO("cavstd")
            .Controls("lblctfinishing").Caption = rsWO("ctscnd")
            .Controls("lblctmachine").Caption = rsWO("ctmachine")
            .Controls("lbltargetshift").Caption = rsWO("targetpshift")
            .Controls("lblmanpower").Caption = ceiling(CDbl(manPower))
            .Controls("lblLeadTime").Caption = rsWO("leadtime")
            .Controls("lbldatesupply").Caption = Format(rsWO("datesupply"), "dd-MMM-yyyy")
            .Controls("lblISno").Caption = rsWO("isno")
            .Controls("lblmanual").Caption = rsWO("tipelabel")
            .Controls("lblNodoc").Caption = rsWO("wo_no")
            .Controls("lbl_k_nodoc").Caption = .Controls("lblNodoc").Caption
            .Controls("lbltimeupdate").Caption = timeupdate
            .Controls("lblcavityact").Caption = cvt_act
            Set .Controls("Image1").Picture = LoadPicture(App.Path & "\Templates\com.bmp")
            Set .Controls("Image3").Picture = LoadPicture(App.Path & "\Templates\comr.bmp")
                                  
            .Controls("lbl_k_matid").Caption = ""
            .Controls("lbl_m_matid").Caption = ""
            .Controls("lbl_k_matvir_nm").Caption = ""
            .Controls("lblqtyReq").Caption = ""
            .Controls("lblqtyReq_m").Caption = ""
            reqQTy = 0
            'reinit_material
            ttl_reqQty = 0
            For r = 1 To rsWO.RecordCount
                rsWO.AbsolutePosition = r
                .Controls("lbl_k_matid").Caption = .Controls("lbl_k_matid").Caption & vbNewLine & rsWO("item_id")
                .Controls("lbl_k_matvir_nm").Caption = .Controls("lbl_k_matvir_nm").Caption & vbNewLine & rsWO("item_nm")

                reqQTy = rsWO("qty") * rsWO("qty_mat")
                ttl_reqQty = ttl_reqQty + reqQTy
                '.Controls("lblqtyReq").Caption = .Controls("lblqtyReq").Caption & vbNewLine & reqQTy & " " & rsWO("um_name")
            Next
            .Controls("lbl_m_matid").Caption = .Controls("lbl_k_matid").Caption

            For r = 1 To rsWO.RecordCount
                rsWO.AbsolutePosition = r
                reqQTy = rsWO("qty") * rsWO("qty_mat")
                reqQTy_purg = (reqQTy / ttl_reqQty) * (rsWO("qty_prg") / 1000)

                .Controls("lblqtyReq").Caption = .Controls("lblqtyReq").Caption & vbNewLine & (reqQTy + reqQTy_purg) & " " & rsWO("um_name")
            Next
            .Controls("lblqtyReq_m").Caption = .Controls("lblqtyReq").Caption
        End With
        .Refresh
        Me.MousePointer = vbDefault
        If pPrintPP Then
            .PrintReport False, rptRangeAllPages
        Else
            .Show
        End If
    End With
    Exit Sub
Exc:
    If Err.Number = 8542 Then
        MsgBox "Ukuruan lebar kertas tidak memungkinkan, " & vbNewLine & " ganti tipe kertas atau Printer yang mendukung kertas A3", vbCritical, "Sorry " & Err.Number
        CommonDialog1.ShowPrinter
        MsgBox "Silahkan coba lagi", vbInformation
    Else
        MsgBox Err.Description, vbCritical, "Error No. : " & Err.Number
    End If
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

Private Sub Form_Activate()
    FocusTab Me
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    ' Is the mouse over the control
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hwnd, xpos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      ' If so, respond accordingly
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, xpos, Ypos
          
        Case TypeOf ctl Is PictureBox
          PictureBoxZoom ctl, MouseKeys, Rotation, xpos, Ypos
          
        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then ctl.SetFocus
          
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
  
End Sub


Private Sub Form_Load()
    AddTab Me
    settingFG
    activeTheme skinFD, Me
    BukaKoneksi
    Me.Height = 6195
    Me.Width = 11395
    dt1 = Now
    dt2 = dt1
'    Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WheelUnHook(Me.hwnd)
    DelTab Me
    Set rsBOM = Nothing
    Set rsWO = Nothing
    Set rsUpdateCT = Nothing
End Sub

Private Sub Label2_Click()
    Picture1.Visible = False
End Sub

Private Sub Label4_Click()
    PicExportContainer.Visible = False
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdfind_Click
        
    End If
End Sub
