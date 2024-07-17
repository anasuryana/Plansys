VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_DelSchedule 
   Caption         =   "Delivery Schedule"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_DelSchedule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   16965
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FrameFilter 
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   7155
      TabIndex        =   10
      Top             =   360
      Width           =   7215
   End
   Begin VB.PictureBox FrameUpload 
      Height          =   1095
      Left            =   7560
      ScaleHeight     =   1035
      ScaleWidth      =   8715
      TabIndex        =   1
      Top             =   360
      Width           =   8775
      Begin VB.PictureBox Picture4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Failed"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label s_failed 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         Begin VB.Label s_success 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Success"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "UPLOAD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Period"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1260
         End
         Begin VB.Label l_period 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         Begin VB.Label l_totalAssy 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Assy"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGridDSch 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483631
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin skn 
      Left            =   0
      OleObjectBlob   =   "Form_DelSchedule.frx":000C
      Top             =   120
   End
   Begin MSComDlg.CommonDialog comDialogUpload 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "Form_DelSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim iLoop As Integer
Dim iGrid As Integer

Private Sub cmdSave_Click()
On Error GoTo errSave
    Dim affSch As Integer
    With MSFlexGridDSch
        For iLoop = 1 To .rows - 1
            On Error Resume Next
            Con.Execute "insert into plansys_schedule (period, date_period, assy_no, total_qty, input_user, input_time) values " _
                & "('" & l_period & "', '" & Left(l_period, 4) & "-" & Right(l_period, 2) & "-01" & "', '" & .TextMatrix(iLoop, 1) & "', " _
                & Val(.TextMatrix(iLoop, 2)) & ", '" & pUserName & "', now())", affSch
            .Row = iLoop
            .Col = 2
            If affSch > 0 Then
                .CellForeColor = vbGreen
                s_success = Val(s_success) + 1
            Else
                .CellForeColor = vbRed
                s_failed = Val(s_failed) + 1
            End If
            For i = 3 To .Cols - 1
                Con.Execute "insert into plansys_schedule_detail (period, date_period, assy_no, date_schedule, qty) values " _
                    & "('" & l_period & "', '" & Left(l_period, 4) & "-" & Right(l_period, 2) & "-01" & "', '" & .TextMatrix(iLoop, 1) & "', " _
                    & "'" & Format(DateSerial(Val(Left(l_period, 4)), Val(Right(l_period, 2)), Val(.TextMatrix(0, i))), "YYYY-MM-DD") & "', " _
                     & Val(.TextMatrix(iLoop, i)) & ")", affSch
                .Row = iLoop
                .Col = i
                If affSch > 0 Then
                    .CellForeColor = vbGreen
                Else
                    .CellForeColor = vbRed
                End If
            Next
        Next
    End With
Exit Sub
errSave:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Save: " & Err.Number
    End If
End Sub

Private Sub cmdUpload_Click()
On Error GoTo errUpload
    comDialogUpload.ShowOpen
    importDelSchedule comDialogUpload.FileName
Exit Sub
errUpload:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Upload: " & Err.Number
    End If
End Sub

Private Sub importDelSchedule(file As String)
On Error GoTo errUpExcel
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    
    Dim stExcel As Boolean
    
    Dim dateLoop As Date
    
    Dim fixSO As String * 30
    Dim chkUpload As Integer
    
    stExcel = False
    
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    
    ExcelObj.Workbooks.Open file
    
    
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
    
    stExcel = True
    
    With ExcelSheet
        l_period = Val(.Cells(3, 1))
        l_totalAssy = 0
    End With
    If Len(l_period) = 6 Then
        dateLoop = DateSerial(Val(Left(l_period, 4)), Val(Right(l_period, 2)), 1)
        With MSFlexGridDSch
            .rows = 1
            .Cols = 3
            .FixedCols = 0
            .Row = 0
            .TextMatrix(0, 0) = "NO"
            .TextMatrix(0, 1) = "ASSY NO"
            .TextMatrix(0, 2) = "TOTAL"
            .Col = 0
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
            .Col = 1
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
            .Col = 2
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
            .RowHeightMin = 300
            .ColWidth(0) = 900
            .ColWidth(1) = 3300
            .ColWidth(2) = 1800
            .ColAlignment(1) = flexAlignLeftCenter
            
            iLoop = 4
            Do Until Month(dateLoop) <> Val(Right(l_period, 2))
                .Cols = iLoop
                .Row = 0
                .Col = iLoop - 1
                .CellAlignment = flexAlignCenterCenter
                .CellFontBold = True
                .TextMatrix(0, iLoop - 1) = Day(dateLoop)
                iLoop = iLoop + 1
                dateLoop = dateLoop + 1
            Loop
            
            iLoop = 6
            iGrid = 1
            Do Until Trim(ExcelSheet.Cells(iLoop, 2)) = ""
                l_totalAssy = Val(l_totalAssy) + 1
                .rows = .rows + 1
                .TextMatrix(iGrid, 0) = iLoop - 5
                .TextMatrix(iGrid, 1) = RTrim(ExcelSheet.Cells(iLoop, 2))
                .TextMatrix(iGrid, 2) = 0
                For i = 3 To .Cols - 1
                    .TextMatrix(iGrid, i) = Val(ExcelSheet.Cells(iLoop, i + 1))
                    .TextMatrix(iGrid, 2) = Val(.TextMatrix(iGrid, 2)) + Val(ExcelSheet.Cells(iLoop, i + 1))
                Next
                iGrid = iGrid + 1
                iLoop = iLoop + 1
            Loop
            
            If .rows > 1 Then
                .FixedRows = 1
                .FixedCols = 3
                cmdSave.Enabled = True
            End If
        End With
    Else
        MsgBox "Periksa Format Periode!", vbExclamation, "Period = False"
    End If
    ExcelObj.Workbooks.Close
    
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
Exit Sub
errUpExcel:
    If Err.Number <> 0 Then
        If stExcel = True Then
            ExcelObj.Workbooks.Close
        End If
        Set ExcelSheet = Nothing
        Set ExcelBook = Nothing
        Set ExcelObj = Nothing
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    activeTheme skn, Me
    Call BukaKoneksi
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    MSFlexGridDSch.Width = Me.Width - 600
    FrameUpload.Left = Me.Width - FrameUpload.Width - 480
End Sub

