VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_ReportDelvSche 
   Caption         =   "Delivery Report"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9090
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
   ScaleHeight     =   6915
   ScaleWidth      =   9090
   Begin VB.CommandButton findPartNo 
      Caption         =   "..."
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtPartNo 
      Height          =   390
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton findSOD 
      Caption         =   "..."
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.CheckBox chkAllSoId 
      Caption         =   "ALL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkAllPart 
      Caption         =   "ALL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1080
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtSO 
      Height          =   390
      Left            =   2040
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker tglFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   164233219
      CurrentDate     =   42664
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_ReportDelvSche.frx":0000
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid flx 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9340
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
   End
   Begin ACTIVESKINLibCtl.Skin Skinfd 
      Left            =   8160
      OleObjectBlob   =   "Form_ReportDelvSche.frx":007A
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   3720
      OleObjectBlob   =   "Form_ReportDelvSche.frx":02AE
      TabIndex        =   4
      Top             =   120
      Width           =   135
   End
   Begin MSComCtl2.DTPicker tglTo 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   164233219
      CurrentDate     =   42664
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_ReportDelvSche.frx":0308
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_ReportDelvSche.frx":0366
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form_ReportDelvSche"
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

Private Sub findPartNo_Click()
    GetForm = Me.Name
    PopUp_Item_Sup.Show 1
End Sub

Private Sub findSOD_Click()
    GetForm = Me.Name
    popUp_SOC.Show 1
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub SettingLV()
    With flx
        .Cols = 10
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "SO No"
'        .ColWidth(i) = 2600
        
        i = 1
        .TextMatrix(0, i) = "Customer"
        
        i = 2
        .TextMatrix(0, i) = "Part No"
        
        i = 3
        .TextMatrix(0, i) = "SO Date"
        
        i = 4
        .TextMatrix(0, i) = "Sch Del Date"
        
        i = 5
        .TextMatrix(0, i) = "Qty Sch Date"
        
        i = 6
        .TextMatrix(0, i) = "Qty Delivery"
        
        i = 7
        .TextMatrix(0, i) = "Qty OS Sched Del"
        
        i = 8
        .TextMatrix(0, i) = "Status"
        
        i = 9
        .TextMatrix(0, i) = "Remark"
        
        .MergeRow(0) = True
    End With
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

Private Sub Form_Load()
    AddTab Me
    Me.Width = 9330
    Me.Height = 7485
    Call SettingLV
    Call TemaAktif(Skinfd, Me)
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        DelTab Me
    End If
End Sub
