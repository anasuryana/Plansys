VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_GenMPP 
   Caption         =   "Generate Menu Loading"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14415
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
   ScaleHeight     =   7305
   ScaleWidth      =   14415
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
      Begin VB.ComboBox cmbKol 
         Height          =   390
         ItemData        =   "Form_GenMPP.frx":0000
         Left            =   120
         List            =   "Form_GenMPP.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
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
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   0
         Width           =   4215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7050
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   24897
            Text            =   "Hello"
            TextSave        =   "Hello"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log"
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   5295
      Left            =   50
      ScaleHeight     =   5235
      ScaleWidth      =   14235
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   14295
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   1215
         Left            =   480
         TabIndex        =   17
         Top             =   3480
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   2143
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Generating data..."
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   735
         Left            =   480
         TabIndex        =   16
         Top             =   2760
         Width           =   11055
      End
   End
   Begin VB.ComboBox CmbDocument 
      Height          =   390
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   1140
      Left            =   10320
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   550
      Width           =   3615
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   550
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CMdNExt 
      Caption         =   ">"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   550
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox txtRevision 
      Height          =   390
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   6840
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid anaGrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Warna kuning artinya mesin telah penuh"
      Top             =   1320
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9763
      _Version        =   393216
      MergeCells      =   1
      AllowUserResizing=   1
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
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
      CustomFormat    =   "yyyyMM"
      Format          =   244580355
      CurrentDate     =   42544
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_GenMPP.frx":0021
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   9195
      OleObjectBlob   =   "Form_GenMPP.frx":0083
      Top             =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form_GenMPP.frx":02B7
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   3360
      OleObjectBlob   =   "Form_GenMPP.frx":031D
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   3360
      OleObjectBlob   =   "Form_GenMPP.frx":038D
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form_GenMPP"
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
Private qry As String
Private idS As String
Private RsA As ADODB.Recordset
Private rsB As ADODB.Recordset
Private rsW As ADODB.Recordset
Private rsX As ADODB.Recordset
Private rsPartMCH As ADODB.Recordset
Private temp_part As String

Private posisisFind As Long
Dim nm_msn_full() As String
Dim nm_mold_full() As String
Dim ar_prodplan() As Variant, c_part() As String, c_part_saved() As String
Dim ar_kpsts_mc_hr() As String
Dim ar_kpsts_mold_hr() As String
Dim ar_nm_msn() As String
Dim ar_nm_mold() As String
Dim ar_hkw(1 To 4) As String
Dim ar_hkw_bln(1 To 4) As String
Dim ar_hkw_bln_th(1 To 4) As String
Dim nmbulan() As String
Dim nextRev As Integer
Dim ym As Date
Dim aPartPrior() As String
Dim dob As Long
Dim dob2 As Long
Dim dob2c As Double
Dim dBariss As Long

' untuk keperluan pembulatan sub isi(....,....,...)

Dim ar_propl() As Variant
Dim ar_propl2() As Variant
Dim ar_propl3() As Variant
Dim ar_propl4() As Variant
Const minND As Byte = 1

'sisa 1 hari
Dim ar_Sisa() As Variant
Dim ar_PartSisa() As String
Dim need_day As Variant
Dim ar_MoldSisa() As String
Dim ar_MesinSisa() As String
Dim ar_fMesin() As String


Dim ttMold As String
Dim c_NDMtZ As Variant
Dim tToutalMold As Integer

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

Private Sub settingFG()
    Dim i As Byte
    With anaGrid
        .Cols = 37: .ColWidth(0) = 700: .ColWidth(1) = 2800: .ColWidth(2) = 3000: .ColWidth(3) = 3000
        .rows = 5
        .FixedRows = 3
        .FixedCols = 4
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "No":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 1
        .TextMatrix(0, i) = "Customer":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 2
        .TextMatrix(0, i) = "Assy no":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = "Assy Desc":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = "STOCK FG":         .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = "STOCK WIP":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 6
        .TextMatrix(0, i) = "FC":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 7
        .TextMatrix(0, i) = "ITO":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 2000
        
        i = 8
        .TextMatrix(0, i) = "SUB CONT":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 700
        
        i = 9
        .TextMatrix(0, i) = "PROD PLAN 1":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 10
        .TextMatrix(0, i) = "PROD PLAN 2":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 0
        
        i = 11
        .TextMatrix(0, i) = "PROD PLAN 3":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 0
        
        
        i = 12
        .TextMatrix(0, i) = "PROD PLAN 4":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 0
        
        i = 13
        .TextMatrix(0, i) = "Cav":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 600
        
        i = 14
        .TextMatrix(0, i) = "C/T":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 600
        
        i = 15
        .TextMatrix(0, i) = "Man Power":          .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 16
        .TextMatrix(0, i) = "2nd Proses":          .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 17
        .TextMatrix(0, i) = "Cap /day":         .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 18
        .TextMatrix(0, i) = "Cap /mold":         .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 19
        .TextMatrix(0, i) = "Cap /Month":         .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 20
        .TextMatrix(0, i) = "Need day": .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 21
        .TextMatrix(0, i) = "Sum Need day": .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 22
        .TextMatrix(0, i) = "MC No": .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 23
        .TextMatrix(0, i) = "Tonage":        .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 24
        .TextMatrix(0, i) = "Mold":        .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 25
        .TextMatrix(0, i) = "%"
        .TextMatrix(1, i) = .TextMatrix(0, i)
        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 0 '750
        
        i = 26
        .TextMatrix(0, i) = "Mach State": .TextMatrix(1, i) = .TextMatrix(0, i):     .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        
        .TextMatrix(0, 27) = "Load Vs Cap Machine"
        .MergeCol(27) = True
        
        .TextMatrix(0, 28) = "Need of Operator"
        .ColWidth(28) = 2000
        .MergeCol(28) = True
        
        .TextMatrix(0, 29) = "Mold Used"
        .MergeCol(29) = True
        .ColWidth(29) = 0
        .TextMatrix(0, 30) = "Productivity Factor"
        .ColWidth(30) = 0
        .TextMatrix(0, 31) = "Hour per shift"
        .ColWidth(31) = 0
        .TextMatrix(0, 32) = "total Shift"
        .ColWidth(32) = 0
        .TextMatrix(0, 33) = "Processed Qty"
        .MergeCol(33) = True
        .TextMatrix(0, 34) = "cavity std"
        .ColWidth(34) = 5
        .TextMatrix(0, 35) = "Time Update"
        .ColAlignment(35) = flexAlignRightCenter
        .TextMatrix(0, 36) = "%prodpan"
        .ColAlignment(36) = flexAlignRightCenter
        .MergeRow(0) = True
        .MergeRow(2) = True
        
        .Refresh
    End With
End Sub

Private Sub settingGridName()
    Dim it As Integer, it2 As Integer
    
    With anaGrid
        For it = 0 To 3
            If Val(Format(DTPicker1, "M")) + it > 12 Then
                it2 = it2 + 1
                ar_hkw_bln(it + 1) = nmbulan(it2)
                ar_hkw_bln_th(it + 1) = Val(Format(DTPicker1, "yyyy")) * 1 + 1
            Else
                ar_hkw_bln(it + 1) = nmbulan(Val(Format(DTPicker1, "M")) + it)
                ar_hkw_bln_th(it + 1) = Format(DTPicker1, "yyyy")
            End If
        Next
    End With
   
    Text1 = ar_hkw_bln(1)
    SkinLabel4 = "HKW [" & ar_hkw(1) & "]"
    settingHeaderMonth
End Sub

Private Sub settingHeaderMonth()
    anaGrid.TextMatrix(1, 27) = Text1
    anaGrid.TextMatrix(2, 27) = Text1
    anaGrid.TextMatrix(1, 28) = Text1
    anaGrid.TextMatrix(2, 28) = Text1
End Sub

Public Function DaysInMonth(ByVal dDate As Date) As Integer
    DaysInMonth = Day(DateAdd("m", 1, dDate - Day(dDate) + 1) - 1)
End Function

Private Sub gridFormatNum()
    Dim v As Long
    For v = 3 To anaGrid.rows - 1
        With anaGrid
            .TextMatrix(v, 6) = FormatNumber(.TextMatrix(v, 6), 0)
            .TextMatrix(v, 7) = FormatNumber(.TextMatrix(v, 7), 4)
            If .TextMatrix(v, 9) <> "" Then
                .TextMatrix(v, 9) = FormatNumber(.TextMatrix(v, 9), 0)
            End If
            
            If .TextMatrix(v, 10) <> "" Then
                .TextMatrix(v, 10) = FormatNumber(.TextMatrix(v, 10), 0)
            End If
            
            If .TextMatrix(v, 11) <> "" Then
                .TextMatrix(v, 11) = FormatNumber(.TextMatrix(v, 11), 0)
            End If
            
            .TextMatrix(v, 12) = FormatNumber(.TextMatrix(v, 12), 0)
            .TextMatrix(v, 17) = FormatNumber(.TextMatrix(v, 17), 0)
            .TextMatrix(v, 18) = FormatNumber(.TextMatrix(v, 18), 0)
            .TextMatrix(v, 19) = FormatNumber(.TextMatrix(v, 19), 0)
            .TextMatrix(v, 20) = FormatNumber(.TextMatrix(v, 20), 2) 'FormatNumber(IIf(IsNumeric(.TextMatrix(v, 20)), .TextMatrix(v, 20),  933), 2)
            .TextMatrix(v, 21) = FormatNumber(.TextMatrix(v, 21), 2)
            
            .TextMatrix(v, 27) = FormatNumber(.TextMatrix(v, 27), 2) & " %"
            .TextMatrix(v, 28) = FormatNumber(.TextMatrix(v, 28), 2)
            If IsNumeric(.TextMatrix(v, 29)) Then
                .TextMatrix(v, 29) = FormatNumber(.TextMatrix(v, 29), 2) 'FormatNumber(IIf(IsNumeric(.TextMatrix(v, 29)), .TextMatrix(v, 29), 0), 2)
            Else
                .TextMatrix(v, 29) = 0
                .TextMatrix(v, 29) = FormatNumber(.TextMatrix(v, 29), 2) 'FormatNumber(IIf(IsNumeric(.TextMatrix(v, 29)), .TextMatrix(v, 29), 0), 2)
            End If
        End With
    Next
End Sub

Private Sub anaGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText anaGrid.Clip
    ElseIf KeyCode = 70 And Shift = 2 Then
        PicFIND.Visible = True
        txtFindNext.SetFocus
    End If
   
End Sub

Private Function checkSisaProdplan(pPART As String) As Variant
    Dim r As Integer
    For r = 1 To UBound(c_part)
        If c_part(r) = pPART Then
            checkSisaProdplan = ar_prodplan(r)
            Exit For
        End If
    Next
End Function

Private Function checkSisaHariMesin(pmesin As String) As Variant
    Dim w As Integer
    For w = 1 To UBound(ar_kpsts_mc_hr)
        If ar_nm_msn(w) = pmesin Then
            checkSisaHariMesin = ar_kpsts_mc_hr(w)
            'If pmesin = "INJ-17" Then List2.AddItem "*inj-17=" & checkSisaHariMesin
            Exit For
        End If
    Next
End Function

Private Function checkSisaHariMold(pMOLD As String) As Variant
    Dim w As Integer
    For w = 1 To UBound(ar_kpsts_mold_hr)
        If ar_nm_mold(w) = pMOLD Then
            checkSisaHariMold = ar_kpsts_mold_hr(w)
            Exit For
        End If
    Next
End Function

Private Sub ProsesPresent(tR As Long, p_statemach As Byte, p_hkw As Variant, p_capday As Variant)
    Dim rx As Integer, rx2 As Integer
    Dim tempnd As Variant
    With anaGrid
        If p_statemach > 0 And .TextMatrix(tR, 24) <> "nomo" Then ' JIKA MESIN dan mould AKTIF
            If checkSisaProdplan(.TextMatrix(tR, 2)) > 0 Then
                If checkSisaHariMesin(.TextMatrix(tR, 22)) > 0 Then
                    If checkSisaHariMold(.TextMatrix(tR, 24)) > 0 Then
                        If checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday * 1 > p_hkw Then 'jika need day > hkw [.TextMatrix(tR, 20) * 1]
                            .TextMatrix(tR, 20) = .TextMatrix(tR, 20) * 1 + (checkSisaHariMesin(.TextMatrix(tR, 22)) * 1)
                            .TextMatrix(tR, 29) = .TextMatrix(tR, 29) * 1 + (checkSisaHariMold(.TextMatrix(tR, 24)) * 1)

                            kurangihariMesin .TextMatrix(tR, 22), checkSisaHariMesin(.TextMatrix(tR, 22))
                            kurangihariMold .TextMatrix(tR, 24), checkSisaHariMold(.TextMatrix(tR, 24))
                            
                            If UBound(nm_msn_full) = 1 And nm_msn_full(1) = "" Then
                                nm_msn_full(1) = .TextMatrix(tR, 22)
                                nm_mold_full(1) = .TextMatrix(tR, 24)
                            Else
                                ReDim Preserve nm_msn_full(1 To UBound(nm_msn_full) + 1) As String
                                ReDim Preserve nm_mold_full(1 To UBound(nm_mold_full) + 1) As String
                                nm_msn_full(UBound(nm_msn_full)) = .TextMatrix(tR, 22)
                                nm_mold_full(UBound(nm_mold_full)) = .TextMatrix(tR, 24)
                            End If
                            kurangiProdplan .TextMatrix(tR, 2), (p_capday * 1) * Val(.TextMatrix(tR, 20) * 1)
                            .TextMatrix(tR, 33) = ((p_capday * 1) * (.TextMatrix(tR, 20) * 1))
                        Else
'                            If .TextMatrix(tR, 2) = "32107-K81-N000-H1" Then
'                                            MsgBox .TextMatrix(tR, 33)
'                                        End If
                            If checkSisaHariMesin(.TextMatrix(tR, 22)) * 1 <= checkSisaHariMold(.TextMatrix(tR, 24)) Then
                                If checkSisaHariMesin(.TextMatrix(tR, 22)) * 1 >= (checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday) Then
                                    If checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday >= minND * 1 Then
                                        tempnd = checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday
                                        .TextMatrix(tR, 20) = .TextMatrix(tR, 20) * 1 + tempnd * 1 '(checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday)
                                        .TextMatrix(tR, 29) = .TextMatrix(tR, 29) * 1 + tempnd * 1 '(checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday)

                                        kurangihariMesin .TextMatrix(tR, 22), tempnd * 1 ' .TextMatrix(tR, 20)
                                        
                                        kurangihariMold .TextMatrix(tR, 24), tempnd * 1 ' '.TextMatrix(tR, 20) * 1
                                        kurangiProdplan .TextMatrix(tR, 2), (p_capday * 1) * tempnd * 1 ' Val(.TextMatrix(tR, 20) * 1)
'                                         If .TextMatrix(tR, 2) = "32107-K81-N000-H1" Then
'                                            MsgBox (p_capday * 1) * (.TextMatrix(tR, 20) * 1)
'                                        End If
                                        .TextMatrix(tR, 33) = ((p_capday * 1) * (.TextMatrix(tR, 20) * 1))  '(.TextMatrix(tR, 20) * 1)
                                        tempnd = 0
                                    End If
                                Else
                                    If checkSisaHariMesin(.TextMatrix(tR, 22)) * 1 >= minND * 1 Then
                                        tempnd = checkSisaHariMesin(.TextMatrix(tR, 22)) * 1
                                        .TextMatrix(tR, 20) = .TextMatrix(tR, 20) * 1 + (tempnd)
                                        .TextMatrix(tR, 29) = .TextMatrix(tR, 29) * 1 + (tempnd)
                                        
                                        kurangihariMesin .TextMatrix(tR, 22), tempnd ' .TextMatrix(tR, 20)
    
                                        kurangihariMold .TextMatrix(tR, 24), tempnd ' .TextMatrix(tR, 20)
    
                                        If checkSisaHariMesin(.TextMatrix(tR, 22)) * 1 = 0 Then
                                            If UBound(nm_msn_full) = 1 And nm_msn_full(1) = "" Then
                                                nm_msn_full(1) = .TextMatrix(tR, 22)
                                            Else
                                                ReDim Preserve nm_msn_full(1 To UBound(nm_msn_full) + 1) As String
                                                nm_msn_full(UBound(nm_msn_full)) = .TextMatrix(tR, 22)
                                            End If
                                        End If
                                        kurangiProdplan .TextMatrix(tR, 2), (p_capday * 1) * tempnd ' Val(.TextMatrix(tR, 20) * 1)
                                        .TextMatrix(tR, 33) = (p_capday * 1) * (.TextMatrix(tR, 20) * 1)
                                        tempnd = 0
                                    End If
                                End If
                            Else
                                If checkSisaHariMold(.TextMatrix(tR, 24)) * 1 >= (checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday) Then
                                    If checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday >= minND * 1 Then
                                        tempnd = checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday
                                        .TextMatrix(tR, 20) = .TextMatrix(tR, 20) * 1 + tempnd * 1 '(checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday)
                                        .TextMatrix(tR, 29) = .TextMatrix(tR, 29) * 1 + tempnd * 1 '(checkSisaProdplan(.TextMatrix(tR, 2)) / p_capday)
                                        
                                        kurangihariMesin .TextMatrix(tR, 22), tempnd * 1 ' .TextMatrix(tR, 29)
                                            
                                        kurangihariMold .TextMatrix(tR, 24), tempnd * 1 '.TextMatrix(tR, 29)
                                        kurangiProdplan .TextMatrix(tR, 2), (p_capday * 1) * tempnd * 1 'Val(.TextMatrix(tR, 29) * 1)
                                        .TextMatrix(tR, 33) = ((p_capday * 1) * .TextMatrix(tR, 29))   'Val(.TextMatrix(tR, 29) * 1)
                                        tempnd = 0
                                    End If
                                Else
                                    If checkSisaHariMold(.TextMatrix(tR, 24)) * 1 >= minND * 1 Then
                                        tempnd = checkSisaHariMold(.TextMatrix(tR, 24))
                                        .TextMatrix(tR, 20) = .TextMatrix(tR, 20) * 1 + (tempnd)
                                        .TextMatrix(tR, 29) = .TextMatrix(tR, 29) * 1 + (tempnd)

                                        kurangihariMesin .TextMatrix(tR, 22), tempnd '.TextMatrix(tR, 29)
    
                                        kurangihariMold .TextMatrix(tR, 24), tempnd '.TextMatrix(tR, 29)
                                        If checkSisaHariMesin(.TextMatrix(tR, 22)) * 1 = 0 Then
                                            If UBound(nm_mold_full) = 1 And nm_mold_full(1) = "" Then
                                                nm_mold_full(1) = .TextMatrix(tR, 24)
                                            Else
                                                ReDim Preserve nm_mold_full(1 To UBound(nm_mold_full) + 1) As String
                                                nm_mold_full(UBound(nm_mold_full)) = .TextMatrix(tR, 24)
                                            End If
                                        End If
                                        kurangiProdplan .TextMatrix(tR, 2), (p_capday * 1) * tempnd
                                        .TextMatrix(tR, 33) = (p_capday * 1) * .TextMatrix(tR, 29) * 1
                                    End If
                                End If
                            End If
                        End If
                    Else
                        .TextMatrix(tR, 20) = 0
                        .TextMatrix(tR, 29) = 0
                    End If
                Else
                    .TextMatrix(tR, 20) = 0
                    .TextMatrix(tR, 29) = 0
                End If
            Else
'                .TextMatrix(tR, 20) = 0
'                .TextMatrix(tR, 29) = 0
            End If
        End If
    End With
End Sub

Private Sub kurangihariMesin(pmesin As String, ppengurang As Variant)
    Dim b As Integer
    For b = 1 To UBound(ar_kpsts_mc_hr)
        If ar_nm_msn(b) = pmesin Then
            ar_kpsts_mc_hr(b) = Val(ar_kpsts_mc_hr(b) * 1) - ppengurang * 1
            Exit Sub
        End If
    Next
End Sub

Private Sub kurangihariMold(pMOLD As String, ppengurang As Variant)
    Dim b As Integer
    For b = 1 To UBound(ar_kpsts_mold_hr)
        If ar_nm_mold(b) = pMOLD Then
            ar_kpsts_mold_hr(b) = Val(ar_kpsts_mold_hr(b) * 1) - ppengurang * 1
            Exit Sub
        End If
    Next
End Sub

Private Sub kurangiProdplan(pPART As String, ppengurang As Variant)
    Dim b As Integer
    For b = 1 To UBound(c_part)
        If c_part(b) = pPART Then
            ar_prodplan(b) = Val(ar_prodplan(b) * 1) - Round(ppengurang * 1)
            Exit Sub
        End If
    Next
End Sub

Private Sub CmbDocument_DropDown()
    'If IsNumeric(txtRevision) = False Then txtRevision.SetFocus: Exit Sub
    qry = "select distinct on (ltpp_doc) ltpp_doc from ltpp_generate where  period='" & Format(DTPicker1.Value, "yyyyMM") & "'" 'rev=" & txtRevision & " and
    Set RsA = Con.Execute(qry)
    CmbDocument.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            CmbDocument.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub

Private Sub NOP_Generate(phkw_tsb As Integer, pi_hkw As Long)
    '---------------------------Bismillah----------------------
    formatWarnaBG
    Dim i As Integer, c_wip As Variant, c_cap_p_day As Variant, j As Integer, x As Integer
    Dim presentMesinUse As Variant
    Dim nomold As String
    Dim hkw1 As Integer
    Dim finPP As Double
    '------------------------PRODUCTION PLAN
    qry = "select distinct on (assy_no) " _
         & "prod_plan_" & pi_hkw & ",g.ct,g.cavity,hour_p_shift,shift_usg,faktor_productivity " _
         & ",item_muloq,item_perbox,assy_no,prc_prodplan from ltpp_generate a " _
         & "inner join mst_item b on a.assy_no=b.item_id " _
         & "inner join r_customer c on b.cust_id=c.cust_id " _
         & "inner join loadcap_mst_product_r d on a.assy_no=d.partno " _
         & "inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc " _
         & "inner join loadcap_proc g on d.partno=g.partno " _
         & "left join loadcap_mst_mach e on g.prod_nomach=e.no_mach " _
         & "where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='')"
         
         '& " and (prod_plan_1>0 or prod_plan_2>0 or prod_plan_3>0 or prod_plan_4>0)"
    Set rsB = Con.Execute(qry)
    
    i = 1
    If rsB.RecordCount > 0 Then
        ReDim ar_prodplan(1 To rsB.RecordCount) As Variant
        While Not rsB.EOF
            If rsB("ct") = 0 Then
                c_cap_p_day = 0
            Else
                c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
            End If
            c_cap_p_day = FormatNumber(c_cap_p_day, 0) * 1
            finPP = rsB(0) * rsB("prc_prodplan") / 100
            If c_cap_p_day * 1 > finPP Then
            
                If rsB(0) > 0 Then
                    If rsX.RecordCount = 0 Then
                        ar_prodplan(i) = c_cap_p_day
                        If rsB("item_perbox") = 0 Then
                            ar_prodplan(i) = isi(rsB("item_muloq"), c_cap_p_day, "b")
                        Else
                            ar_prodplan(i) = isi(rsB("item_perbox"), c_cap_p_day, "b")
                        End If
                    Else
                        If pi_hkw = 1 Then
                            If wo_runeqpp(Trim(rsB("assy_no"))) Then
                                ar_prodplan(i) = 0
                            Else
                                ar_prodplan(i) = c_cap_p_day
                                If rsB("item_perbox") = 0 Then
                                    ar_prodplan(i) = isi(rsB("item_muloq"), c_cap_p_day, "b") - wo_prodplan(Trim(rsB("assy_no")))
                                Else
                                    ar_prodplan(i) = isi(rsB("item_perbox"), c_cap_p_day, "b") - wo_prodplan(Trim(rsB("assy_no")))
                                End If
                            End If
                        End If
                    End If
                Else
                    ar_prodplan(i) = 0
                End If
            Else
                If finPP > 0 Then
                    If rsX.RecordCount = 0 Then
                        ar_prodplan(i) = finPP
                        If rsB("item_perbox") = 0 Then
                            ar_prodplan(i) = isi(rsB("item_muloq"), finPP, "b")
                        Else
                            ar_prodplan(i) = isi(rsB("item_perbox"), finPP, "b")
                        End If
                    Else
                        If pi_hkw = 1 Then
                            ar_prodplan(i) = finPP
                            If rsB("item_perbox") = 0 Then
                                ar_prodplan(i) = isi(rsB("item_muloq"), finPP, "b") - wo_prodplan(Trim(rsB("assy_no")))
                            Else
                                ar_prodplan(i) = isi(rsB("item_perbox"), finPP, "b") - wo_prodplan(Trim(rsB("assy_no")))
                            End If
                        End If
                    End If
                Else
                    ar_prodplan(i) = 0
                End If
            End If
            
            i = i + 1
            rsB.MoveNext
        Wend
    Else
        ReDim ar_prodplan(1) As Variant
    End If
    
    '----------------mendapatkan mesin yang dipakai (distict) GEUR-----------------
    qry = "select distinct on (prod_nomach) prod_nomach " _
         & "  from ltpp_generate a " _
         & " inner join mst_item b on a.assy_no=b.item_id " _
         & " inner join r_customer c on b.cust_id=c.cust_id " _
         & " inner join loadcap_mst_product_r d on a.assy_no=d.partno " _
         & " inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc " _
         & " inner join loadcap_proc g on d.partno=g.partno " _
         & " left join loadcap_mst_mach e on g.prod_nomach=e.no_mach " _
         & " where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='')"
    Set rsB = Con.Execute(qry)
    i = 1
    If rsB.RecordCount > 0 Then
        ReDim ar_kpsts_mc_hr(1 To rsB.RecordCount) As String
        ReDim ar_nm_msn(1 To rsB.RecordCount) As String
        While Not rsB.EOF
            If pi_hkw = 1 Then
                ar_kpsts_mc_hr(i) = phkw_tsb - wo_hkwmachine(rsB(0)) 'rsB(0)
            Else
                ar_kpsts_mc_hr(i) = phkw_tsb
            End If
            ar_nm_msn(i) = IIf(IsNull(rsB(0)), "nom", rsB(0))
            i = i + 1
            rsB.MoveNext
        Wend
    Else
        ReDim ar_kpsts_mc_hr(1) As String
        ReDim ar_nm_msn(1) As String
    End If
    
     '---------------------KAPASITAS HARI MOLD-------------
    qry = "select distinct on (mold_no) mold_no " _
          & " from ltpp_generate a " _
          & " inner join mst_item b on a.assy_no=b.item_id " _
          & " inner join r_customer c on b.cust_id=c.cust_id " _
          & " inner join loadcap_mst_product_r d on a.assy_no=d.partno " _
          & " inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc " _
          & " inner join loadcap_proc g on d.partno=g.partno " _
          & " left join loadcap_mst_mach e on g.prod_nomach=e.no_mach " _
          & " where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='')"
    Set rsB = Con.Execute(qry)
    i = 1
    If rsB.RecordCount > 0 Then
        ReDim ar_kpsts_mold_hr(1 To rsB.RecordCount) As String
        While Not rsB.EOF
            If pi_hkw = 1 Then
                ar_kpsts_mold_hr(i) = phkw_tsb - wo_hkwmold(rsB(0)) 'rsB(0)
            Else
                ar_kpsts_mold_hr(i) = phkw_tsb
            End If
            ar_nm_mold(i) = IIf(IsNull(rsB(0)), "nomo", rsB(0))
            i = i + 1
            rsB.MoveNext
        Wend
    Else
        ReDim ar_kpsts_mold_hr(1) As String
        ReDim ar_nm_mold(1) As String
    End If
        
    '-----------------------------YUK KITA PIKIRKAN---------------------------
    i = 0
    qry = "select cust_name,assy_no,a.item_name,fg,p1,p2,p3,fc1 " _
        & " ,prod_plan_1,prod_plan_2,prod_plan_3,prod_plan_4 " _
        & " ,g.cavity,g.ct,g.manpower,g.ct_2,g.prod_nomach " _
        & " ,coalesce(e.tonage_mach,0) tonage_mach,case when (g.cavity=0 or g.ct=0) then 0 else round((prod_plan_" & pi_hkw & "/((60 / g.ct) * g.cavity * hour_p_shift * shift_usg * 60 )*faktor_productivity)/" & phkw_tsb & "*100,6) end presenku " _
        & " ,faktor_productivity,state_mach,mold_no,subcont,hour_p_shift,shift_usg,cavity_std,item_muloq,item_perbox, " _
        & " priorit,submch,timeupdate,prc_prodplan from ltpp_generate a " _
        & " inner join mst_item b on a.assy_no=b.item_id " _
        & " inner join r_customer c on b.cust_id=c.cust_id " _
        & " inner join loadcap_mst_product_r d on a.assy_no=d.partno" _
        & " inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc" _
        & " inner join loadcap_proc g on d.partno=g.partno" _
        & " left join loadcap_mst_mach e on g.prod_nomach=e.no_mach" _
        & " where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='')" _
        & " order by subcont asc,19 desc,priorit asc,2 " ' ,prod_nomach asc"
        '& " and (prod_plan_1>0 or prod_plan_2>0 or prod_plan_3>0 or prod_plan_4>0)"
    Set rsB = Con.Execute(qry)
    Erase aPartPrior
    ReDim aPartPrior(1 To 1) As String
    resetArrayOVR
    ReDim nm_msn_full(1 To 1) As String
    ReDim nm_mold_full(1 To 1) As String
    If rsB.RecordCount > 0 Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = rsB.RecordCount
        anaGrid.rows = 3
        While Not rsB.EOF
            rsKeArray RTrim(rsB(1))
            rsB.MoveNext
        Wend
        rsB.Fields("assy_no").Properties("Optimize") = True
        rsB.Fields("priorit").Properties("Optimize") = True
        rsB.Fields("subcont").Properties("Optimize") = True
        rsB.Fields("mold_no").Properties("Optimize") = True
        
        '&&& ane
        StatusBar1.Panels(1).Text = rsB.RecordCount & " row(s)"
        rsW.Filter = adFilterNone
        
        If rsW.RecordCount = 0 Then
'            MsgBox "Sini 2"
            For dob = 1 To UBound(aPartPrior) - 1
                rsB.Filter = adFilterNone
                rsB.Filter = "assy_no='" & aPartPrior(dob) & "'"
                rsB.Sort = "priorit ASC"
                dob2 = 1
'                If Trim(aPartPrior(dob)) = "BEI-KMI-017" Then
'                    MsgBox "ke sini 1"
'                End If
'                MsgBox "ke sini #" & dob & " ## " & aPartPrior(dob)
                temp_part = aPartPrior(dob)
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                        rsB.AbsolutePosition = dob2
                        c_wip = rsB("p1") + rsB("p2") + rsB("p3")
                        If rsB("ct") = 0 Then
                            c_cap_p_day = 0
                        Else
                            c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
'                            If rsB("item_perbox") = 0 Then
'                                c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
'                            Else
'                                c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
'                            End If
                        End If
                        hkw1 = phkw_tsb
                        presentMesinUse = rsB("presenku")
                        
                        If IsNull(rsB("mold_no")) Or rsB("mold_no") = "" Or rsB("mold_no") = "0" Then
                            nomold = "nomo"
                        Else
                            nomold = rsB("mold_no")
                        End If
                        
                        anaGrid.rows = anaGrid.rows + 1
                        dBariss = anaGrid.rows - 1
                        anaGrid.TextMatrix(dBariss, 0) = anaGrid.rows - 3
                        anaGrid.TextMatrix(dBariss, 1) = rsB(0)
                        anaGrid.TextMatrix(dBariss, 2) = RTrim(rsB(1))
                        anaGrid.TextMatrix(dBariss, 3) = rsB(2)
                        anaGrid.TextMatrix(dBariss, 4) = FormatNumber(rsB(3), 0)
                        anaGrid.TextMatrix(dBariss, 5) = c_wip
                        anaGrid.TextMatrix(dBariss, 6) = rsB("fc1")
    
                        If rsB("fc1") = 0 Then
                            anaGrid.TextMatrix(dBariss, 7) = 0
                        Else
                            anaGrid.TextMatrix(dBariss, 7) = (rsB(3) + c_wip) / rsB("fc1")
                        End If
    
                        anaGrid.TextMatrix(dBariss, 8) = IIf(IsNull(rsB("subcont")), "no", rsB("subcont")) 'kebijakan_subc
                        anaGrid.TextMatrix(dBariss, 9) = getProdplan(RTrim(rsB(1)), 1) 'rsB("prod_plan_1")
                        anaGrid.TextMatrix(dBariss, 10) = getProdplan(RTrim(rsB(1)), 2) 'rsB("prod_plan_2")
                        anaGrid.TextMatrix(dBariss, 11) = getProdplan(RTrim(rsB(1)), 3) 'rsB("prod_plan_3")
                        anaGrid.TextMatrix(dBariss, 12) = getProdplan(RTrim(rsB(1)), 4) 'rsB("prod_plan_4")
                        anaGrid.TextMatrix(dBariss, 13) = rsB("cavity")
                        anaGrid.TextMatrix(dBariss, 14) = rsB("ct")
                        anaGrid.TextMatrix(dBariss, 15) = rsB("manpower")
                        anaGrid.TextMatrix(dBariss, 16) = rsB("ct_2")
                        anaGrid.TextMatrix(dBariss, 17) = c_cap_p_day
                        anaGrid.TextMatrix(dBariss, 18) = c_cap_p_day
                        anaGrid.TextMatrix(dBariss, 19) = c_cap_p_day * hkw1
                        anaGrid.TextMatrix(dBariss, 31) = rsB("hour_p_shift")
                        anaGrid.TextMatrix(dBariss, 32) = rsB("shift_usg")
                        anaGrid.TextMatrix(dBariss, 20) = 0
                        anaGrid.TextMatrix(dBariss, 29) = 0
                        anaGrid.TextMatrix(dBariss, 21) = 0
                        anaGrid.TextMatrix(dBariss, 22) = IIf(IsNull(rsB("prod_nomach")), "nom", rsB("prod_nomach")) 'rsB("tonage_mach")
                        anaGrid.TextMatrix(dBariss, 23) = IIf(IsNull(rsB("tonage_mach")), "nom", rsB("tonage_mach"))
                        anaGrid.TextMatrix(dBariss, 24) = nomold 'IIf(IsNull(rsB("mold_no")) Or IsNumeric(rsB("mold_no")) = False, "nomo", rsB("mold_no"))
                        anaGrid.TextMatrix(dBariss, 25) = 0
    
                        anaGrid.TextMatrix(dBariss, 26) = IIf(IsNull(rsB("state_mach")), 0, rsB("state_mach")) 'IIf(ovrd_msnutama < 0, 0, ovrd_msnutama)
                        anaGrid.TextMatrix(dBariss, 34) = IIf(IsNull(rsB("cavity_std")), 0, rsB("cavity_std"))
                        anaGrid.TextMatrix(dBariss, 27) = 0
                        anaGrid.TextMatrix(dBariss, 30) = IIf(IsNull(rsB("faktor_productivity")), 0, rsB("faktor_productivity"))
                        anaGrid.TextMatrix(dBariss, 33) = 0
                        anaGrid.TextMatrix(dBariss, 35) = IIf(IsNull(rsB("timeupdate")), "", Format(rsB("timeupdate"), "yyyy-MM-dd HH:mm:ss"))
                        anaGrid.TextMatrix(dBariss, 36) = rsB("prc_prodplan")
                        
                        If IsNull(rsB("state_mach")) Or nomold = "nomo" Or nomold = "0" Or anaGrid.TextMatrix(dBariss, 8) = "yes" Then
                            If rsB("submch") = True Then
                                ProsesPresent dBariss, 1, hkw1, c_cap_p_day
                            Else
                                ProsesPresent dBariss, 0, hkw1, c_cap_p_day
                            End If
                        Else
                            rsPartMCH.Filter = adFilterNone
                            rsPartMCH.Filter = "part_mch = '" & anaGrid.TextMatrix(dBariss, 22) & "'"
                            If rsPartMCH.RecordCount > 0 Then
    
                                rsPartMCH.Filter = adFilterNone
                                rsPartMCH.Filter = "part_mch = '" & anaGrid.TextMatrix(dBariss, 22) & "' and part_used='" & anaGrid.TextMatrix(dBariss, 2) & "'"
                                If rsPartMCH.RecordCount > 0 Then
                                    ProsesPresent dBariss, rsB("state_mach"), hkw1, c_cap_p_day
                                Else
                                    ProsesPresent dBariss, 0, hkw1, c_cap_p_day
                                End If
                            Else
                                ProsesPresent dBariss, rsB("state_mach"), hkw1, c_cap_p_day
                            End If
                        End If
    
                        ProgressBar1.Value = ProgressBar1.Value + 1
                    dob2 = dob2 + 1
                Loop
            Next
        Else
'            MsgBox "Sini 3"
            For dob = 1 To UBound(aPartPrior) - 1
'                MsgBox "sini 4"
                temp_part = aPartPrior(dob)
                rsB.Filter = adFilterNone
                rsB.Filter = "assy_no='" & aPartPrior(dob) & "'"
                rsB.Sort = "priorit ASC"
                dob2 = 1
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                        rsB.AbsolutePosition = dob2
                        c_wip = rsB("p1") + rsB("p2") + rsB("p3")
                        If rsB("ct") = 0 Then
                            c_cap_p_day = 0
                        Else
                            c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
                            If rsB("item_perbox") = 0 Then
                                If IsNull(rsB("item_muloq")) Then
                                Else
                                    c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
                                End If
                            Else
                                c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
                            End If
                        End If
                        hkw1 = phkw_tsb
                        presentMesinUse = rsB("presenku")
                        
                        If IsNull(rsB("mold_no")) Or rsB("mold_no") = "" Or rsB("mold_no") = "0" Then
                            nomold = "nomo"
                        Else
                            nomold = rsB("mold_no")
                        End If
                        
                        anaGrid.rows = anaGrid.rows + 1
                        dBariss = anaGrid.rows - 1
                        anaGrid.TextMatrix(dBariss, 0) = anaGrid.rows - 3
                        anaGrid.TextMatrix(dBariss, 1) = rsB(0)
                        anaGrid.TextMatrix(dBariss, 2) = RTrim(rsB(1))
                        anaGrid.TextMatrix(dBariss, 3) = rsB(2)
                        anaGrid.TextMatrix(dBariss, 4) = FormatNumber(rsB(3), 0)
                        anaGrid.TextMatrix(dBariss, 5) = c_wip
                        anaGrid.TextMatrix(dBariss, 6) = rsB("fc1")
    
                        If rsB("fc1") = 0 Then
                            anaGrid.TextMatrix(dBariss, 7) = 0
                        Else
                            anaGrid.TextMatrix(dBariss, 7) = (rsB(3) + c_wip) / rsB("fc1")
                        End If
    
                        anaGrid.TextMatrix(dBariss, 8) = IIf(IsNull(rsB("subcont")), "no", rsB("subcont")) 'kebijakan_subc
                        anaGrid.TextMatrix(dBariss, 9) = getProdplan(RTrim(rsB(1)), 1) 'rsB("prod_plan_1")
                        anaGrid.TextMatrix(dBariss, 10) = getProdplan(RTrim(rsB(1)), 2) 'rsB("prod_plan_2")
                        anaGrid.TextMatrix(dBariss, 11) = getProdplan(RTrim(rsB(1)), 3) 'rsB("prod_plan_3")
                        anaGrid.TextMatrix(dBariss, 12) = getProdplan(RTrim(rsB(1)), 4) 'rsB("prod_plan_4")
                        anaGrid.TextMatrix(dBariss, 13) = rsB("cavity")
                        anaGrid.TextMatrix(dBariss, 14) = rsB("ct")
                        anaGrid.TextMatrix(dBariss, 15) = rsB("manpower")
                        anaGrid.TextMatrix(dBariss, 16) = rsB("ct_2")
                        anaGrid.TextMatrix(dBariss, 17) = c_cap_p_day
                        anaGrid.TextMatrix(dBariss, 18) = c_cap_p_day
                        anaGrid.TextMatrix(dBariss, 19) = c_cap_p_day * hkw1
                        
                        
                        anaGrid.TextMatrix(dBariss, 21) = 0
                        anaGrid.TextMatrix(dBariss, 22) = IIf(IsNull(rsB("prod_nomach")), "nom", rsB("prod_nomach")) 'rsB("tonage_mach")
                        anaGrid.TextMatrix(dBariss, 23) = IIf(IsNull(rsB("tonage_mach")), "nom", rsB("tonage_mach"))
                        anaGrid.TextMatrix(dBariss, 24) = nomold 'IIf(IsNull(rsB("mold_no")) Or IsNumeric(rsB("mold_no")) = False, "nomo", rsB("mold_no"))
                        anaGrid.TextMatrix(dBariss, 25) = 0
    
                        anaGrid.TextMatrix(dBariss, 26) = IIf(IsNull(rsB("state_mach")), 0, rsB("state_mach")) 'IIf(ovrd_msnutama < 0, 0, ovrd_msnutama)
                        
                        anaGrid.TextMatrix(dBariss, 27) = 0
                        anaGrid.TextMatrix(dBariss, 30) = IIf(IsNull(rsB("faktor_productivity")), 0, rsB("faktor_productivity"))
                        anaGrid.TextMatrix(dBariss, 31) = rsB("hour_p_shift")
                        anaGrid.TextMatrix(dBariss, 32) = rsB("shift_usg")
                        If pi_hkw = 1 Then
                        anaGrid.TextMatrix(dBariss, 20) = wo_daymamo(rsB("prod_nomach"), nomold, RTrim$(rsB(1)))
                        anaGrid.TextMatrix(dBariss, 29) = anaGrid.TextMatrix(dBariss, 20)
                        anaGrid.TextMatrix(dBariss, 33) = wo_qtymamo(rsB("prod_nomach"), nomold, RTrim$(rsB(1)))
                        Else
                        anaGrid.TextMatrix(dBariss, 20) = 0
                        anaGrid.TextMatrix(dBariss, 29) = 0
                        anaGrid.TextMatrix(dBariss, 33) = 0
                        End If
                        anaGrid.TextMatrix(dBariss, 34) = IIf(IsNull(rsB("cavity_std")), 0, rsB("cavity_std"))
                        anaGrid.TextMatrix(dBariss, 35) = IIf(IsNull(rsB("timeupdate")), "", Format(rsB("timeupdate"), "yyyy-MM-dd HH:mm:ss"))
                        anaGrid.TextMatrix(dBariss, 36) = rsB("prc_prodplan")
                        ProgressBar1.Value = ProgressBar1.Value + 1
                    dob2 = dob2 + 1
                Loop
            Next
            Dim stmach As Byte
            With anaGrid
                For dob = 3 To .rows - 1
                    If .TextMatrix(dob, 26) = "0" Or .TextMatrix(dob, 24) = "nomo" Or .TextMatrix(dob, 24) = "0" Or .TextMatrix(dob, 8) = "yes" Then
                        If .TextMatrix(dob, 8) = "yes" Then '
                            ProsesPresent dob, 1, hkw1, .TextMatrix(dob, 17)
                        Else
                            ProsesPresent dob, 0, hkw1, .TextMatrix(dob, 17)
                        End If
                    Else
                        If .TextMatrix(dob, 8) = "yes" Then
                            stmach = 1
                        Else
                            stmach = 1
                        End If
                        rsPartMCH.Filter = adFilterNone
                        rsPartMCH.Filter = "part_mch = '" & .TextMatrix(dob, 22) & "'"
                        
                        If rsPartMCH.RecordCount > 0 Then
                            rsPartMCH.Filter = adFilterNone
                            rsPartMCH.Filter = "part_mch = '" & .TextMatrix(dob, 22) & "' and part_used='" & .TextMatrix(dob, 2) & "'"
                            If rsPartMCH.RecordCount > 0 Then
                                ProsesPresent dob, stmach, hkw1, .TextMatrix(dob, 17) ' c_cap_p_day
                            Else
                                ProsesPresent dob, 0, hkw1, .TextMatrix(dob, 17)
                            End If
                        Else
                            ProsesPresent dob, stmach, hkw1, .TextMatrix(dob, 17)
                        End If
                    End If
                Next
            End With
        End If
        '&&& end ane
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False

    Else
        anaGrid.rows = 3
    End If
    For j = 3 To anaGrid.rows - 1
        For x = 1 To UBound(nm_msn_full)
            If nm_msn_full(x) = anaGrid.TextMatrix(j, 22) Then
                anaGrid.Col = 22
                anaGrid.Row = j
                anaGrid.CellBackColor = RGB(255, 255, 62)
            End If
        Next
    Next
    '**periksa yang OVERLOAD hanya 1 hari
    '*** uniq part no, mold, mesin
    'FILTER1
    For dBariss = 1 To UBound(c_part)
        If ar_prodplan(dBariss) * 1 > 0 Then
            rsB.Filter = adFilterNone
            rsB.Filter = "assy_no='" & c_part(dBariss) & "' and subcont='no'"
            rsB.Sort = "mold_no ASC"
            ttMold = ""
            tToutalMold = 0
            'Periksa mold 1
            For dob2 = 1 To rsB.RecordCount
                rsB.AbsolutePosition = dob2
                If ttMold <> rsB("mold_no") Then
                    tToutalMold = tToutalMold + 1
                    ttMold = rsB("mold_no")
                End If
            Next
            rsB.Sort = ""
            dob2 = 1
            If (tToutalMold) <= 1 Then '<=
                dob2c = 0
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                    rsB.AbsolutePosition = dob2
                    c_NDMtZ = NDMtZ(rsB("prod_nomach"), rsB("assy_no"))
                    If rsB("state_mach") = 1 And c_NDMtZ > 0 Then
                        dob2c = dob2
                    End If
                    dob2 = dob2 + 1
                Loop
                If dob2c > 0 Then
                    rsB.AbsolutePosition = dob2c
                    If rsB("ct") = 0 Then
                        c_cap_p_day = 0
                    Else
                        c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
                        If rsB("item_perbox") = 0 Then
                            c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
                        Else
                            c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
                        End If
                    End If
                    need_day = FormatNumber(ar_prodplan(dBariss) / c_cap_p_day, 2)
                    If need_day * 1 <= 1 And need_day * 1 > 0 Then
                        If PerMcNow(rsB("prod_nomach"), hkw1) + (need_day / hkw1 * 100) <= 105 Then
                            PlotSISA c_part(dBariss), rsB("mold_no"), need_day, rsB("prod_nomach")
                        End If
                    End If
                Else
                    dob2 = 1
                    Do
                        If dob2 > rsB.RecordCount Then Exit Do
                        rsB.AbsolutePosition = dob2
                        If rsB("state_mach") = 1 Then
                            If rsB("ct") = 0 Then
                                c_cap_p_day = 0
                            Else
                                c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
                                If rsB("item_perbox") = 0 Then
                                    c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
                                Else
                                    c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
                                End If
                            End If
                            need_day = FormatNumber(ar_prodplan(dBariss) / c_cap_p_day, 2)
                            
                            If need_day * 1 <= 1 And need_day * 1 > 0 Then
                                If PerMcNow(rsB("prod_nomach"), hkw1) + (need_day / hkw1 * 100) <= 105 Then
                                    If blockSpec(rsB("prod_nomach"), c_part(dBariss)) Then
                                        PlotSISA c_part(dBariss), rsB("mold_no"), need_day, rsB("prod_nomach")
                                    End If
                                End If
                            End If
                        End If
                        dob2 = dob2 + 1
                    Loop
                End If
            Else
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                    rsB.AbsolutePosition = dob2
                    If rsB("state_mach") = 1 Then
                        If rsB("ct") = 0 Then
                            c_cap_p_day = 0
                        Else
                            c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
                            If rsB("item_perbox") = 0 Then
                                c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
                            Else
                                c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
                            End If
                        End If
                        need_day = FormatNumber(ar_prodplan(dBariss) / c_cap_p_day, 2)
                        If need_day * 1 <= 1 And need_day * 1 > 0 Then
                            If PerMcNow(rsB("prod_nomach"), hkw1) + (need_day / hkw1 * 100) <= 105 Then
                                If blockSpec(rsB("prod_nomach"), c_part(dBariss)) Then
                                    PlotSISA c_part(dBariss), rsB("mold_no"), need_day, rsB("prod_nomach")
                                End If
'                                rsKeArray2 c_part(dBariss), need_day, rsB("mold_no"), rsB("prod_nomach")
                            End If
                        End If
                    End If
                    dob2 = dob2 + 1
                Loop
            End If
        End If
    Next
    
    '-----------------rekap need day
    For x = 3 To anaGrid.rows - 1
        For j = 3 To anaGrid.rows - 1
            If anaGrid.TextMatrix(x, 2) = anaGrid.TextMatrix(j, 2) Then
                If Val(anaGrid.TextMatrix(x, 20)) > 0 Then
                    anaGrid.TextMatrix(x, 21) = Val(anaGrid.TextMatrix(x, 21)) + Val(anaGrid.TextMatrix(j, 20))
                End If
            End If
        Next
    Next

    '--------sum perbulan load cap
    For x = 3 To anaGrid.rows - 1
        anaGrid.TextMatrix(x, 27) = Val(anaGrid.TextMatrix(x, 20)) / hkw1 * 100
        anaGrid.TextMatrix(x, 28) = (Val(anaGrid.TextMatrix(x, 20)) / hkw1) * Val(anaGrid.TextMatrix(x, 15))
    Next
    gridFormatNum
    settingHeaderMonth
    anaGrid.Refresh
End Sub

Private Sub PlotSISA(pPART As String, pMOLD As String, pHari As Variant, pmesin As String)
    'bangun
    Dim j As Long
    With anaGrid
        For j = 3 To .rows - 1
            .Col = 2
            If .TextMatrix(j, 2) = pPART And .TextMatrix(j, 24) = pMOLD And .TextMatrix(j, 22) = pmesin Then
                .Row = j
                .CellBackColor = vbGreen
                .TextMatrix(j, 20) = pHari * 1 + (.TextMatrix(j, 20) * 1)
                .TextMatrix(j, 33) = (.TextMatrix(j, 17) * 1) * (.TextMatrix(j, 20) * 1)
                setProdPlan0 pPART
                kurangihariMesin pmesin, pHari
            End If
        Next
    End With
    'end bangun
End Sub

Private Function PerMcNow(pmesin As String, phkw As Integer) As Single
    Dim xi As Long
    Dim xiv As Single
    xiv = 0
    For xi = 3 To anaGrid.rows - 1
        If anaGrid.TextMatrix(xi, 22) = pmesin Then
            xiv = xiv + anaGrid.TextMatrix(xi, 20) / phkw * 100
        End If
    Next
    PerMcNow = xiv
End Function

Private Sub CMdNExt_Click()
    If Len(Text1) < 1 Then Exit Sub
    Screen.MousePointer = 11
    Dim u As Integer, posisiBulan As Long, hkw_tsb As Integer
    For u = 1 To UBound(ar_hkw)
        If ar_hkw_bln(u) = Text1 Then
            posisiBulan = u
            Exit For
        End If
    Next
    If posisiBulan <= 4 Then
        posisiBulan = posisiBulan + 1
        If posisiBulan > 4 Then posisiBulan = 4
        Text1 = ar_hkw_bln(posisiBulan)
    End If
    '---------------cari HKW dengan diketahui bulan--------------
    For u = 1 To UBound(ar_hkw)
        If ar_hkw_bln(u) = Text1 Then
             hkw_tsb = ar_hkw(u)
             posisiBulan = u
            Exit For
        End If
    Next
    SkinLabel4.Caption = "HKW [" & hkw_tsb & "]"
    NOP_Generate hkw_tsb, posisiBulan
    Screen.MousePointer = 0
    If CmbDocument <> "" Then checkNeedMoldMachine
End Sub

Private Sub formatWarnaBG()
    Dim j As Integer
    For j = 3 To anaGrid.rows - 1
        anaGrid.Col = 22
        anaGrid.Row = j
        anaGrid.CellBackColor = vbWhite
    Next
End Sub



Private Sub checkNeedMoldMachine()
    List1.Clear
    Dim r As Integer
    For r = 1 To UBound(c_part)
        If Val(ar_prodplan(r)) > 0 Then
            List1.AddItem c_part(r) & " Overload"
        End If
    Next
    If List1.ListCount > 0 Then
        List1.Visible = True
    Else
        List1.Visible = False
    End If
End Sub

Private Function nmBulankeAngka(pis As String) As String
    Dim x As Integer
    For x = 1 To UBound(nmbulan)
        If pis = nmbulan(x) Then
            nmBulankeAngka = x
            Exit For
        End If
    Next
End Function

Private Function checkHeaderSaved(pPART As String) As String
    Dim a As Integer
    For a = 1 To UBound(c_part)
        If c_part(a) = pPART Then
            checkHeaderSaved = c_part_saved(a)
            Exit For
        End If
    Next
End Function

Private Sub setHeaderSaved(pPART As String)
    Dim a As Integer
    For a = 1 To UBound(c_part)
        If c_part(a) = pPART Then
            c_part_saved(a) = "1"
            Exit For
        End If
    Next
End Sub

Private Sub setHeaderSavedReset()
    Dim a As Integer
    For a = 1 To UBound(c_part)
        c_part_saved(a) = 0
    Next
End Sub

Private Function checkLoadCapSaved() As Boolean
    qry = "select fltpp_rev from mpp_gen_d where fltpp_doc ='" & CmbDocument & "' and fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "' " _
        & " order by fltpp_rev desc " _
        & "limit 1"
    Set RsA = Con.Execute(qry)
    If RsA.RecordCount > 0 Then
        nextRev = RsA("fltpp_rev") + 1
        checkLoadCapSaved = True
    Else
        nextRev = 0
        checkLoadCapSaved = False
    End If
End Function

Private Sub cmdSave_Click()
'On Error GoTo excEp
    Dim i As Long, u As Long, totalBaris As Long
    Dim rsLC_d As ADODB.Recordset
    If MsgBox("Are you sure want to save the data ?", vbQuestion + vbYesNo) = vbYes Then
        If checkLoadCapSaved Then
            If MsgBox("You have entered the data." & vbNewLine _
             & "do you want to re-generate data ?", vbQuestion + vbYesNo) = vbYes Then
            Else
                Exit Sub
            End If
        End If
        Screen.MousePointer = 11
        Picture1.Visible = True
        Picture1.Refresh
        Set rsLC_d = New ADODB.Recordset
        
        rsLC_d.Open "select * from mpp_gen_d where fltpp_doc='" & CmbDocument & "' and fltpp_rev = " & nextRev, Con, adOpenDynamic, adLockOptimistic
        
        ProgressBar2.Value = 0
        '----------reset END
        
        ym = DTPicker1.Value
        With anaGrid
        For u = 1 To 1 'UBound(ar_hkw)
'            setHeaderSavedReset
            NOP_Generate Val(ar_hkw(u)), u
            totalBaris = anaGrid.rows - 3
            For i = 3 To anaGrid.rows - 1
                rsLC_d.AddNew
                rsLC_d!lcd_itemdid = .TextMatrix(i, 2)
                rsLC_d!lc_customer = .TextMatrix(i, 1)
                rsLC_d!lc_subcont = .TextMatrix(i, 8)
                rsLC_d!no_mach = .TextMatrix(i, 22)
                rsLC_d!ton_mach = Val(.TextMatrix(i, 23))
                rsLC_d!reg_mold = .TextMatrix(i, 24)
                rsLC_d!fltpp_period = Format(DTPicker1.Value, "yyyyMM")
                rsLC_d!cav = .TextMatrix(i, 13) * 1
                rsLC_d!ct = .TextMatrix(i, 14) * 1
                rsLC_d!mpower = .TextMatrix(i, 15) * 1
                rsLC_d!ct_scnd = .TextMatrix(i, 16) * 1
                rsLC_d!cap_p_day = .TextMatrix(i, 17) * 1
                rsLC_d!neday = .TextMatrix(i, 20) * 1
                rsLC_d!sum_nedady = .TextMatrix(i, 21) * 1
                rsLC_d!lcvsmach = Left(.TextMatrix(i, 27), Len(.TextMatrix(i, 27)) - 2) * 1
                rsLC_d!lcneed_mp = .TextMatrix(i, 28) * 1
                rsLC_d!fltpp_doc = CmbDocument
                rsLC_d!fltpp_rev = nextRev
                rsLC_d!fltpp_ym = Format(ym, "yyyymm") ' Format(DTPicker1, "yyyy") & Right("00" & nmBulankeAngka(ar_hkw_bln(u)), 2)
                rsLC_d!rstate_mach = .TextMatrix(i, 26)
                rsLC_d!fltpp_hkw = ar_hkw(u)
                rsLC_d!lc_sisa_pp = checkSisaProdplan(.TextMatrix(i, 2))
                rsLC_d!lc_pp = .TextMatrix(i, 8 + u) * 1
                rsLC_d!lc_fprodtvty = .TextMatrix(i, 30)
                rsLC_d!lc_itemname = .TextMatrix(i, 3)
                rsLC_d!lc_stockqty = .TextMatrix(i, 4) * 1
                rsLC_d!lc_stockwip = .TextMatrix(i, 5)
                rsLC_d!lc_fc = .TextMatrix(i, 6) * 1
                rsLC_d!lc_ito = .TextMatrix(i, 7) * 1
                rsLC_d!shiftusg = .TextMatrix(i, 32) * 1
                rsLC_d!hourpshift = .TextMatrix(i, 31) * 1
                If Len(.TextMatrix(i, 33)) > 0 Then
                    rsLC_d!neqty = .TextMatrix(i, 33) * 1
                Else
                    rsLC_d!neqty = 0
                End If
                If Len(.TextMatrix(i, 34)) > 0 Then
                    rsLC_d!cav_std = .TextMatrix(i, 34)
                Else
                    rsLC_d!cav_std = 0
                End If
                If .TextMatrix(i, 35) <> "" Then
                    rsLC_d!timeupdate = .TextMatrix(i, 35)
                End If
                rsLC_d!prc_prodplan = .TextMatrix(i, 36) * 1
                rsLC_d.Update

                ProgressBar2.Value = FormatNumber(((i - 3) * 100) / totalBaris, 0)
                Label1.Caption = "Generating data..." & ProgressBar2.Value & "%"
                Label1.Refresh
            Next
            ym = DateAdd("m", 1, ym)
        Next
        Screen.MousePointer = 0
        Set rsLC_d = Nothing
        MsgBox "tersimpan"
        End With
    End If
    Exit Sub
'excEp:
 '   MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub Command1_Click()
    If Picture1.Visible Then
        Picture1.Visible = False
    Else
        Picture1.Visible = True
    End If
End Sub

Private Sub Command2_Click()
    Dim xf As Double, pos As Integer
    Dim ttlrows As Double
    Dim stringCari As String
    Dim kolom As Byte
    If cmbKol.ListIndex = 0 Then
        kolom = 2
    Else
        kolom = 22
    End If
    With anaGrid
        ttlrows = .rows - 1
        If posisisFind + 1 >= ttlrows Then
            posisisFind = 2
        Else
            posisisFind = 1 + posisisFind
        End If
        For xf = posisisFind To ttlrows
            stringCari = LCase$(.TextMatrix(xf, kolom))
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

Private Sub DTPicker1_Change()
    Picture1.Visible = False
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

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingFG
    DTPicker1.Value = Now
    Me.Height = 7755
    Me.Width = 14640
'    Call WheelHook(Me.hwnd)

    ReDim nmbulan(1 To 12) As String
    nmbulan(1) = "January"
    nmbulan(2) = "February"
    nmbulan(3) = "March"
    nmbulan(4) = "April"
    nmbulan(5) = "May"
    nmbulan(6) = "June"
    nmbulan(7) = "July"
    nmbulan(8) = "August"
    nmbulan(9) = "September"
    nmbulan(10) = "October"
    nmbulan(11) = "November"
    nmbulan(12) = "December"
    cmbKol.ListIndex = 0
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    CmbDocument.Left = cmdPrev.Left
    CmbDocument.Top = SkinLabel3.Top
    txtRevision.Top = SkinLabel2.Top
    txtRevision.Left = DTPicker1.Left
    cmbKol.Width = Label2.Width
    cmbKol.Top = txtFindNext.Top
    cmbKol.Left = Label2.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Call WheelUnHook(Me.hwnd)
        DelTab Me
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

Private Sub List1_Click()
    MsgBox List1.ListCount
End Sub

Private Sub rsKeArray(paray As String)
    Dim ada As Boolean
    For dob = 1 To UBound(aPartPrior)
        If aPartPrior(dob) = paray Then
            ada = True
            Exit For
        Else
            ada = False
        End If
    Next
    If ada = False Then
        aPartPrior(UBound(aPartPrior)) = paray
        ReDim Preserve aPartPrior(1 To UBound(aPartPrior) + 1) As String
    End If
End Sub

Private Sub rsKeArray2(pitem As String, pneday As Variant, pMOLD As String, pmesin As String)
    Dim ada As Boolean, r As Integer
    For dob = 1 To UBound(ar_PartSisa)
        If ar_PartSisa(dob) = pitem Then
            ada = True
            r = dob
            Exit For
        Else
            ada = False
        End If
    Next
    If ada = False Then
        ar_PartSisa(UBound(ar_PartSisa)) = pitem
        ReDim Preserve ar_PartSisa(1 To UBound(ar_PartSisa) + 1) As String
        ar_Sisa(UBound(ar_Sisa)) = pneday
        ReDim Preserve ar_Sisa(1 To UBound(ar_Sisa) + 1) As Variant
        ar_MesinSisa(UBound(ar_MesinSisa)) = pmesin
        ReDim Preserve ar_MesinSisa(1 To UBound(ar_MesinSisa) + 1) As String
        ar_MoldSisa(UBound(ar_MoldSisa)) = pMOLD
        ReDim Preserve ar_MoldSisa(1 To UBound(ar_MoldSisa) + 1) As String
    Else
        If ar_Sisa(r) * 1 >= pneday * 1 Then
            ar_Sisa(r) = pneday
            ar_PartSisa(r) = pitem
            ar_MoldSisa(r) = pMOLD
            ar_MesinSisa(r) = pmesin
        End If
    End If
End Sub

Private Function isi(pMPQ As Double, pCapPDay As Variant, atasBawah As String)
    Dim MPQ As Variant
    Dim bReach As Boolean
    
    bReach = True
    MPQ = pMPQ
    If pMPQ = 0 Then
        MsgBox "Packing " & temp_part & "=0"
        Exit Function
    End If
    While bReach
        If MPQ * 1 > pCapPDay * 1 Then
            If atasBawah = "a" Then
                isi = MPQ '- pMPQ
            Else
                isi = MPQ - pMPQ
            End If
            bReach = False
        Else
            If MPQ = pCapPDay Then
                isi = pCapPDay
                bReach = False
            Else
                isi = MPQ
            End If
        End If
        MPQ = MPQ * 1 + pMPQ * 1
    Wend
End Function

Private Sub WORun()
    ym = DTPicker1
    qry = "select mesinno,moldno,partno,sum(qty) ttlrun,sum(qty)/max(cappday) dayrun from worko " _
    & " where extract(YEAR from issudate)=" & Format(ym, "yyyy") & " and extract(month from issudate)=" & Format(ym, "MM") _
    & " group by mesinno,moldno,partno order by 1 asc"
    Set rsW = Con.Execute(qry)
    If rsW.RecordCount > 0 Then
        rsW.Fields("mesinno").Properties("Optimize") = True
        rsW.Fields("dayrun").Properties("Optimize") = True
        rsW.Fields("moldno").Properties("Optimize") = True
        rsW.Fields("partno").Properties("Optimize") = True
    End If
    qry = "select partno,sum(qty) ttlrun,max(mpp_pp) pp from worko " _
    & " where extract(YEAR from issudate)=" & Format(ym, "yyyy") & " and extract(month from issudate)=" & Format(ym, "MM") _
    & " group by partno order by 1 asc"
    Set rsX = Con.Execute(qry)
    If rsX.RecordCount > 0 Then
        rsX.Fields("partno").Properties("Optimize") = True
        rsX.Fields("ttlrun").Properties("Optimize") = True
    End If
End Sub

Private Function wo_daymamo(pMCH As String, pmld As String, pPART As String) As Single
    rsW.Filter = "mesinno='" & pMCH & "' and moldno='" & pmld & "' and partno='" & pPART & "'"
    If rsW.RecordCount > 0 Then
        wo_daymamo = rsW("dayrun")
    Else
        wo_daymamo = 0
    End If
End Function

Private Function wo_qtymamo(pMCH As String, pmld As String, pPART As String) As Long
    rsW.Filter = "mesinno='" & pMCH & "' and moldno='" & pmld & "' and partno='" & pPART & "'"
    If rsW.RecordCount > 0 Then
        wo_qtymamo = rsW("ttlrun")
    Else
        wo_qtymamo = 0
    End If
End Function

Private Function wo_prodplan(pPART As String) As Long
    rsX.Filter = "partno='" & pPART & "'"
    If rsX.RecordCount > 0 Then
        wo_prodplan = rsX("ttlrun")
    Else
        wo_prodplan = 0
    End If
    rsX.Filter = adFilterNone
End Function

Private Function wo_runeqpp(pPART As String) As Boolean
    rsX.Filter = "partno='" & pPART & "'"
    If rsX.RecordCount > 0 Then
        If rsX("ttlrun") = rsX("pp") Then
            wo_runeqpp = True
        Else
            wo_runeqpp = False
        End If
    Else
        wo_runeqpp = False
    End If
    rsX.Filter = adFilterNone
End Function

Private Function wo_hkwmold(mold As String) As Single
    Dim ttldayrun As Single
    ttldayrun = 0
    rsW.Filter = "moldno='" & mold & "'"
    If rsW.RecordCount > 0 Then
        If rsW.RecordCount = 1 Then
            wo_hkwmold = rsW("dayrun")
        Else
            rsW.MoveFirst
            While Not rsW.EOF
                ttldayrun = ttldayrun + rsW("dayrun")
                rsW.MoveNext
            Wend
            wo_hkwmold = ttldayrun
        End If
    Else
        wo_hkwmold = 0
    End If
End Function

Private Function wo_hkwmachine(mch As String) As Single
    Dim ttldayrun As Single
    ttldayrun = 0
    rsW.Filter = "mesinno='" & mch & "'"
    If rsW.RecordCount > 0 Then
        If rsW.RecordCount = 1 Then
            wo_hkwmachine = rsW("dayrun")
        Else
            rsW.MoveFirst
            While Not rsW.EOF
                ttldayrun = ttldayrun + rsW("dayrun")
                rsW.MoveNext
            Wend
            wo_hkwmachine = ttldayrun
        End If
    Else
        wo_hkwmachine = 0
    End If
End Function

Private Sub txtFindNext_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2_Click
    ElseIf KeyAscii = vbKeyEscape Then
        PicFIND.Visible = False
    ElseIf KeyAscii = 1 Then
        txtFindNext.SelStart = 0
        txtFindNext.SelLength = Len(txtFindNext.Text)
    End If
End Sub

Private Sub txtRevision_Click()
    Screen.MousePointer = 11
    formatWarnaBG
    qry = "SELECT part_used,part_mch from loadcap_partused"
    Set rsPartMCH = Con.Execute(qry)
    rsPartMCH.Fields("part_mch").Properties("Optimize") = True
    Dim i As Integer, c_wip As Variant, c_cap_p_day As Variant, j As Integer, x As Integer
    Dim presentMesinUse As Variant
    Dim nomold As String
    Dim bulanHKW As String
    Dim tahunHKW As String, ttlKoma As Integer
    Dim hkw1 As Integer
    Dim dateFlag As Date
    Dim finPP As Double
    '--------------GET WO RU
    Call WORun
    
    '-------------====================START HKW DARI SETTOFF===============----------
    dateFlag = DTPicker1
    For i = 1 To 4
        If Format(dateFlag, "yyyy") <= Format(DTPicker1, "yyyy") Then
            bulanHKW = bulanHKW & Format(dateFlag, "m") & ","
            tahunHKW = Format(dateFlag, "yyyy")
        End If
        dateFlag = DateAdd("m", 1, dateFlag)
    Next
    
    bulanHKW = Left(bulanHKW, Len(bulanHKW) - 1)
    qry = "select max(work_date) tgl, count(work_status) hkw from plansys_setoffday " _
            & " where  work_status=true and extract(month from work_date) in (" & bulanHKW & ") and extract(year from work_date)=" & tahunHKW _
            & " group by extract(month from work_date) , extract(year from work_date)" _
            & " order by extract(year from work_date) asc,extract(month from work_date) asc"
    Set rsB = Con.Execute(qry)
    If rsB.RecordCount > 0 Then
        i = 1
        While Not rsB.EOF
            ar_hkw(i) = rsB("hkw")
            rsB.MoveNext
            i = i + 1
        Wend
    End If
    For i = 1 To Len(bulanHKW)
        If Mid(bulanHKW, i, 1) = "," Then
            ttlKoma = ttlKoma + 1
        End If
    Next
    If ttlKoma <> 3 Then
        bulanHKW = ""
        dateFlag = DateAdd("m", ttlKoma + 1, DTPicker1)
        For i = ttlKoma + 2 To 4
           bulanHKW = bulanHKW & Format(dateFlag, "m") & ","
           tahunHKW = Format(dateFlag, "yyyy")
           dateFlag = DateAdd("m", 1, dateFlag)
        Next
        bulanHKW = Left(bulanHKW, Len(bulanHKW) - 1)
        
        qry = "select max(work_date) tgl, count(work_status) hkw from plansys_setoffday " _
            & " where  work_status=true and extract(month from work_date) in (" & bulanHKW & ") and extract(year from work_date)=" & tahunHKW _
            & " group by extract(month from work_date) , extract(year from work_date)" _
            & " order by extract(year from work_date) asc,extract(month from work_date) asc"
        Set rsB = Con.Execute(qry)
        If rsB.RecordCount > 0 Then
            i = ttlKoma + 2
            While Not rsB.EOF
                ar_hkw(i) = rsB("hkw")
                rsB.MoveNext
                i = i + 1
            Wend
        End If
    End If
   
    '---------------------===========END HKW DARI SETOFF
    
    qry = "select distinct on (assy_no) cust_name,assy_no" _
         & ",prod_plan_1,prod_plan_2,prod_plan_3,prod_plan_4,g.ct,g.cavity,hour_p_shift,shift_usg,faktor_productivity,item_muloq,item_perbox " _
         & ", prc_prodplan from ltpp_generate a " _
         & "inner join mst_item b on a.assy_no=b.item_id " _
         & "inner join r_customer c on b.cust_id=c.cust_id " _
         & "inner join loadcap_mst_product_r d on a.assy_no=d.partno " _
         & "inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc " _
         & "inner join loadcap_proc g on d.partno=g.partno " _
         & "left join loadcap_mst_mach e on g.prod_nomach=e.no_mach " _
         & "where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='') "

    Set rsB = Con.Execute(qry)
    i = 1
    If rsB.RecordCount > 0 Then
        ReDim c_part(1 To rsB.RecordCount) As String
        ReDim c_part_saved(1 To rsB.RecordCount) As String
        ReDim ar_prodplan(1 To rsB.RecordCount) As Variant
        ReDim ar_propl(1 To rsB.RecordCount) As Variant
        ReDim ar_propl2(1 To rsB.RecordCount) As Variant
        ReDim ar_propl3(1 To rsB.RecordCount) As Variant
        ReDim ar_propl4(1 To rsB.RecordCount) As Variant
        While Not rsB.EOF
            temp_part = RTrim(rsB("assy_no"))
            c_part(i) = RTrim(rsB("assy_no"))
            c_part_saved(i) = "0"
            If rsB("ct") = 0 Then
                c_cap_p_day = 0
            Else
                c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
            End If
            c_cap_p_day = FormatNumber(c_cap_p_day, 0) * 1
            finPP = rsB("prod_plan_1") * rsB("prc_prodplan") / 100
            If c_cap_p_day * 1 > finPP Then
                If finPP > 0 Then
                    If rsX.RecordCount = 0 Then
                        If rsB("item_perbox") = 0 Then
                            ar_prodplan(i) = isi(rsB("item_muloq"), c_cap_p_day, "b")
                            ar_propl(i) = isi(rsB("item_muloq"), c_cap_p_day, "b")
                        Else
                            ar_prodplan(i) = isi(rsB("item_perbox"), c_cap_p_day, "b")
                            ar_propl(i) = isi(rsB("item_perbox"), c_cap_p_day, "b")
                        End If
                    Else
                        If wo_runeqpp(Trim(rsB("assy_no"))) Then
                            ar_prodplan(i) = 0
                            ar_propl(i) = wo_prodplan(Trim(rsB("assy_no")))
                        Else
                            If rsB("item_perbox") = 0 Then
                                ar_prodplan(i) = isi(rsB("item_muloq"), c_cap_p_day, "b") - wo_prodplan(Trim(rsB("assy_no")))
                                ar_propl(i) = isi(rsB("item_muloq"), c_cap_p_day, "b") ' - wo_prodplan(Trim(rsB("assy_no")))
                            Else
                                ar_prodplan(i) = isi(rsB("item_perbox"), c_cap_p_day, "b") - wo_prodplan(Trim(rsB("assy_no")))
                                ar_propl(i) = isi(rsB("item_perbox"), c_cap_p_day, "b") ' - wo_prodplan(Trim(rsB("assy_no")))
                            End If
                        End If
                    End If
                    
                Else
                    ar_prodplan(i) = 0
                    ar_propl(i) = 0
                End If
            Else
                If finPP > 0 Then
                    If rsX.RecordCount = 0 Then
                        If rsB("item_perbox") = 0 Then
                            ar_prodplan(i) = isi(rsB("item_muloq"), finPP, "b")
                            ar_propl(i) = isi(rsB("item_muloq"), finPP, "b")
                        Else
                            ar_prodplan(i) = isi(rsB("item_perbox"), finPP, "b")
                            ar_propl(i) = isi(rsB("item_perbox"), finPP, "b")
                        End If
                    Else
                        If rsB("item_perbox") = 0 Then
                            ar_prodplan(i) = isi(rsB("item_muloq"), finPP, "b") - wo_prodplan(Trim(rsB("assy_no")))
                            ar_propl(i) = isi(rsB("item_muloq"), finPP, "b") '- wo_prodplan(Trim("assy_no"))
                        Else
                            ar_prodplan(i) = isi(rsB("item_perbox"), finPP, "b") - wo_prodplan(Trim(rsB("assy_no")))
                            ar_propl(i) = isi(rsB("item_perbox"), finPP, "b") '- wo_prodplan(Trim("assy_no"))
                        End If
                    End If
                   
                Else
                    ar_prodplan(i) = 0
                    ar_propl(i) = 0
                End If
            End If
            
            '* prod plan 2
            If c_cap_p_day * 1 > rsB("prod_plan_2") Then
                If rsB("prod_plan_2") > 0 Then
                    If rsB("item_perbox") = 0 Then
                        ar_propl2(i) = isi(rsB("item_muloq"), c_cap_p_day, "b")
                    Else
                        ar_propl2(i) = isi(rsB("item_perbox"), c_cap_p_day, "b")
                    End If
                Else
                    ar_propl2(i) = 0
                End If
            Else
                If rsB("prod_plan_2") > 0 Then
                    If rsB("item_perbox") = 0 Then
                        ar_propl2(i) = isi(rsB("item_muloq"), rsB("prod_plan_2"), "b")
                    Else
                        ar_propl2(i) = isi(rsB("item_perbox"), rsB("prod_plan_2"), "b")
                    End If
                Else
                    ar_propl2(i) = 0
                End If
            End If
            
            '* prod plan 3
            If c_cap_p_day * 1 > rsB("prod_plan_3") Then
                If rsB("prod_plan_3") > 0 Then
                    If rsB("item_perbox") = 0 Then
                        ar_propl3(i) = isi(rsB("item_muloq"), c_cap_p_day, "b")
                    Else
                        ar_propl3(i) = isi(rsB("item_perbox"), c_cap_p_day, "b")
                    End If
                Else
                    ar_propl3(i) = 0
                End If
            Else
                If rsB("prod_plan_3") > 0 Then
                    If rsB("item_perbox") = 0 Then
                        ar_propl3(i) = isi(rsB("item_muloq"), rsB("prod_plan_3"), "b")
                    Else
                        ar_propl3(i) = isi(rsB("item_perbox"), rsB("prod_plan_3"), "b")
                    End If
                Else
                    ar_propl3(i) = 0
                End If
            End If
            '* prod plan 4
            If c_cap_p_day * 1 > rsB("prod_plan_4") Then
                If rsB("prod_plan_4") > 0 Then
                    If rsB("item_perbox") = 0 Then
                        ar_propl4(i) = isi(rsB("item_muloq"), c_cap_p_day, "b")
                    Else
                        ar_propl4(i) = isi(rsB("item_perbox"), c_cap_p_day, "b")
                    End If
                Else
                    ar_propl4(i) = 0
                End If
            Else
                If rsB("prod_plan_4") > 0 Then
                    If rsB("item_perbox") = 0 Then
                        ar_propl4(i) = isi(rsB("item_muloq"), rsB("prod_plan_4"), "b")
                    Else
                        ar_propl4(i) = isi(rsB("item_perbox"), rsB("prod_plan_4"), "b")
                    End If
                Else
                    ar_propl4(i) = 0
                End If
            End If
            i = i + 1
            rsB.MoveNext
        Wend
    Else
        ar_hkw(1) = "-"
        ar_hkw(2) = "-"
        ar_hkw(3) = "-"
        ar_hkw(4) = "-"
        ar_hkw_bln(1) = "-"
        ar_hkw_bln(2) = "-"
        ar_hkw_bln(3) = "-"
        ar_hkw_bln(4) = "-"
        ReDim c_part(1) As String
        ReDim c_part_saved(1) As String
        ReDim ar_prodplan(1) As Variant
        ReDim ar_propl(1) As Variant
        ReDim ar_propl2(1) As Variant
        ReDim ar_propl3(1) As Variant
        ReDim ar_propl4(1) As Variant
    End If
    
    '----------------mendapatkan mesin yang dipakai (distict) GEUR-----------------,a.hkw_1
    qry = "select distinct on (prod_nomach) prod_nomach " _
         & " from ltpp_generate a " _
         & " inner join mst_item b on a.assy_no=b.item_id " _
         & " inner join r_customer c on b.cust_id=c.cust_id " _
         & " inner join loadcap_mst_product_r d on a.assy_no=d.partno " _
         & " inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc " _
         & " inner join loadcap_proc g on d.partno=g.partno " _
         & " left join loadcap_mst_mach e on g.prod_nomach=e.no_mach " _
         & " where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='')"
    Set rsB = Con.Execute(qry)
    
    i = 1
    If rsB.RecordCount > 0 Then
        ReDim ar_kpsts_mc_hr(1 To rsB.RecordCount) As String
        ReDim ar_nm_msn(1 To rsB.RecordCount) As String
        While Not rsB.EOF
            ar_kpsts_mc_hr(i) = ar_hkw(1) - wo_hkwmachine(rsB(0))
            ar_nm_msn(i) = IIf(IsNull(rsB(0)), "nom", rsB(0))
            i = i + 1
            rsB.MoveNext
        Wend
    Else
        ReDim ar_kpsts_mc_hr(1) As String
        ReDim ar_nm_msn(1) As String
    End If
    
    '---------------------KAPASITAS HARI MOLD-------------,a.hkw_1
    qry = "select distinct on (mold_no) mold_no  " _
          & " from ltpp_generate a " _
          & " inner join mst_item b on a.assy_no=b.item_id " _
          & " inner join r_customer c on b.cust_id=c.cust_id " _
          & " inner join loadcap_mst_product_r d on a.assy_no=d.partno " _
          & " inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc " _
          & " inner join loadcap_proc g on d.partno=g.partno " _
          & " left join loadcap_mst_mach e on g.prod_nomach=e.no_mach " _
          & " where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='')"
    Set rsB = Con.Execute(qry)
    
    i = 1
    If rsB.RecordCount > 0 Then
        ReDim ar_kpsts_mold_hr(1 To rsB.RecordCount) As String
        ReDim ar_nm_mold(1 To rsB.RecordCount) As String
        While Not rsB.EOF
            ar_kpsts_mold_hr(i) = ar_hkw(1) - wo_hkwmold(rsB(0))
            ar_nm_mold(i) = IIf(IsNull(rsB(0)), "nomo", rsB(0))
            i = i + 1
            rsB.MoveNext
        Wend
    Else
        ReDim ar_kpsts_mold_hr(1) As String
        ReDim ar_nm_mold(1) As String
    End If

    '-----------------------------YUK KITA PIKIRKAN---------------------------
    i = 0
    qry = "select cust_name,assy_no,a.item_name,fg,p1,p2,p3,fc1 " _
        & " ,prod_plan_1,prod_plan_2,prod_plan_3,prod_plan_4 " _
        & " ,g.cavity,g.ct,g.manpower,g.ct_2,g.prod_nomach " _
        & " ,coalesce(e.tonage_mach,0) tonage_mach,case when (g.cavity=0 or g.ct=0) then 0 else round ((prod_plan_1/((60 / g.ct) * g.cavity * hour_p_shift * shift_usg * 60 )*faktor_productivity)/" & ar_hkw(1) * 1 & "*100,6) end presenku " _
        & " ,faktor_productivity,state_mach, mold_no,subcont,hour_p_shift,shift_usg,cavity_std,item_muloq,item_perbox " _
        & " ,priorit,submch,timeupdate,prc_prodplan from ltpp_generate a " _
        & " inner join mst_item b on a.assy_no=b.item_id " _
        & " inner join r_customer c on b.cust_id=c.cust_id " _
        & " inner join loadcap_mst_product_r d on a.assy_no=d.partno" _
        & " inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc" _
        & " inner join loadcap_proc g on d.partno=g.partno" _
        & " left join loadcap_mst_mach e on g.prod_nomach=e.no_mach" _
        & " where stscode_id='01' AND a.ltpp_doc='" & CmbDocument & "' and a.rev=" & txtRevision & " and (prod_nomach is not null and prod_nomach!='0' and prod_nomach!='')" _
        & " order by subcont asc,19 desc, priorit asc,2"
    Set rsB = Nothing
    Set rsB = New ADODB.Recordset
    rsB.Open qry, Con, adOpenStatic, adLockOptimistic
    
    Erase aPartPrior
    ReDim aPartPrior(1 To 1) As String
    resetArrayOVR
    ReDim nm_msn_full(1 To 1) As String
    ReDim nm_mold_full(1 To 1) As String
    settingGridName
    For i = 1 To 4
        If IsNumeric(ar_hkw(i)) = False Then MsgBox "Set HKW value in " & ar_hkw_bln(i) & " first", vbExclamation: Screen.MousePointer = 0: Exit Sub
    Next
    If rsB.RecordCount > 0 Then
        ProgressBar1.Visible = True
        ProgressBar1.Max = rsB.RecordCount
        anaGrid.rows = 3
        
        While Not rsB.EOF
            rsKeArray RTrim(rsB(1))
            rsB.MoveNext
        Wend
        
        rsB.Fields("assy_no").Properties("Optimize") = True
        rsB.Fields("priorit").Properties("Optimize") = True
        rsB.Fields("subcont").Properties("Optimize") = True
        rsB.Fields("mold_no").Properties("Optimize") = True
        
        StatusBar1.Panels(1).Text = rsB.RecordCount & " row(s)"
        rsW.Filter = adFilterNone
        If rsW.RecordCount = 0 Then
            For dob = 1 To UBound(aPartPrior) - 1
                rsB.Filter = adFilterNone
                rsB.Filter = "assy_no='" & aPartPrior(dob) & "'"
                rsB.Sort = "priorit ASC"
                dob2 = 1
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                    rsB.AbsolutePosition = dob2
                    temp_part = RTrim(rsB(1))
                    c_wip = rsB("p1") + rsB("p2") + rsB("p3")
                    If rsB("ct") = 0 Then
                        c_cap_p_day = 0
                    Else
                        c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
'                        If rsB("item_perbox") = 0 Then
'                            c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
'                        Else
'                            c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
'                        End If
                    End If
                    hkw1 = ar_hkw(1) * 1
                    presentMesinUse = rsB("presenku")
                    
                    If IsNull(rsB("mold_no")) Or rsB("mold_no") = "" Or rsB("mold_no") = "0" Then
                        nomold = "nomo"
                    Else
                        nomold = rsB("mold_no")
                    End If
                    
                    anaGrid.rows = anaGrid.rows + 1
                    dBariss = anaGrid.rows - 1
                    anaGrid.TextMatrix(dBariss, 0) = anaGrid.rows - 3
                    anaGrid.TextMatrix(dBariss, 1) = rsB(0)
                    anaGrid.TextMatrix(dBariss, 2) = RTrim(rsB(1))
                    anaGrid.TextMatrix(dBariss, 3) = rsB(2)
                    anaGrid.TextMatrix(dBariss, 4) = FormatNumber(rsB(3), 0)
                    anaGrid.TextMatrix(dBariss, 5) = c_wip
                    anaGrid.TextMatrix(dBariss, 6) = rsB("fc1")

                    If rsB("fc1") = 0 Then
                        anaGrid.TextMatrix(dBariss, 7) = 0
                    Else
                        anaGrid.TextMatrix(dBariss, 7) = (rsB(3) + c_wip) / rsB("fc1")
                    End If
    
                    anaGrid.TextMatrix(dBariss, 8) = IIf(IsNull(rsB("subcont")), "no", rsB("subcont")) 'kebijakan_subc
                    anaGrid.TextMatrix(dBariss, 9) = getProdplan(RTrim(rsB(1)), 1) 'rsB("prod_plan_1")
                    anaGrid.TextMatrix(dBariss, 10) = getProdplan(RTrim(rsB(1)), 2) 'rsB("prod_plan_2")
                    anaGrid.TextMatrix(dBariss, 11) = getProdplan(RTrim(rsB(1)), 3) 'rsB("prod_plan_3")
                    anaGrid.TextMatrix(dBariss, 12) = getProdplan(RTrim(rsB(1)), 4) 'rsB("prod_plan_4")
                    anaGrid.TextMatrix(dBariss, 13) = rsB("cavity")
                    anaGrid.TextMatrix(dBariss, 14) = rsB("ct")
                    anaGrid.TextMatrix(dBariss, 15) = rsB("manpower")
                    anaGrid.TextMatrix(dBariss, 16) = rsB("ct_2")
                    anaGrid.TextMatrix(dBariss, 17) = c_cap_p_day
                    anaGrid.TextMatrix(dBariss, 18) = c_cap_p_day
                    anaGrid.TextMatrix(dBariss, 19) = c_cap_p_day * hkw1
                    anaGrid.TextMatrix(dBariss, 20) = 0
                    
                    anaGrid.TextMatrix(dBariss, 29) = 0
                    anaGrid.TextMatrix(dBariss, 21) = 0
                    anaGrid.TextMatrix(dBariss, 22) = IIf(IsNull(rsB("prod_nomach")), "nom", rsB("prod_nomach")) 'rsB("tonage_mach")
                    anaGrid.TextMatrix(dBariss, 23) = IIf(IsNull(rsB("tonage_mach")), "nom", rsB("tonage_mach"))
                    anaGrid.TextMatrix(dBariss, 24) = nomold 'IIf(IsNull(rsB("mold_no")) Or IsNumeric(rsB("mold_no")) = False, "nomo", rsB("mold_no"))
                    anaGrid.TextMatrix(dBariss, 25) = 0
                    anaGrid.TextMatrix(dBariss, 26) = IIf(IsNull(rsB("state_mach")), 0, rsB("state_mach")) 'IIf(ovrd_msnutama < 0, 0, ovrd_msnutama)
                    anaGrid.TextMatrix(dBariss, 27) = 0
                    anaGrid.TextMatrix(dBariss, 30) = IIf(IsNull(rsB("faktor_productivity")), 0, rsB("faktor_productivity"))
                    anaGrid.TextMatrix(dBariss, 31) = IIf(IsNull(rsB("hour_p_shift")), 0, rsB("hour_p_shift"))
                    anaGrid.TextMatrix(dBariss, 32) = rsB("shift_usg")
                    anaGrid.TextMatrix(dBariss, 33) = 0
                    anaGrid.TextMatrix(dBariss, 34) = IIf(IsNull(rsB("cavity_std")), 0, rsB("cavity_std"))
                    anaGrid.TextMatrix(dBariss, 35) = IIf(IsNull(rsB("timeupdate")), "", Format(rsB("timeupdate"), "yyyy-MM-dd HH:mm:ss"))
                    anaGrid.TextMatrix(dBariss, 36) = rsB("prc_prodplan")
                    
                    If IsNull(rsB("state_mach")) Or nomold = "nomo" Or nomold = "0" Or anaGrid.TextMatrix(dBariss, 8) = "yes" Then
                        If rsB("submch") = True Then
                            ProsesPresent dBariss, 1, hkw1, c_cap_p_day
                        Else
                            ProsesPresent dBariss, 0, hkw1, c_cap_p_day
                        End If
                    Else
                        rsPartMCH.Filter = adFilterNone
                        rsPartMCH.Filter = "part_mch = '" & anaGrid.TextMatrix(dBariss, 22) & "'"
                        If rsPartMCH.RecordCount > 0 Then
                            rsPartMCH.Filter = adFilterNone
                            rsPartMCH.Filter = "part_mch = '" & anaGrid.TextMatrix(dBariss, 22) & "' and part_used='" & anaGrid.TextMatrix(dBariss, 2) & "'"
                            If rsPartMCH.RecordCount > 0 Then
                                ProsesPresent dBariss, rsB("state_mach"), hkw1, c_cap_p_day
                            Else
                                ProsesPresent dBariss, 0, hkw1, c_cap_p_day
                            End If
                        Else
                            ProsesPresent dBariss, rsB("state_mach"), hkw1, c_cap_p_day
                        End If
                    End If
    
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    dob2 = dob2 + 1
                Loop
            Next
        Else
            For dob = 1 To UBound(aPartPrior) - 1
                rsB.Filter = adFilterNone
                rsB.Filter = "assy_no='" & aPartPrior(dob) & "'"
                rsB.Sort = "priorit ASC"
                dob2 = 1
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                    rsB.AbsolutePosition = dob2
                    c_wip = rsB("p1") + rsB("p2") + rsB("p3")
                    If rsB("ct") = 0 Then
                        c_cap_p_day = 0
                    Else
                        c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
'                        If rsB("item_perbox") = 0 Then
'                            c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
'                        Else
'                            c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
'                        End If
                    End If
                    hkw1 = ar_hkw(1) * 1
                    presentMesinUse = rsB("presenku")
                    
                    If IsNull(rsB("mold_no")) Or rsB("mold_no") = "" Or rsB("mold_no") = "0" Then
                        nomold = "nomo"
                    Else
                        nomold = rsB("mold_no")
                    End If
                    
                    anaGrid.rows = anaGrid.rows + 1
                    dBariss = anaGrid.rows - 1
                    anaGrid.TextMatrix(dBariss, 0) = anaGrid.rows - 3
                    anaGrid.TextMatrix(dBariss, 1) = rsB(0)
                    anaGrid.TextMatrix(dBariss, 2) = RTrim(rsB(1))
                    anaGrid.TextMatrix(dBariss, 3) = rsB(2)
                    anaGrid.TextMatrix(dBariss, 4) = FormatNumber(rsB(3), 0)
                    anaGrid.TextMatrix(dBariss, 5) = c_wip
                    anaGrid.TextMatrix(dBariss, 6) = rsB("fc1")
    
                    If rsB("fc1") = 0 Then
                        anaGrid.TextMatrix(dBariss, 7) = 0
                    Else
                        anaGrid.TextMatrix(dBariss, 7) = (rsB(3) + c_wip) / rsB("fc1")
                    End If
    
                    anaGrid.TextMatrix(dBariss, 8) = IIf(IsNull(rsB("subcont")), "no", rsB("subcont")) 'kebijakan_subc
                    anaGrid.TextMatrix(dBariss, 9) = getProdplan(RTrim(rsB(1)), 1) 'rsB("prod_plan_1")
                    anaGrid.TextMatrix(dBariss, 10) = getProdplan(RTrim(rsB(1)), 2) 'rsB("prod_plan_2")
                    anaGrid.TextMatrix(dBariss, 11) = getProdplan(RTrim(rsB(1)), 3) 'rsB("prod_plan_3")
                    anaGrid.TextMatrix(dBariss, 12) = getProdplan(RTrim(rsB(1)), 4) 'rsB("prod_plan_4")
                    anaGrid.TextMatrix(dBariss, 13) = rsB("cavity")
                    anaGrid.TextMatrix(dBariss, 14) = rsB("ct")
                    anaGrid.TextMatrix(dBariss, 15) = rsB("manpower")
                    anaGrid.TextMatrix(dBariss, 16) = rsB("ct_2")
                    anaGrid.TextMatrix(dBariss, 17) = c_cap_p_day
                    anaGrid.TextMatrix(dBariss, 18) = c_cap_p_day
                    anaGrid.TextMatrix(dBariss, 19) = c_cap_p_day * hkw1
                    anaGrid.TextMatrix(dBariss, 31) = IIf(IsNull(rsB("hour_p_shift")), 0, rsB("hour_p_shift"))
                    anaGrid.TextMatrix(dBariss, 32) = rsB("shift_usg")
                    anaGrid.TextMatrix(dBariss, 20) = wo_daymamo(rsB("prod_nomach"), nomold, RTrim$(rsB(1)))
                    anaGrid.TextMatrix(dBariss, 29) = anaGrid.TextMatrix(dBariss, 20)
                    anaGrid.TextMatrix(dBariss, 21) = 0
                    anaGrid.TextMatrix(dBariss, 22) = IIf(IsNull(rsB("prod_nomach")), "nom", rsB("prod_nomach")) 'rsB("tonage_mach")
                    anaGrid.TextMatrix(dBariss, 23) = IIf(IsNull(rsB("tonage_mach")), "nom", rsB("tonage_mach"))
                    anaGrid.TextMatrix(dBariss, 24) = nomold 'IIf(IsNull(rsB("mold_no")) Or IsNumeric(rsB("mold_no")) = False, "nomo", rsB("mold_no"))
                    anaGrid.TextMatrix(dBariss, 25) = 0

                    anaGrid.TextMatrix(dBariss, 26) = IIf(IsNull(rsB("state_mach")), 0, rsB("state_mach")) 'IIf(ovrd_msnutama < 0, 0, ovrd_msnutama)
                    anaGrid.TextMatrix(dBariss, 34) = IIf(IsNull(rsB("cavity_std")), 0, rsB("cavity_std"))
                    anaGrid.TextMatrix(dBariss, 27) = 0
                    anaGrid.TextMatrix(dBariss, 30) = IIf(IsNull(rsB("faktor_productivity")), 0, rsB("faktor_productivity"))
                    anaGrid.TextMatrix(dBariss, 33) = wo_qtymamo(rsB("prod_nomach"), nomold, RTrim$(rsB(1)))
                    anaGrid.TextMatrix(dBariss, 35) = IIf(IsNull(rsB("timeupdate")), "", Format(rsB("timeupdate"), "yyyy-MM-dd HH:mm:ss"))
                    anaGrid.TextMatrix(dBariss, 36) = rsB("prc_prodplan")
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    dob2 = dob2 + 1
                Loop
            Next
            Dim stmach As Byte
            With anaGrid
                For dob = 3 To .rows - 1
                    If .TextMatrix(dob, 26) = "0" Or .TextMatrix(dob, 24) = "nomo" Or .TextMatrix(dob, 24) = "0" Or .TextMatrix(dob, 8) = "yes" Then
                        If .TextMatrix(dob, 8) = "yes" Then '
                            ProsesPresent dob, 1, hkw1, .TextMatrix(dob, 17)
                        Else
                            ProsesPresent dob, 0, hkw1, .TextMatrix(dob, 17)
                        End If
                    Else
                        If .TextMatrix(dob, 8) = "yes" Then
                            stmach = 1
                        Else
                            stmach = 1
                        End If
                        rsPartMCH.Filter = adFilterNone
                        rsPartMCH.Filter = "part_mch = '" & .TextMatrix(dob, 22) & "'"
                        
                        If rsPartMCH.RecordCount > 0 Then
'                            If .TextMatrix(dob, 2) = "6040-6111-30" And .TextMatrix(dob, 22) = "INJ-06" Then
'                                MsgBox "INI DIA"
'                            End If
                            rsPartMCH.Filter = adFilterNone
                            rsPartMCH.Filter = "part_mch = '" & .TextMatrix(dob, 22) & "' and part_used='" & .TextMatrix(dob, 2) & "'"
                            If rsPartMCH.RecordCount > 0 Then
                                'MsgBox "SINI INI DIA "
'                                If .TextMatrix(dob, 2) = "6040-6111-30" And .TextMatrix(dob, 22) = "INJ-06" Then
'                                    MsgBox "DAN INI DIA"
'                                End If
                                ProsesPresent dob, stmach, hkw1, .TextMatrix(dob, 17) ' c_cap_p_day
                            Else
'                                If .TextMatrix(dob, 2) = "6040-6111-30" And .TextMatrix(dob, 22) = "INJ-06" Then
'                                    MsgBox "ATAU INI DIA"
'                                End If
                                ProsesPresent dob, 0, hkw1, .TextMatrix(dob, 17)
                            End If
                        Else
                            ProsesPresent dob, stmach, hkw1, .TextMatrix(dob, 17)
                        End If
                    End If
                Next
            End With
        End If
        

        '&&& end ane
        
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False

    Else
        anaGrid.rows = 3
    End If
    For j = 3 To anaGrid.rows - 1
        For x = 1 To UBound(nm_msn_full)
            If nm_msn_full(x) = anaGrid.TextMatrix(j, 22) Then
                anaGrid.Col = 22
                anaGrid.Row = j
                anaGrid.CellBackColor = RGB(255, 255, 62)
            End If
        Next
    Next
    
    '**periksa yang OVERLOAD hanya 1 hari
    '*** uniq part no, mold, mesin
    '    List2.AddItem "*$$$$$$$$*"
    For dBariss = 1 To UBound(c_part)
        If ar_prodplan(dBariss) * 1 > 0 Then
            rsB.Filter = adFilterNone
            rsB.Filter = "assy_no='" & c_part(dBariss) & "' and subcont='no'"
            rsB.Sort = "mold_no ASC"
            ttMold = ""
            tToutalMold = 0
            'Periksa mold 1
            For dob2 = 1 To rsB.RecordCount
                rsB.AbsolutePosition = dob2
                If ttMold <> rsB("mold_no") Then
                    tToutalMold = tToutalMold + 1
                    ttMold = rsB("mold_no")
                End If
            Next
            rsB.Sort = ""
            dob2 = 1
            If (tToutalMold) <= 1 Then '<=
                dob2c = 0
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                    rsB.AbsolutePosition = dob2
                    If IsNull(rsB("prod_nomach")) = False Then
                        c_NDMtZ = NDMtZ(rsB("prod_nomach"), rsB("assy_no"))
                        If rsB("state_mach") = 1 And c_NDMtZ > 0 Then
                            dob2c = dob2
                        End If
                    End If
                    dob2 = dob2 + 1
                Loop
                If dob2c > 0 Then
                    rsB.AbsolutePosition = dob2c
                    If rsB("ct") = 0 Then
                        c_cap_p_day = 0
                    Else
                        c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
                        If rsB("item_perbox") = 0 Then
                            c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
                        Else
                            c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
                        End If
                    End If
                    need_day = FormatNumber(ar_prodplan(dBariss) / c_cap_p_day, 2)
                    If need_day * 1 <= 1 And need_day * 1 > 0 Then
                        If PerMcNow(rsB("prod_nomach"), hkw1) + (need_day / hkw1 * 100) <= 105 Then
                            PlotSISA c_part(dBariss), rsB("mold_no"), need_day, rsB("prod_nomach")
                        End If
                    End If
                Else
                    dob2 = 1
                    Do
                        If dob2 > rsB.RecordCount Then Exit Do
                        rsB.AbsolutePosition = dob2
                        If rsB("state_mach") = 1 Then
                            If rsB("ct") = 0 Then
                                c_cap_p_day = 0
                            Else
                                c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
                                If rsB("item_perbox") = 0 Then
                                    c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
                                Else
                                    c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
                                End If
                            End If
                            need_day = FormatNumber(ar_prodplan(dBariss) / c_cap_p_day, 2)

                            If need_day * 1 <= 1 And need_day * 1 > 0 Then
                                If PerMcNow(rsB("prod_nomach"), hkw1) + (need_day / hkw1 * 100) <= 105 Then
                                    If blockSpec(rsB("prod_nomach"), c_part(dBariss)) Then
                                        PlotSISA c_part(dBariss), rsB("mold_no"), need_day, rsB("prod_nomach")
                                    End If
                                End If
                            End If
                        End If
                        dob2 = dob2 + 1
                    Loop
                End If
            Else
                Do
                    If dob2 > rsB.RecordCount Then Exit Do
                    rsB.AbsolutePosition = dob2
                    If rsB("state_mach") = 1 Then
                        If rsB("ct") = 0 Then
                            c_cap_p_day = 0
                        Else
                            c_cap_p_day = ((60 / rsB("ct")) * rsB("cavity") * rsB("hour_p_shift") * rsB("shift_usg") * 60) * rsB("faktor_productivity")
                            If rsB("item_perbox") = 0 Then
                                c_cap_p_day = isi(rsB("item_muloq"), c_cap_p_day, "b")
                            Else
                                c_cap_p_day = isi(rsB("item_perbox"), c_cap_p_day, "b")
                            End If
                        End If
                        need_day = FormatNumber(ar_prodplan(dBariss) / c_cap_p_day, 2)
                        If need_day * 1 <= 1 And need_day * 1 > 0 Then
                            If PerMcNow(rsB("prod_nomach"), hkw1) + (need_day / hkw1 * 100) <= 105 Then
                                If blockSpec(rsB("prod_nomach"), c_part(dBariss)) Then
                                    PlotSISA c_part(dBariss), rsB("mold_no"), need_day, rsB("prod_nomach")
                                End If
                            End If
                        End If
                    End If
                    dob2 = dob2 + 1
                Loop
            End If
        End If
    Next
    
    '-----------------rekap need day
    For x = 3 To anaGrid.rows - 1
        For j = 3 To anaGrid.rows - 1
            If anaGrid.TextMatrix(x, 2) = anaGrid.TextMatrix(j, 2) Then
                If Val(anaGrid.TextMatrix(x, 20)) > 0 Then
                    anaGrid.TextMatrix(x, 21) = Val(anaGrid.TextMatrix(x, 21)) + Val(anaGrid.TextMatrix(j, 20))
                End If
            End If
        Next
    Next
    
    '--------sum perbulan load cap machine
    For x = 3 To anaGrid.rows - 1
        anaGrid.TextMatrix(x, 27) = Val(anaGrid.TextMatrix(x, 20)) / hkw1 * 100
        anaGrid.TextMatrix(x, 28) = (Val(anaGrid.TextMatrix(x, 20)) / hkw1) * Val(anaGrid.TextMatrix(x, 15))
    Next
    gridFormatNum
    If CmbDocument <> "" Then checkNeedMoldMachine
    anaGrid.Refresh
    Screen.MousePointer = 0
End Sub

Private Function blockSpec(pmesin As String, pPART As String) As Boolean
    rsPartMCH.Filter = adFilterNone
    rsPartMCH.Filter = "part_mch = '" & pmesin & "'"
    
    If rsPartMCH.RecordCount > 0 Then
        rsPartMCH.Filter = adFilterNone
        rsPartMCH.Filter = "part_mch = '" & pmesin & "' and part_used='" & pPART & "'"
        If rsPartMCH.RecordCount > 0 Then
            blockSpec = True
        Else
            blockSpec = False
        End If
    Else
        blockSpec = True
    End If
End Function

Private Function getProdplan(part As String, ppKe As Integer)
    Dim X1 As Long
    X1 = 1
    Do
        If X1 > UBound(ar_propl) Then Exit Do
        If c_part(X1) = part Then
            If ppKe = 1 Then
                getProdplan = ar_propl(X1)
            ElseIf ppKe = 2 Then
                getProdplan = ar_propl2(X1)
            ElseIf ppKe = 3 Then
                getProdplan = ar_propl3(X1)
            ElseIf ppKe = 4 Then
                getProdplan = ar_propl4(X1)
            End If
            Exit Do
        End If
        X1 = X1 + 1
    Loop
End Function

Private Function NDMtZ(pmesin As String, pPART As String) As Variant 'CARI Mesin need day >0
    Dim k As Long
    With anaGrid
        For k = 3 To .rows - 1
            If .TextMatrix(k, 22) = pmesin And .TextMatrix(k, 2) = RTrim(pPART) And .TextMatrix(k, 20) * 1 > 0 Then
                NDMtZ = .TextMatrix(k, 20)
            End If
        Next
    End With
End Function

Private Sub resetArrayOVR()
    Erase ar_Sisa
    ReDim ar_Sisa(1 To 1) As Variant
    Erase ar_PartSisa
    ReDim ar_PartSisa(1 To 1) As String
    Erase ar_MoldSisa
    ReDim ar_MoldSisa(1 To 1) As String
    Erase ar_MesinSisa
    ReDim ar_MesinSisa(1 To 1) As String
End Sub

Private Sub addFmesin(pmesin As String)
    Dim ade As Boolean, dib As Long
    For dib = 1 To UBound(ar_fMesin)
        If ar_fMesin(dib) = pmesin Then
            ade = True
            Exit For
        Else
            ade = False
        End If
    Next
    If ade = False Then
        ar_fMesin(UBound(ar_fMesin)) = pmesin
        ReDim Preserve ar_fMesin(1 To UBound(ar_fMesin) + 1) As String
    End If
End Sub

Private Function MchovrExsist(pmesin As String) As Boolean
    Dim m As Long
    For m = 1 To UBound(ar_fMesin)
        If pmesin = ar_fMesin(m) Then
            MchovrExsist = True
            Exit For
        Else
            MchovrExsist = False
        End If
    Next
End Function

Private Sub setProdPlan0(pPART As String)
    Dim b As Long
    b = 1
    Do
        If b > UBound(c_part) Then Exit Do
        
        If c_part(b) = pPART Then
            ar_prodplan(b) = 0
            Exit Do
        End If
       b = b + 1
    Loop
End Sub

Private Sub txtRevision_DropDown()
    qry = "select distinct on (rev) rev from ltpp_generate where period='" & Format(DTPicker1.Value, "yyyyMM") & "' and ltpp_doc='" & CmbDocument & "'"
    Set RsA = Con.Execute(qry)
    txtRevision.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            txtRevision.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub


