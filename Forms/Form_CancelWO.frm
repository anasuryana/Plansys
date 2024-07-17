VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_CancelWO 
   Caption         =   "Cancel Work Order"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8865
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
   ScaleHeight     =   6075
   ScaleWidth      =   8865
   Begin ACTIVESKINLibCtl.SkinLabel lblplandate 
      Height          =   375
      Left            =   4320
      OleObjectBlob   =   "Form_CancelWO.frx":0000
      TabIndex        =   22
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   266
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   600
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Detail"
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
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   8625
      TabIndex        =   7
      Top             =   720
      Width           =   8655
      Begin VB.Label lblplanqty 
         BackColor       =   &H00FFFFC0&
         Caption         =   "..."
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblmoldno 
         BackColor       =   &H00FFFFC0&
         Caption         =   "..."
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblmachine 
         BackColor       =   &H00FFFFC0&
         Caption         =   "..."
         Height          =   255
         Left            =   6240
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         Caption         =   " Plan Qty"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         Caption         =   " Mold No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         Caption         =   " Machine"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbllotno 
         BackColor       =   &H00FFFFC0&
         Caption         =   "..."
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblitemname 
         BackColor       =   &H00FFFFC0&
         Caption         =   "..."
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblitemid 
         BackColor       =   &H00FFFFC0&
         Caption         =   "..."
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         Caption         =   "  Lot No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "  Item Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "  Item Id"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      Height          =   270
      Left            =   720
      TabIndex        =   6
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "Commit"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      ToolTipText     =   "Commit tha you are want to cancel the DO"
      Top             =   5640
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5318
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdFindWo 
      Caption         =   "..."
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtWoid 
      BackColor       =   &H00FFFF00&
      Height          =   390
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_CancelWO.frx":006E
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8400
      OleObjectBlob   =   "Form_CancelWO.frx":00D0
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblrevisiMPP 
      Height          =   255
      Left            =   6120
      OleObjectBlob   =   "Form_CancelWO.frx":0304
      TabIndex        =   23
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "Form_CancelWO"
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
Dim i As Byte
Const centang As String = "þ"
Const ga_centang As String = "ý"

Private Sub cmdAdd_Click()
    Dim u As Byte
    With grid1
        If .rows > 1 Then
        For i = 1 To .rows - 1
            If .TextMatrix(i, 1) = txtWoid Then
                MsgBox "Duplicate", vbExclamation
                Exit Sub
            End If
        Next
        End If
        .rows = .rows + 1
        .TextMatrix(.rows - 1, 0) = .rows - 1
        .TextMatrix(.rows - 1, 1) = txtWoid
        .TextMatrix(.rows - 1, 2) = lbllotno
        .TextMatrix(.rows - 1, 3) = lblmachine
        .TextMatrix(.rows - 1, 4) = lblmoldno
        .TextMatrix(.rows - 1, 5) = lblitemid
        .TextMatrix(.rows - 1, 6) = lblitemname
        .TextMatrix(.rows - 1, 7) = Replace(lblplandate, "Plan date : ", "")
        .TextMatrix(.rows - 1, 8) = lblplanqty
        .TextMatrix(.rows - 1, 9) = Replace(lblrevisiMPP, "MPP Revision : ", "")
        .Col = .Cols - 1
        .TextMatrix(.rows - 1, .Cols - 1) = centang
        .Row = .rows - 1
        .CellFontName = "Wingdings"
    End With
End Sub

Private Sub gridtoform()
    With grid1
        If .rows > 1 Then
            lblitemid = .TextMatrix(.RowSel, 5)
            lblitemname = .TextMatrix(.RowSel, 6)
            txtWoid = .TextMatrix(.RowSel, 1)
            lbllotno = .TextMatrix(.RowSel, 2)
            lblmachine = .TextMatrix(.RowSel, 3)
            lblmoldno = .TextMatrix(.RowSel, 4)
            lblplandate = "Plan date : " & .TextMatrix(.RowSel, 7)
            lblplanqty = .TextMatrix(.RowSel, 8)
        End If
    End With
End Sub

Private Sub cmdCommit_Click()
    Dim qry As String
    Dim r As Byte
    Dim cplandate As String
    Dim cmold As String
    Dim cpart As String
    Dim cmach As String
    If MsgBox("Are you sure want to conitnue ?", vbQuestion + vbYesNo) = vbYes Then
        With grid1
            For r = 1 To .rows - 1
                cplandate = .TextMatrix(r, 7)
                cpart = .TextMatrix(r, 5)
                cmold = .TextMatrix(r, 4)
                cmach = .TextMatrix(r, 3)
                qry = "update worko set lotno=NULL, qty=0 " _
                & " WHERE wo_no='" & .TextMatrix(r, 1) & "'"
                Con.Execute qry
                
                qry = "UPDATE mpp_gen a SET planqty=0 where " _
                & " plandate='" & cplandate & "' and lcd_itemdid='" & cpart & "' and " _
                & " reg_mold='" & cmold & "' and no_mach='" & cmach & "' and mpp_revisi =(select max(mpp_revisi) from mpp_gen " _
                & " where plandate=a.plandate and lcd_itemdid=a.lcd_itemdid " _
                & " and reg_mold=a.reg_mold and no_mach=a.no_mach)"
                Con.Execute qry
            Next
            .rows = 1
            MsgBox "Processed successfully"
        End With
    End If
End Sub

Private Sub cmdFindWo_Click()
    GetForm = Me.Name
    popup_wono.Show 1
    cmdAdd.SetFocus
End Sub


Private Sub cmdMinus_Click()
    Dim k As Byte
    Dim R1 As Byte
    Dim r As Byte
    With grid1
        If .rows > 1 Then
            For k = 0 To .Cols - 1
                .TextMatrix(.Row, k) = ""
            Next
            If .rows = 2 Then
                .rows = 1
            Else
                R1 = .RowSel + 1
                For r = R1 To .rows - 1
                    For k = 1 To .Cols - 1
                        .TextMatrix(r - 1, k) = .TextMatrix(r, k)
                    Next
                Next
                .rows = .rows - 1
                For r = 1 To .rows - 1
                    .TextMatrix(r, 0) = r
                Next
            End If
        End If
    End With
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

Private Sub settingFG()
    
    With grid1
        .Cols = 11
        .rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .ColAlignment(2) = flexAlignLeftCenter
        
        i = 0
        .TextMatrix(0, i) = "No."
        .ColWidth(i) = 700
        .ColAlignment(i) = flexAlignLeftCenter
        
        
        i = 1
        .TextMatrix(0, i) = "WO No."
        .ColWidth(i) = 1700
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 2
        .TextMatrix(0, i) = "Lot No."
        .ColWidth(i) = 1000
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 3
        .TextMatrix(0, i) = "Machine"
        .ColWidth(i) = 1100
        
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
        .TextMatrix(0, i) = "Issue Date"
        .ColWidth(i) = 1900
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 8
        .TextMatrix(0, i) = "Qty"
        
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 9
        .TextMatrix(0, i) = "MPP Revision"
        .ColWidth(i) = 1800
        .ColAlignment(i) = flexAlignCenterCenter
        
        i = 10
        .TextMatrix(0, i) = "..."
        .ColWidth(i) = 500
        .ColAlignment(i) = flexAlignCenterCenter
        
    End With
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

Private Sub Form_Load()
    AddTab Me
    Width = 9105
    Height = 6645
    activeTheme Skin1, Me
    settingFG
    
    
    grid1.rows = 1
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub grid1_Click()
    With grid1
        If .Col = .Cols - 1 Then
            If .Text = centang Then
                .Text = ga_centang
            Else
                .Text = centang
            End If
        End If
    End With
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
    With grid1
        If .Col = .Cols - 1 Then
            If KeyAscii = vbKeySpace Then
                If .TextMatrix(.RowSel, .ColSel) = centang Then
                    .TextMatrix(.RowSel, .ColSel) = ga_centang
                Else
                    .TextMatrix(.RowSel, .ColSel) = centang
                End If
            End If
        End If
    End With
End Sub

Private Sub grid1_RowColChange()
    gridtoform
End Sub
