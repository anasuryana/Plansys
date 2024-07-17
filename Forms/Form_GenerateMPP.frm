VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_GenerateMPP 
   Caption         =   "Generate MPP"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_GenerateMPP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   19815
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrameLog 
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6075
      ScaleWidth      =   19395
      TabIndex        =   61
      Top             =   3360
      Visible         =   0   'False
      Width           =   19455
      Begin VB.CommandButton closeLog 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   255
         Left            =   19080
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtLog 
         Height          =   5775
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Top             =   240
         Width           =   19335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG:"
         Height          =   255
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox frameGenerate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   19395
      TabIndex        =   0
      Top             =   240
      Width           =   19455
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14400
         TabIndex        =   60
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "EXCEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13080
         TabIndex        =   59
         Top             =   2280
         Width           =   1215
      End
      Begin VB.PictureBox FrameStatusMPP 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         ScaleHeight     =   555
         ScaleWidth      =   8835
         TabIndex        =   44
         Top             =   2280
         Width           =   8895
         Begin VB.CommandButton cmdLog 
            Caption         =   "LOG"
            Height          =   255
            Left            =   2520
            TabIndex        =   64
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label stSave 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   840
            TabIndex        =   57
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "LOG:"
            Height          =   255
            Left            =   2520
            TabIndex        =   56
            Top             =   0
            Width           =   855
         End
         Begin VB.Label logWOMf 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   7920
            TabIndex        =   55
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "Material"
            Height          =   255
            Left            =   7080
            TabIndex        =   54
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label logWOMs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   7080
            TabIndex        =   53
            Top             =   240
            Width           =   855
         End
         Begin VB.Label logWOCf 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   6120
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "WO Card"
            Height          =   255
            Left            =   5280
            TabIndex        =   51
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label logWOCs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   5280
            TabIndex        =   50
            Top             =   240
            Width           =   855
         End
         Begin VB.Label logWIPf 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   4320
            TabIndex        =   49
            Top             =   240
            Width           =   855
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000000&
            BorderWidth     =   2
            X1              =   2400
            X2              =   2400
            Y1              =   0
            Y2              =   600
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "WIP"
            Height          =   255
            Left            =   3480
            TabIndex        =   48
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label logWIPs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3480
            TabIndex        =   47
            Top             =   240
            Width           =   855
         End
         Begin VB.Label stGenerate 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   840
            TabIndex        =   46
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "STATUS:"
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   975
         End
      End
      Begin MSComctlLib.ProgressBar progBar 
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2040
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11280
         TabIndex        =   43
         Top             =   2280
         Width           =   1575
      End
      Begin VB.PictureBox FrameKET 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   15840
         ScaleHeight     =   3075
         ScaleWidth      =   3555
         TabIndex        =   18
         Top             =   -120
         Width           =   3615
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            X1              =   0
            X2              =   3480
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   0
            X2              =   3480
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Total WP"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   42
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label sTotalWP 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2280
            TabIndex        =   41
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label sCdLine 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   40
            Top             =   1080
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2040
            TabIndex        =   39
            Top             =   600
            Width           =   255
         End
         Begin VB.Label sWPTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2280
            TabIndex        =   38
            Top             =   600
            Width           =   975
         End
         Begin VB.Label sLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   37
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Line"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label sRev 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   35
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Rev."
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   840
            Width           =   975
         End
         Begin VB.Label sWPFr 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   33
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " WP Date"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.Label sPeriod 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   31
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Period"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
         Begin VB.Label sDocno 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   29
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Doc No"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Width           =   975
         End
         Begin VB.Label sCapJoint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2280
            TabIndex        =   27
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label sCapCrimp 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2280
            TabIndex        =   26
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label sCapCut 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2280
            TabIndex        =   25
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label sLabelJoint 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Capacity Jointing /Day"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label sLabelCrimp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Capacity Crimping /Day"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label sLabelCut 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Capacity Cutting /Day"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label sLoadData 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2280
            TabIndex        =   21
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblStatusGenerate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   2760
            Width           =   3495
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   " Load Assy"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   1440
            Width           =   2175
         End
      End
      Begin VB.ComboBox cmbLine 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox txtRev 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   2040
         TabIndex        =   15
         Top             =   1560
         Width           =   6015
      End
      Begin MSComCtl2.DTPicker dtWPStart 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   143851521
         CurrentDate     =   42104
      End
      Begin VB.TextBox txtDocNo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   2040
         TabIndex        =   5
         Top             =   120
         Width           =   6015
      End
      Begin VB.ComboBox cmbMM 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtNote 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   1365
         Left            =   9600
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   120
         Width           =   6015
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "GENERATE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   2
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtYear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   5880
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtWPEnd 
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   144048129
         CurrentDate     =   42104
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  Line"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8400
         TabIndex        =   17
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  WP Start Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  MPP Period"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  Document No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  Last Revision"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  Note"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8400
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGridMPP 
      Height          =   6135
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
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
      Left            =   120
      OleObjectBlob   =   "Form_GenerateMPP.frx":000C
      Top             =   240
   End
   Begin VB.Timer queReGenerate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   360
   End
   Begin VB.Timer queryTime 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   360
   End
   Begin MSComDlg.CommonDialog comSave 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "Form_GenerateMPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Const xlContinuous = 1
Private Const xlLeft = -4131

Dim i As Integer
Dim iLoop As Integer
Dim iROW As Integer
Dim iCol As Integer
Dim rowHeader As Integer
Dim rowPlan As Long
Dim arrLine() As Variant
Dim setPeriod As String
Dim queOffset As Integer
Dim totalWP As Integer
Dim rowPosFooter As Integer
Dim stEditor As Boolean
Dim posAssy As Integer
Dim lastPos As Integer
Dim repDate As Date
Dim repNotes As String

Private Sub getListMonth()
    Dim iMonth As Integer
    cmbMM.Clear
    For iMonth = 1 To 12
        cmbMM.AddItem Format(DateSerial(Year(Now), iMonth, 1), "MMMM")
    Next
    cmbMM.ListIndex = Month(Now) - 1
End Sub

Private Sub getListLine()
    Set RsGet = Con.Execute("select * from wip_mst_line order by nm_line")
    If RsGet.RecordCount > 0 Then
        ReDim arrLine(0 To RsGet.RecordCount - 1, 0 To 4)
        i = 0
        cmbLine.Clear
        Do Until RsGet.EOF
            arrLine(i, 0) = RsGet!cd_line
            arrLine(i, 1) = RsGet!nm_line
            arrLine(i, 2) = RsGet!mpp_line
            arrLine(i, 3) = RsGet!form_line
            arrLine(i, 4) = RsGet!nick_line
            cmbLine.AddItem RsGet!nm_line
            i = i + 1
            RsGet.MoveNext
        Loop
        cmbLine.ListIndex = 0
    End If
    RsGet.Close
End Sub

Private Sub closeLog_Click()
    FrameLog.Visible = False
End Sub

Private Sub cmdExcel_Click()
On Error GoTo errExcel
    FrameLog.Visible = False
    If Val(sLoadData) <= 0 Then
        MsgBox "No Assy Found!", vbExclamation, "Warning..."
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Dim oExcel As Object 'Excel.Application 'Object
    Dim oBook As Object 'Excel.Workbook 'Object
    Dim oSheet As Object 'Excel.Worksheet 'Object
    Dim getSheet As Object
    Dim rowLine As Long
    Dim lastRow As Long
    
    With comSave
        .DefaultExt = ".xls"
        .Filter = "Excel Workbook (*.xls)|*.xls"
        .ShowSave
    End With
    
    rowPlan = 0
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open pTemplateMPP
    
    Set oBook = oExcel.Workbooks(1)
    Set oSheet = oBook.Worksheets(1)
    
    oSheet.Range("B5") = "PERIODE : " & sPeriod
    oSheet.Range("F7") = ": " & txtDocNo
    oSheet.Range("F8") = ": " & Format(repDate, "DD MMMM YYYY")
    oSheet.Range("I9") = sRev
    oSheet.Range("J9") = repNotes
    oSheet.Range("A15") = sLine
    
    With MSFlexGridMPP
    'HEADER
    For iLoop = 10 To .Cols - 1
        oSheet.Cells(16, iLoop - 1) = .TextMatrix(0, iLoop)
        oSheet.Cells(17, iLoop - 1) = .TextMatrix(1, iLoop)
        oSheet.Cells(18, iLoop - 1) = .TextMatrix(2, iLoop)
        oSheet.Cells(19, iLoop - 1) = .TextMatrix(3, iLoop)
        oSheet.Cells(20, iLoop - 1) = .TextMatrix(4, iLoop)
        oSheet.Cells(21, iLoop - 1) = .TextMatrix(5, iLoop)
        oSheet.Cells(22, iLoop - 1) = .TextMatrix(6, iLoop)
        oSheet.Cells(23, iLoop - 1) = getColor(Val(.TextMatrix(7, iLoop)))
        oSheet.Cells(24, iLoop - 1) = .TextMatrix(8, iLoop)
        If Trim(.TextMatrix(0, iLoop)) = "-" Then
            oSheet.Cells(16, iLoop - 1).Resize(9, 1).Interior.Color = RGB(255, 222, 222)
        End If
    Next
    
    'DETAIL
    progBar.Max = sLoadData
    progBar.value = 0
    For i = 0 To Val(sLoadData) - 1
        rowLine = 25 + (i * 3)
        rowPlan = rowHeader + (i * 7) + 1
        oSheet.Cells(rowLine, 1) = i + 1
        oSheet.Cells(rowLine, 2) = .TextMatrix(rowPlan, 1)
        oSheet.Cells(rowLine, 3) = .TextMatrix(rowPlan, 2)
        oSheet.Cells(rowLine, 4) = .TextMatrix(rowPlan, 3)
        oSheet.Cells(rowLine, 5) = .TextMatrix(rowPlan, 4)
        oSheet.Cells(rowLine, 6) = .TextMatrix(rowPlan, 5)
        oSheet.Cells(rowLine, 7) = .TextMatrix(rowPlan, 6)
        oSheet.Cells(rowLine + 1, 7) = .TextMatrix(rowPlan + 1, 6)
        oSheet.Cells(rowLine + 2, 7) = .TextMatrix(rowPlan + 2, 6)
        oSheet.Cells(rowLine, 8) = .TextMatrix(rowPlan, 7)
        oSheet.Cells(rowLine + 1, 8) = .TextMatrix(rowPlan + 1, 7)
        oSheet.Cells(rowLine + 2, 8) = .TextMatrix(rowPlan + 2, 7)
        oSheet.Range("A" & rowLine & ":A" & rowLine + 2).Merge
        oSheet.Range("B" & rowLine & ":B" & rowLine + 2).Merge
        oSheet.Range("C" & rowLine & ":C" & rowLine + 2).Merge
        oSheet.Range("D" & rowLine & ":D" & rowLine + 2).Merge
        oSheet.Range("E" & rowLine & ":E" & rowLine + 2).Merge
        oSheet.Range("F" & rowLine & ":F" & rowLine + 2).Merge
        
        For iLoop = 10 To .Cols - 1
            oSheet.Cells(rowLine, iLoop - 1) = .TextMatrix(rowPlan, iLoop)
            oSheet.Cells(rowLine + 1, iLoop - 1) = .TextMatrix(rowPlan + 1, iLoop)
            oSheet.Cells(rowLine + 2, iLoop - 1) = .TextMatrix(rowPlan + 2, iLoop)
        Next
    
        progBar.value = i + 1
    Next
    
    'FOOTER
    For i = 0 To 7
        oSheet.Cells(rowLine + 4 + i, 1) = .TextMatrix(rowPosFooter + i, 1) & " " & .TextMatrix(rowPosFooter + i, 2)
        oSheet.Cells(rowLine + 4 + i, 7) = .TextMatrix(rowPosFooter + i, 9)
        oSheet.Range("A" & rowLine + 4 + i & ":F" & rowLine + 4 + i).Merge
        oSheet.Range("A" & rowLine + 4 + i & ":F" & rowLine + 4 + i).HorizontalAlignment = xlLeft
        oSheet.Range("G" & rowLine + 4 + i & ":H" & rowLine + 4 + i).Merge
        For iLoop = 10 To .Cols - 1
            oSheet.Cells(rowLine + 4 + i, iLoop - 1) = .TextMatrix(rowPosFooter + i, iLoop)
        Next
    Next
    i = i + 1
    oSheet.Cells(rowLine + 4 + i, 1) = UCase(sLabelCut)
    oSheet.Cells(rowLine + 5 + i, 1) = UCase(sLabelCrimp)
    oSheet.Cells(rowLine + 6 + i, 1) = UCase(sLabelJoint)
    oSheet.Cells(rowLine + 4 + i, 7) = Val(sCapCut)
    oSheet.Cells(rowLine + 5 + i, 7) = Val(sCapCrimp)
    oSheet.Cells(rowLine + 6 + i, 7) = Val(sCapJoint)
    oSheet.Range("A" & rowLine + 4 + i & ":F" & rowLine + 4 + i).Merge
    oSheet.Range("A" & rowLine + 5 + i & ":F" & rowLine + 5 + i).Merge
    oSheet.Range("A" & rowLine + 6 + i & ":F" & rowLine + 6 + i).Merge
    oSheet.Range("G" & rowLine + 4 + i & ":H" & rowLine + 4 + i).Merge
    oSheet.Range("G" & rowLine + 5 + i & ":H" & rowLine + 5 + i).Merge
    oSheet.Range("G" & rowLine + 6 + i & ":H" & rowLine + 6 + i).Merge
    oSheet.Range("A" & rowLine + 4 + i & ":F" & rowLine + 4 + i).HorizontalAlignment = xlLeft
    oSheet.Range("A" & rowLine + 5 + i & ":F" & rowLine + 5 + i).HorizontalAlignment = xlLeft
    oSheet.Range("A" & rowLine + 6 + i & ":F" & rowLine + 6 + i).HorizontalAlignment = xlLeft
    oSheet.Range("A16").Resize(14 + (Val(sLoadData) * 3), .Cols - 2).Borders.LineStyle = xlContinuous
    End With
    '--
'    oSheet.Range("B18").Resize(totalRow, 1).NumberFormat = "@"
'    oSheet.Range("A18").Resize(totalRow, 51).value = DataArray
    
    oBook.SaveAs comSave.FileName
    oExcel.Quit
    Set oExcel = Nothing
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Excel has been saved...", vbInformation, "Exported..."
Exit Sub
errExcel:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Export: " & Err.Number
    End If
End Sub

Private Function getColor(Optional colorID As Integer = 0) As String
Select Case colorID
    Case 1
        getColor = "KUNING"
    Case 2
        getColor = "BIRU MUDA"
    Case 3
        getColor = "HIJAU"
    Case 4
        getColor = "KUNING"
    Case 5
        getColor = "MERAH"
    Case 6
        getColor = "PUTIH"
    Case 7
        getColor = ""
    Case Else
        getColor = "-"
End Select
End Function

Private Sub cmdGenerate_Click()
On Error GoTo errGenerate
    cmdGenerate.Enabled = False
    cmdExcel.Enabled = False
    cmdprint.Enabled = False
    clearSummary
    setPeriod = txtYear & Format(cmbMM.ListIndex + 1, "00")
    sPeriod = setPeriod
    sWPFr = Format(dtWPStart, "YYYY-MM-DD")
    sWPTo = Format(dtWPEnd, "YYYY-MM-DD")
    sLine = arrLine(cmbLine.ListIndex, 1)
    sCdLine = arrLine(cmbLine.ListIndex, 0)
    stEditor = False
    stGenerate = "WAIT..."
    cmdSave.Enabled = False
    txtRev.BackColor = vbWhite
    Set RsDB = Con.Execute("select * from mpp_generate where doc_mpp = '" & setPeriod & sCdLine & "'")
    clearLog
    If Not RsDB.EOF Then
        txtDocNo = RsDB!doc_mpp
        txtRev = RsDB!rev
        sRev = RsDB!rev
        dtWPStart.value = RsDB!wp_dt_fr
        dtWPEnd.value = RsDB!wp_dt_to
        repDate = RsDB!time_update
        repNotes = RsDB!mpp_notes
        txtNote = repNotes
        stSave = "UPDATE"
    Else
        stSave = "NEW"
        repNotes = "<< GENERATING... >>"
        repDate = Now
    End If
    LoadDataMPP arrLine(cmbLine.ListIndex, 0), setPeriod, dtWPStart, dtWPEnd
    RsDB.Close
Exit Sub
errGenerate:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Generate: " & Err.Number
    End If
End Sub

Private Sub clearLog()
    logWIPs = 0
    logWIPf = 0
    logWOCs = 0
    logWOCf = 0
    logWOMs = 0
    logWOMf = 0
End Sub

Private Sub cmdLog_Click()
    FrameLog.Visible = True
End Sub

Private Sub cmdprint_Click()
    FrameLog.Visible = False
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSave
    Dim debugStr As String
    Dim getAssyNo As String
    Dim qtyMppWP As Double
    Dim qtyTotMpp As Double
    Dim qtyProdplan As Double
    Dim woID As String
    Dim wpID As String
    Dim wpMPP As String
    Dim wpDate As String
    Dim wpColor As Integer
    Dim wpHoldWo As String
    Dim wpHoldQty As Double
    Dim wpHoldSeri As String
    Dim ketMppProdplan As String
    Dim affMpp As Long
    Dim affWoc As Long
    Dim affWom As Long
    Dim stLog As String
    Dim stExec As String
    Dim stWoEx As String
    Dim stSkip As Boolean
    
    Dim seriHead As String
    Dim seriUrut As String
    Dim SerialNo As String
    
    rowPlan = 0
    clearLog
    cmdSave.Enabled = False
    
    FrameLog.Visible = False
    
    If Val(sLoadData) <= 0 Then
        MsgBox "No Assy Found!", vbExclamation, "Warning..."
        Exit Sub
    End If
    
    Select Case stSave
    Case "NEW"
        Con.Execute "insert into mpp_generate (doc_mpp, rev, period, line, wp_dt_fr, wp_dt_to, total_wp, user_update, time_update, mpp_notes) " _
            & "values ('" & sPeriod & sCdLine & "', " & Val(sRev) & ", '" & sPeriod & "', '" & sCdLine & "', '" & Format(dtWPStart, "YYYY-MM-DD") & "', '" & Format(dtWPEnd, "YYYY-MM-DD") & "', " & Val(sTotalWP) & ", '" & pUserName & "', now(), '" & Trim(txtNote) & "')"
        lastPos = 0
    Case "UPDATE"
        'syntax for re-generate (update)
        Con.Execute "update mpp_generate set rev = " & Val(sRev) & ", total_wp = " & Val(sTotalWP) & ", user_update = '" & pUserName & "', time_update = now(), mpp_notes = '" & txtNote & "' where doc_mpp = '" & sPeriod & sCdLine & "'"
        Con.Execute "update wip_trx_mpp set rev = " & sRev & " where periode = '" & sPeriod & "' and line = '" & sCdLine & "'"
        Set RsGet = Con.Execute("select coalesce(max(pos_assy), '00') pos_assy from wip_trx_mpp where periode = '" & sPeriod & "' and line = '" & sCdLine & "'")
        lastPos = Val(RsGet!pos_assy)
        RsGet.Close
    End Select
        With MSFlexGridMPP
        Set RsGet = Con.Execute("select coalesce(max(serial_mpp), '0') serial, lpad(extract(day from now())::varchar(2), 2, '0') daycode from wip_trx_mpp where periode = '" & sPeriod & "' and line = '" & sCdLine & "' " _
            & "and substr(serial_mpp, 1, 1) = '1' and substr(serial_mpp, 9, 1) = '2' and substr(serial_mpp, 7, 2) = lpad(extract(day from now())::varchar(2), 2, '0')")
        If Trim(RsGet!serial) <> "0" Then
            seriHead = Left(RTrim(RsGet!serial), 9)
            seriUrut = Right(RTrim(RsGet!serial), 5)
        Else
            seriHead = "1" & Right(sPeriod, 4) & sCdLine & RsGet!daycode & "2"
            seriUrut = "00000"
        End If
        RsGet.Close
        progBar.Max = sLoadData
        progBar.value = 0
        posAssy = 0
        For i = 0 To Val(sLoadData) - 1
            Sleep 100
            rowPlan = rowHeader + (i * 7) + 1
            qtyProdplan = Val(.TextMatrix(rowPlan, 7))
            qtyTotMpp = Val(.TextMatrix(rowPlan, 9))
            getAssyNo = .TextMatrix(rowPlan, 1)
            'schedule
            .Row = rowPlan - 1 'sch
            .Col = 7
            If .CellBackColor = RGB(220, 220, 75) Then
                Con.Execute "update plansys_schedule set total_qty = " & Val(.Text) & ", input_user = '" & pUserName & "', input_time = now() where period = '" & sPeriod & "' and assy_no = '" & getAssyNo & "'"
                For iLoop = 10 To .Cols - 1
                    .Col = iLoop
                    If .CellBackColor = RGB(220, 220, 75) Then
                        wpDate = .TextMatrix(9, iLoop)
                        Con.Execute "update plansys_schedule_detail set qty = " & Val(.Text) & " where period = '" & sPeriod & "' and assy_no = '" & getAssyNo & "' and date_schedule = '" & wpDate & "'"
                    End If
                Next
            End If
            'plan
            If qtyProdplan >= qtyTotMpp Then
                debugStr = debugStr & "<< NO  " & i + 1 & " >> [PRODPLAN >= QTY MPP] : TRUE   --- " _
                    & " [ASSY NO] " & getAssyNo & " [PRODPLAN] " & qtyProdplan & " [QTY MPP] " & qtyTotMpp & vbCrLf
                If Val(.TextMatrix(rowPlan, 9)) >= 0 Then
                    .Row = rowPlan
                    .Col = 6
                    If .CellBackColor = RGB(165, 240, 165) Then
                        posAssy = Val(.TextMatrix(rowPlan + 3, 0))
                        .Col = 7
                        If .CellBackColor = RGB(220, 220, 75) Then
                            stExec = "UPDATE"
                        Else
                            stExec = "SKIP"
                        End If
                    Else
                        lastPos = lastPos + 1
                        posAssy = lastPos
                        .TextMatrix(rowPlan + 3, 0) = Format(posAssy, "00")
                        stExec = "INSERT"
                    End If
                End If
                For iLoop = 10 To .Cols - 1
                    On Error Resume Next
                    affMpp = 0
                    stSkip = False
                    If .TextMatrix(1, iLoop) <> "" Then
                        qtyMppWP = Val(.TextMatrix(rowPlan, iLoop))
                        wpID = Format(Val(.TextMatrix(1, iLoop)), "00")
                        wpMPP = wpID & "/" & Right(sPeriod, 2)
                        wpDate = .TextMatrix(9, iLoop)
                        wpColor = Val(.TextMatrix(7, iLoop))
                        wpHoldQty = Val(.TextMatrix(rowPlan + 3, iLoop))
                        wpHoldSeri = .TextMatrix(rowPlan + 4, iLoop)
                        wpHoldWo = .TextMatrix(rowPlan + 5, iLoop)
                        If qtyMppWP > 0 And stExec <> "SKIP" Then
                            stLog = ""
                            woID = Format(posAssy, "00") & "W/P" & wpID & Right(sPeriod, 2) & Mid(sPeriod, 3, 2) & "-" & arrLine(cmbLine.ListIndex, 4)
                            If stExec = "INSERT" Then
                                seriUrut = Format(Val(seriUrut) + 1, "00000")
                                SerialNo = seriHead & seriUrut
                                Con.Execute "insert into wip_trx_mpp (serial_mpp, periode, assy_no, wp, qty, rev, color, line, wp_date, trx_user, plant, trx_date, wp_id, pos_assy, temp_woc_id) " _
                                    & "values ('" & SerialNo & "', '" & sPeriod & "', '" & getAssyNo & "', '" & wpMPP & "', " & qtyMppWP & ", " & Val(sRev) & ", " & wpColor & ", " _
                                    & "'" & sCdLine & "', '" & wpDate & "', '" & pUserId & "', '2', now(), '" & wpID & "', '" & Format(posAssy, "00") & "', '" & woID & "')", affMpp
                                stWoEx = "[I]"
                            ElseIf stExec = "UPDATE" And wpHoldQty <> qtyMppWP Then
                                If wpHoldQty = 0 Then
                                    seriUrut = Format(Val(seriUrut) + 1, "00000")
                                    SerialNo = seriHead & seriUrut
                                    Con.Execute "insert into wip_trx_mpp (serial_mpp, periode, assy_no, wp, qty, rev, color, line, wp_date, trx_user, plant, trx_date, wp_id, pos_assy, temp_woc_id) " _
                                        & "values ('" & SerialNo & "', '" & sPeriod & "', '" & getAssyNo & "', '" & wpMPP & "', " & qtyMppWP & ", " & Val(sRev) & ", " & wpColor & ", " _
                                        & "'" & sCdLine & "', '" & wpDate & "', '" & pUserId & "', '2', now(), '" & wpID & "', '" & Format(posAssy, "00") & "', '" & woID & "')", affMpp
                                    stWoEx = "[I]"
                                Else
                                    SerialNo = wpHoldSeri
                                    Con.Execute "update wip_trx_mpp set qty = " & qtyMppWP & ", rev = " & Val(sRev) & ", trx_user = '" & pUserId & "', trx_date = now() where serial_mpp = '" & SerialNo & "'", affMpp
                                    stWoEx = "[U]"
                                End If
                            Else
                                stSkip = True
                            End If
                            If affMpp > 0 Then
                                logWIPs = Val(logWIPs) + 1
                                If stWoEx = "[I]" Then
                                    stLog = "[I]"
                                    Con.Execute "insert into woc (woc_id, woc_item_id, woc_req_qty, woc_startqty, woc_start_date, woc_req_date, woc_status, woc_employee, periode, cd_line, wp_id, serial_mpp) " _
                                        & "values ('" & woID & "', '" & getAssyNo & "', " & qtyMppWP & ", " & qtyMppWP & ", '" & wpDate & "', '" & wpDate & "', '1', '" & pUserName & "', '" & sPeriod & "', '" & sCdLine & "', '" & wpID & "', '" & SerialNo & "')", affWoc
                                Else
                                    stLog = "[U]"
                                    Con.Execute "update woc set woc_req_qty = " & qtyMppWP & ", woc_startqty = " & qtyMppWP & " where woc_id = '" & wpHoldWo & "'", affWoc
                                    Con.Execute "delete from wom where wom_woc_id = '" & wpHoldWo & "'"
                                End If
                                If affWoc > 0 Then
                                    logWOCs = Val(logWOCs) + 1
                                    stLog = stLog & "[I]"
                                    Con.Execute "insert into wom (wom_woc_id, wom_item_id, wom_routid, wom_qty_perassy, wom_orireq_qty, wom_req_qty) " _
                                        & "(select '" & woID & "', bom_com_item, bom_routid, bom_qty_perassy,  bom_qty_perassy * " & qtyMppWP & ",  bom_qty_perassy * " & qtyMppWP & " from mst_bom where bom_par_item = '" & getAssyNo & "')", affWom
                                    If affWom > 0 Then
                                        logWOMs = Val(logWOMs) + affWom
                                    Else
                                        logWOMs = Val(logWOMs) + 1
                                    End If
                                Else
                                    logWOCf = Val(logWOCf) + 1
                                    stLog = stLog & "[x]"
                                End If
                            ElseIf stSkip = True Then
                                stLog = "[S][-]"
                            Else
                                logWIPf = Val(logWIPf) + 1
                                logWOCf = Val(logWOCf) + 1
                                stLog = "[x][-]"
                            End If
                            debugStr = debugStr & stLog & "[Serial MPP] " & SerialNo & " [WO ID] " & woID & " [WP] " & wpMPP _
                            & " | [Date] " & wpDate _
                            & " | [Qty] " & qtyMppWP & " [Total Material] " & affWom & vbCrLf
                            Sleep 10
                        Else
                            If qtyMppWP = 0 And wpHoldQty > 0 And stExec <> "SKIP" Then
                                Con.Execute "delete from wom where wom_woc_id = '" & wpHoldWo & "'"
                                Con.Execute "delete from woc where woc_id = '" & wpHoldWo & "'"
                                Con.Execute "delete from wip_trx_wos where trxwos_wdscode in (select serial_wds from wip_trx_wds where serial_mpp = '" & wpHoldSeri & "')"
                                Con.Execute "delete from wip_trx_wds_ser where serial_mpp = '" & wpHoldSeri & "'"
                                Con.Execute "delete from wip_trx_wds where serial_mpp = '" & wpHoldSeri & "'"
                                Con.Execute "delete from wip_trx_mpp where serial_mpp = '" & wpHoldSeri & "'"
                                debugStr = debugStr & "[D][D][DELETE: " & wpHoldSeri & "]" & " [WO ID] " & wpHoldWo & " [WP] " & wpMPP & vbCrLf
                            Else
                                debugStr = debugStr & "[-][-][SKIP]" & " | [WP] " & wpMPP & vbCrLf
                            End If
                        End If
                    End If
                Next
            Else
                debugStr = debugStr & "<< NO  " & i + 1 & " >> [PRODPLAN >= QTY MPP] : FALSE --- " _
                    & " [ASSY NO] " & getAssyNo & " [PRODPLAN] " & qtyProdplan & " [QTY MPP] " & qtyTotMpp & vbCrLf
            End If
            debugStr = debugStr & vbCrLf
            progBar.value = i + 1
        Next
        'MsgBox "SAVED", vbInformation, "Information..."
        stGenerate = "SAVED"
        repNotes = txtNote
        txtLog = debugStr
        cmdLog.Visible = True
        End With
    
'    Dim locDebug As String
'    locDebug = App.Path & "/DEBUG/debug_mpp.log"
'    Open locDebug For Output As #1
'    Print #1, debugStr
'    Close #1
Exit Sub
errSave:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Save: " & Err.Number
        cmdSave.Enabled = False
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call activeTheme(skn, Me)
    Call BukaKoneksi
    txtYear = Year(Now)
    getListMonth
    getListLine
    dtWPStart.value = DateSerial(Year(Now), Month(Now), 1)
    dtWPEnd.value = Date
    Call WheelHook(Me.hWnd)
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xPos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hWnd, xPos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, xPos, Ypos ', 9
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
End Sub

Private Sub LoadDataMPP(ByVal line As String, ByVal period As String, ByVal dtFrom As Date, ByVal dtTo As Date)
    Dim dateCount As Integer
    Dim totalDay As Integer
    Dim allDay As Integer
    Dim offDay As Integer
    Dim iAddDate As Integer
    Dim stAddDate As Boolean
    stAddDate = True
    iAddDate = 4
    queOffset = 0
    rowHeader = 10
    totalWP = 0
    Do Until stAddDate = False
        iAddDate = iAddDate + 1
        Set RsGet = Con.Execute("select count(*) from plansys_setoffday where work_date between '" & Format(dtTo + 1, "YYYY-MM-DD") & "' and '" & Format(dtTo + iAddDate, "YYYY-MM-DD") & "' and work_status = true")
        If RsGet.Fields(0) = 5 Then
            stAddDate = False
            dtTo = dtTo + iAddDate
        End If
        If iAddDate > 31 Then
            MsgBox "Silahkan Setting Off Day Terlebih Dahulu.", vbExclamation, "Warning..."
            RsGet.Close
            Exit Sub
        End If
        RsGet.Close
    Loop
    dateCount = (dtTo - dtFrom) + 1
    With MSFlexGridMPP
        .Clear
        .MergeCells = flexMergeFree
        
        Set RsGet = Con.Execute("select * from plansys_setoffday where work_date between '" & Format(dtFrom, "YYYY-MM-DD") & "' and '" & Format(dtTo, "YYYY-MM-DD") & "' order by work_date")
        totalDay = RsGet.RecordCount
        If dateCount <> totalDay Then
            MsgBox "Silahkan Setting Off Day Terlebih Dahulu.", vbExclamation, "Warning..."
            Exit Sub
        End If
        iLoop = 9
        .Cols = iLoop + totalDay + 1
        .rows = rowHeader + 1
        allDay = 0
        offDay = 0
        totalWP = 0
        Do Until RsGet.EOF
            iLoop = iLoop + 1
            allDay = allDay + 1
            .Col = iLoop
            If iLoop Mod 2 = 1 Then
                .TextMatrix(0, iLoop) = "W/P"
            Else
                .TextMatrix(0, iLoop) = " W/P "
            End If
            .TextMatrix(7, iLoop) = Weekday(RsGet!work_date)
            .TextMatrix(8, iLoop) = Format(RsGet!work_date, "DD MMM")
            .TextMatrix(9, iLoop) = Format(RsGet!work_date, "YYYY-MM-DD")
            .RowHeight(9) = 15
            If RsGet!work_status = 1 Then
                If RsGet!work_date <= (dtTo - iAddDate) Then 'CUT
                    totalWP = totalWP + 1
                    .TextMatrix(1, iLoop) = allDay - offDay
                    sTotalWP = totalWP
                End If
                If RsGet!work_date > dtFrom + offDay And totalWP >= (allDay - offDay - 1) Then 'JOINT
                    .TextMatrix(2, iLoop) = allDay - offDay - 1
                End If
                If RsGet!work_date > dtFrom + offDay + 1 And totalWP >= (allDay - offDay - 2) Then 'W/C
                    .TextMatrix(3, iLoop) = allDay - offDay - 2
                End If
                If RsGet!work_date > dtFrom + offDay + 2 And totalWP >= (allDay - offDay - 3) Then 'HAV
                    .TextMatrix(4, iLoop) = allDay - offDay - 3
                End If
                If RsGet!work_date > dtFrom + offDay + 3 And totalWP >= (allDay - offDay - 4) Then 'CP
                    .TextMatrix(5, iLoop) = allDay - offDay - 4
                End If
                If RsGet!work_date > dtFrom + offDay + 4 Then 'FG
                    .TextMatrix(6, iLoop) = allDay - offDay - 5
                End If
            Else
                offDay = offDay + 1
                For i = 0 To 8
                    .Row = i
                    .CellBackColor = RGB(255, 222, 222)
                    .TextMatrix(0, iLoop) = "-"
                Next
            End If
            RsGet.MoveNext
        Loop
        RsGet.Close
        
        .FixedCols = 10
        .FixedRows = 9
        .RowHeightMin = 300
        
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        .MergeRow(3) = True
        .MergeRow(4) = True
        .MergeRow(5) = True
        .MergeRow(6) = True
        .MergeRow(7) = True
        
        .ColWidth(0) = 600
        .ColWidth(1) = 2400
        .ColWidth(2) = 2400
        .ColWidth(3) = 600
        .ColWidth(4) = 600
        .ColWidth(5) = 600
        .ColWidth(6) = 600
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        For i = 0 To .Cols - 1
            For iROW = 0 To 8
                .Row = iROW
                .Col = i
                .CellAlignment = flexAlignCenterCenter
            Next
        Next
        
        .TextMatrix(0, 0) = "NO"
        .TextMatrix(1, 0) = "NO"
        .TextMatrix(2, 0) = "NO"
        .TextMatrix(3, 0) = "NO"
        .TextMatrix(4, 0) = "NO"
        .TextMatrix(5, 0) = "NO"
        .TextMatrix(6, 0) = "NO"
        .TextMatrix(7, 0) = "NO"
        .TextMatrix(8, 0) = "NO"
        
        .TextMatrix(0, 1) = "ASSY NO"
        .TextMatrix(1, 1) = "ASSY NO"
        .TextMatrix(2, 1) = "ASSY NO"
        .TextMatrix(3, 1) = "ASSY NO"
        .TextMatrix(4, 1) = "ASSY NO"
        .TextMatrix(5, 1) = "ASSY NO"
        .TextMatrix(6, 1) = "ASSY NO"
        .TextMatrix(7, 1) = "ASSY NO"
        .TextMatrix(8, 1) = "ASSY NO"
        
        .TextMatrix(0, 2) = "ASSY NAME"
        .TextMatrix(1, 2) = "ASSY NAME"
        .TextMatrix(2, 2) = "ASSY NAME"
        .TextMatrix(3, 2) = "ASSY NAME"
        .TextMatrix(4, 2) = "ASSY NAME"
        .TextMatrix(5, 2) = "ASSY NAME"
        .TextMatrix(6, 2) = "ASSY NAME"
        .TextMatrix(7, 2) = "ASSY NAME"
        .TextMatrix(8, 2) = "ASSY NAME"
        
        .TextMatrix(0, 3) = TextToDown("CCT")
        .TextMatrix(1, 3) = TextToDown("CCT")
        .TextMatrix(2, 3) = TextToDown("CCT")
        .TextMatrix(3, 3) = TextToDown("CCT")
        .TextMatrix(4, 3) = TextToDown("CCT")
        .TextMatrix(5, 3) = TextToDown("CCT")
        .TextMatrix(6, 3) = TextToDown("CCT")
        .TextMatrix(7, 3) = TextToDown("CCT")
        .TextMatrix(8, 3) = TextToDown("CCT")
        
        .TextMatrix(0, 4) = TextToDown("CRIMPING")
        .TextMatrix(1, 4) = TextToDown("CRIMPING")
        .TextMatrix(2, 4) = TextToDown("CRIMPING")
        .TextMatrix(3, 4) = TextToDown("CRIMPING")
        .TextMatrix(4, 4) = TextToDown("CRIMPING")
        .TextMatrix(5, 4) = TextToDown("CRIMPING")
        .TextMatrix(6, 4) = TextToDown("CRIMPING")
        .TextMatrix(7, 4) = TextToDown("CRIMPING")
        .TextMatrix(8, 4) = TextToDown("CRIMPING")
        
        .TextMatrix(0, 5) = TextToDown("JOINT")
        .TextMatrix(1, 5) = TextToDown("JOINT")
        .TextMatrix(2, 5) = TextToDown("JOINT")
        .TextMatrix(3, 5) = TextToDown("JOINT")
        .TextMatrix(4, 5) = TextToDown("JOINT")
        .TextMatrix(5, 5) = TextToDown("JOINT")
        .TextMatrix(6, 5) = TextToDown("JOINT")
        .TextMatrix(7, 5) = TextToDown("JOINT")
        .TextMatrix(8, 5) = TextToDown("JOINT")
        
        .TextMatrix(0, 6) = TextToDown("KET")
        .TextMatrix(1, 6) = TextToDown("KET")
        .TextMatrix(2, 6) = TextToDown("KET")
        .TextMatrix(3, 6) = TextToDown("KET")
        .TextMatrix(4, 6) = TextToDown("KET")
        .TextMatrix(5, 6) = TextToDown("KET")
        .TextMatrix(6, 6) = TextToDown("KET")
        .TextMatrix(7, 6) = TextToDown("KET")
        .TextMatrix(8, 6) = TextToDown("KET")
        
        .TextMatrix(0, 7) = " "
        .TextMatrix(1, 7) = "CUT"
        .TextMatrix(2, 7) = "JOINT"
        .TextMatrix(3, 7) = "W/C"
        .TextMatrix(4, 7) = "HAV"
        .TextMatrix(5, 7) = "CP"
        .TextMatrix(6, 7) = "F/G"
        .TextMatrix(7, 7) = "COLOR"
        .TextMatrix(8, 7) = "QTY"
        
        .TextMatrix(0, 8) = " "
        .TextMatrix(1, 8) = "CUT"
        .TextMatrix(2, 8) = "JOINT"
        .TextMatrix(3, 8) = "W/C"
        .TextMatrix(4, 8) = "HAV"
        .TextMatrix(5, 8) = "CP"
        .TextMatrix(6, 8) = "F/G"
        .TextMatrix(7, 8) = "COLOR"
        .TextMatrix(8, 8) = "F/G STOCK"
        
        .TextMatrix(0, 9) = " "
        .TextMatrix(1, 9) = "CUT"
        .TextMatrix(2, 9) = "JOINT"
        .TextMatrix(3, 9) = "W/C"
        .TextMatrix(4, 9) = "HAV"
        .TextMatrix(5, 9) = "CP"
        .TextMatrix(6, 9) = "F/G"
        .TextMatrix(7, 9) = "COLOR"
        .TextMatrix(8, 9) = "T. MPP"
    End With
    Set RsGet = Con.Execute("select count(a.assy_no) c_data from ltpp_generate a " _
        & "inner join mst_item_line b on a.assy_no = b.item_id and b.cd_line_1 = '" & arrLine(cmbLine.ListIndex, 0) & "' " _
        & "inner join mst_item c on a.assy_no = c.item_id " _
        & "left join plansys_schedule d on d.assy_no = a.assy_no and d.period = a.period " _
        & "where a.period = '" & setPeriod & "' and rev = (select max(rev) from ltpp_generate where period = '" & setPeriod & "')")
    If RsGet!c_data = 0 Then
        progBar.Max = 1
    Else
        progBar.Max = RsGet!c_data
    End If
    progBar.value = 0
    RsGet.Close
    queryTime = True
End Sub

Function TextToDown(Str As String) As String
    Dim tC As Integer
    Dim sConvert As String
    sConvert = ""
    For tC = 1 To Len(Str)
        sConvert = sConvert & Mid(Str, tC, 1) & vbCrLf
    Next
    TextToDown = sConvert
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call WheelUnHook(Me.hWnd)
End Sub

Private Sub MSFlexGridMPP_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Or KeyCode = 46 Or KeyCode = 8 Or KeyCode = 13) Then
        If (KeyCode >= 96 And KeyCode <= 105) Then
            KeyCode = KeyCode - 48
        End If
        flexEditor MSFlexGridMPP, KeyCode
    End If
End Sub

Private Sub flexEditor(argFlexGrid As MSFlexGrid, KeyCode As Integer)
On Error GoTo errEditor
    Dim textHolder As String
    Dim qtyHolder As Long
    Dim colBack As Integer
    Dim rowBack As Integer
    Dim gWO As String
    Dim gSeri As String
    textHolder = argFlexGrid.Text
    With argFlexGrid
        colBack = .Col
        rowBack = .Row
        If stEditor = False Then Exit Sub
        If KeyCode = 13 Or KeyCode = 9 Then Exit Sub
        If Not Trim(.TextMatrix(.Row, 6)) = "PLAN" And Not Trim(.TextMatrix(.Row, 6)) = "SCH" Then Exit Sub
        If Not Trim(.TextMatrix(.Row, 6)) = "SCH" Then
            If Not (Val(.TextMatrix(.Row + 1, 7)) > 0 And Val(.TextMatrix(1, .Col)) > 0) Then Exit Sub
        End If
        If .Text = "-" Then Exit Sub
        If .CellBackColor = RGB(225, 100, 100) Then Exit Sub
        
        gWO = .TextMatrix(.Row + 5, .Col)
        gSeri = .TextMatrix(.Row + 4, .Col)
        qtyHolder = Val(.TextMatrix(.Row + 3, .Col))
        If RTrim(gWO) <> "-" And .CellBackColor <> RGB(220, 220, 75) Then
            'Set RsGet = Con.Execute("select kanbanrm_woc_id from serial_detail_kanbanrm where kanbanrm_woc_id = '" & gWO & "' limit 1")
            Set RsGet = Con.Execute("select distinct serial_mpp from wip_trx_wds where serial_mpp = '" & gSeri & "' and precut_sts = '1'")
            If Not RsGet.EOF Then
                .CellBackColor = RGB(225, 100, 100)
                Exit Sub
            Else
                Set RsBantu = Con.Execute("select kanbanrm_woc_id from serial_detail_kanbanrm where kanbanrm_woc_id = '" & gWO & "' limit 1")
                If Not RsBantu.EOF Then
                    .CellBackColor = RGB(225, 100, 100)
                    Exit Sub
                End If
            End If
            RsGet.Close
        End If
        
        If KeyCode = 8 Then 'backspace
            If Len(Trim(argFlexGrid.Text)) <> 0 Then 'if not empty
                .Text = Val(Left(.Text, (Len(.Text) - 1))) 'Removing a character from the right
            End If
        ElseIf KeyCode = 46 Then
            .Text = 0
        Else
            If Val(.Text) = 0 Then
                .Text = Chr(KeyCode)
            Else
                .Text = .Text + Chr(KeyCode)
            End If
        End If
        If .Text <> textHolder Then
            stGenerate = "EDITING..."
            repNotes = "<< EDITING... >>"
            If Val(.Text) = qtyHolder And Trim(.TextMatrix(.Row, 6)) = "PLAN" Then
                .CellBackColor = vbDefault
            Else
                .CellBackColor = RGB(220, 220, 75)
            End If
            'change footer
            If Trim(.TextMatrix(.Row, 6)) = "PLAN" Then
                .TextMatrix(rowPosFooter, .Col) = 0
                .TextMatrix(rowPosFooter + 1, .Col) = 0
                For i = 0 To Val(sLoadData) - 1
                    'total MPP
                    .TextMatrix(rowPosFooter, .Col) = Val(.TextMatrix(rowPosFooter, .Col)) + Val(.TextMatrix(rowHeader + (i * 7) + 1, .Col))
                    'total Cutting
                    .TextMatrix(rowPosFooter + 1, .Col) = Val(.TextMatrix(rowPosFooter + 1, .Col)) + (Val(.TextMatrix(rowHeader + (i * 7) + 1, .Col)) * Val(.TextMatrix(rowHeader + (i * 7) + 1, 3)))
                Next
                .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 9)) - Val(textHolder) + Val(.Text)
                .TextMatrix(rowPosFooter, 9) = Val(.TextMatrix(rowPosFooter, 9)) - Val(textHolder) + Val(.Text)
                .TextMatrix(rowPosFooter + 1, 9) = Val(.TextMatrix(rowPosFooter + 1, 9)) - (Val(textHolder) * Val(.TextMatrix(.Row, 3))) + (Val(.Text) * Val(.TextMatrix(.Row, 3)))
                .TextMatrix(rowPosFooter + 5, .Col) = Round(Val(.TextMatrix(rowPosFooter + 1, .Col)) * 100 / Val(sCapCut), 2) & "%"
            ElseIf Trim(.TextMatrix(.Row, 6)) = "SCH" Then
                .TextMatrix(.Row, 7) = Val(.TextMatrix(.Row, 7)) - Val(textHolder) + Val(.Text)
            End If
        End If
        
        .Col = 7
        .CellBackColor = RGB(220, 220, 75)
        If Val(.TextMatrix(.Row, 9)) > Val(.TextMatrix(.Row, 7)) Then
            .Col = 9
            .CellBackColor = RGB(225, 100, 100)
        Else
            .Col = 9
            .CellBackColor = vbDefault
        End If
        .Col = colBack
    End With
Exit Sub
errEditor:
    If Err.Number = 11 Then Resume Next
    If Err.Number = 6 Then Resume Next
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Editor: " & Err.Number
    End If
End Sub

Private Sub queryTime_Timer()
On Error GoTo errQuery
    Dim posList As Integer
    Dim RemProdPlan As Double
    Dim qtySch As Double
    Dim lotQty As Long
    Dim lotSpace As Long
    Dim stfgWP As Boolean
    Dim fgStock As Long
    Dim fgWP As Long
    Dim fgSch As Long
    Dim fgPlan As Long
    Dim schQty As Long
    Dim planWP As Long
    Dim planQty As Long
    Dim roundDiv As Long
    Dim rStat As Long
    Dim rWPid As String
    
    Set RsGet = Con.Execute("select a.ltpp_doc, a.rev, a.assy_no, a.item_name, a.cct, coalesce(lotqty, 0) lotqty, prod_plan_1, prod_plan_2, prod_plan_3, prod_plan_4, a.t_stock, a.bal_1, coalesce(d.total_qty, 0) del_schedule from ltpp_generate a " _
        & "inner join mst_item_line b on a.assy_no = b.item_id and b.cd_line_1 = '" & arrLine(cmbLine.ListIndex, 0) & "' " _
        & "inner join mst_item c on a.assy_no = c.item_id " _
        & "left join plansys_schedule d on d.assy_no = a.assy_no and d.period = a.period " _
        & "where a.period = '" & setPeriod & "' and rev = (select max(rev) from ltpp_generate where period = '" & setPeriod & "') order by assy_no limit 1 offset " & queOffset)
    If RsGet.RecordCount > 0 Then
        txtRev = RsGet!rev
        sRev = RsGet!rev
        queOffset = queOffset + 1
        sLoadData = queOffset
        progBar.value = sLoadData
        posList = ((queOffset - 1) * 7)
        With MSFlexGridMPP
            .rows = rowHeader + (queOffset * 7)
            .RowHeight(rowHeader + posList + 4) = 0
            .RowHeight(rowHeader + posList + 5) = 0
            .RowHeight(rowHeader + posList + 6) = 0
            
            .TextMatrix(rowHeader + posList, 0) = queOffset
            .TextMatrix(rowHeader + posList + 1, 0) = queOffset
            .TextMatrix(rowHeader + posList + 2, 0) = queOffset
            .TextMatrix(rowHeader + posList + 3, 0) = queOffset
            
            .TextMatrix(rowHeader + posList, 1) = RTrim(RsGet!assy_no)
            .TextMatrix(rowHeader + posList + 1, 1) = RTrim(RsGet!assy_no)
            .TextMatrix(rowHeader + posList + 2, 1) = RTrim(RsGet!assy_no)
            .TextMatrix(rowHeader + posList + 3, 1) = RTrim(RsGet!assy_no)
            
            .TextMatrix(rowHeader + posList, 2) = RTrim(RsGet!item_name)
            .TextMatrix(rowHeader + posList + 1, 2) = RTrim(RsGet!item_name)
            .TextMatrix(rowHeader + posList + 2, 2) = RTrim(RsGet!item_name)
            .TextMatrix(rowHeader + posList + 3, 2) = RTrim(RsGet!item_name)
            
            .TextMatrix(rowHeader + posList, 3) = RsGet!cct
            .TextMatrix(rowHeader + posList + 1, 3) = RsGet!cct
            .TextMatrix(rowHeader + posList + 2, 3) = RsGet!cct
            .TextMatrix(rowHeader + posList + 3, 3) = RsGet!cct
            
            .TextMatrix(rowHeader + posList, 4) = 0
            .TextMatrix(rowHeader + posList + 1, 4) = 0
            .TextMatrix(rowHeader + posList + 2, 4) = 0
            .TextMatrix(rowHeader + posList + 3, 4) = 0
            
            .TextMatrix(rowHeader + posList, 5) = 0
            .TextMatrix(rowHeader + posList + 1, 5) = 0
            .TextMatrix(rowHeader + posList + 2, 5) = 0
            .TextMatrix(rowHeader + posList + 3, 5) = 0
            
            .TextMatrix(rowHeader + posList, 6) = "SCH"
            .TextMatrix(rowHeader + posList + 1, 6) = "PLAN"
            .TextMatrix(rowHeader + posList + 2, 6) = "LOT"
            .TextMatrix(rowHeader + posList + 3, 6) = "ACT"
            
            fgStock = RsGet!bal_1
            .TextMatrix(rowHeader + posList + 1, 7) = RsGet!prod_plan_1
            .TextMatrix(rowHeader + posList + 1, 8) = fgStock
            .TextMatrix(rowHeader + posList + 1, 9) = 0
            
        'SCHEDULE---
            qtySch = 0
            Set RsBantu = Con.Execute("select b.*, a.total_qty from plansys_schedule a inner join plansys_schedule_detail b on b.period = a.period and b.assy_no = a.assy_no " _
                & "where a.assy_no = '" & RTrim(RsGet!assy_no) & "' and date_schedule between '" & Format(dtWPStart.value, "YYYY-MM-DD") & "' " _
                & "and '" & Format(dtWPEnd.value, "YYYY-MM-DD") & "' order by date_schedule")
            For i = 10 To .Cols - 1
                If Not RsBantu.EOF Then
                    If Format(RsBantu!date_schedule, "DD MMM") = .TextMatrix(8, i) Then
                        .TextMatrix(rowHeader + posList, 7) = Val(.TextMatrix(rowHeader + posList, 7)) + RsBantu!qty
                        .TextMatrix(rowHeader + posList, i) = RsBantu!qty
                        qtySch = RsBantu!qty
                        RsBantu.MoveNext
                    Else
                        .TextMatrix(rowHeader + posList, i) = "-"
                    End If
                Else
                    .TextMatrix(rowHeader + posList, i) = "-"
                End If
            Next
            RsBantu.Close
            
            RemProdPlan = RsGet!prod_plan_1
            lotQty = RsGet!lotQty
            lotSpace = 0
            fgWP = 0
            fgSch = 0
            fgPlan = 0
            stfgWP = False
            
            .TextMatrix(rowHeader + posList + 2, 7) = lotQty
            
        'PLAN----
            Set RsBantu = Con.Execute("select qty, pos_assy from wip_trx_mpp where periode = '" & sPeriod & "' and assy_no = '" & RTrim(RsGet!assy_no) & "' and line = '" & sCdLine & "' limit 1")
            rStat = RsBantu.RecordCount
            If rStat > 0 Then .TextMatrix(rowHeader + posList + 4, 0) = RsBantu!pos_assy
            RsBantu.Close
            Select Case rStat
            Case 0 'generate
                .Row = rowHeader + posList + 1
                .Col = 6
                .CellBackColor = RGB(220, 220, 75)
                'FG to WP
                iCol = 9
                If RemProdPlan <= 0 Then stfgWP = True
                Do Until stfgWP = True
                    iCol = iCol + 1
                    fgSch = fgSch + Val(.TextMatrix(rowHeader + posList, iCol))
                    If fgSch >= fgStock Then stfgWP = True
                    If Val(.TextMatrix(1, iCol)) = Val(sTotalWP) Then stfgWP = True
                Loop
                fgWP = Val(.TextMatrix(1, iCol))
                If lotQty > 0 And fgWP > 0 Then
                    schQty = -Int(Int(-fgStock / fgWP) / lotQty) * lotQty
                Else
                    schQty = 0
                End If
                For i = 10 To iCol
                    If lotQty > 0 And Val(.TextMatrix(1, i)) > 0 And fgWP > 0 Then
                        If RemProdPlan > schQty Then
                            .TextMatrix(rowHeader + posList + 1, i) = schQty
                            .TextMatrix(rowHeader + posList + 4, iCol) = schQty 'temp qty
                            .TextMatrix(rowHeader + posList + 1, 9) = Val(.TextMatrix(rowHeader + posList + 1, 9)) + schQty
                            fgPlan = fgPlan + Val(.TextMatrix(rowHeader + posList + 1, i))
                            RemProdPlan = RemProdPlan - Val(.TextMatrix(rowHeader + posList + 1, i))
                        Else
                            .TextMatrix(rowHeader + posList + 1, i) = RemProdPlan
                            .TextMatrix(rowHeader + posList + 4, i) = RemProdPlan 'temp qty
                            .TextMatrix(rowHeader + posList + 1, 9) = Val(.TextMatrix(rowHeader + posList + 1, 9)) + RemProdPlan
                            RemProdPlan = 0
                        End If
                        .Row = rowHeader + posList + 1
                        .Col = i
                        .CellBackColor = RGB(190, 220, 220)
                    End If
                Next
                
                planWP = sTotalWP - fgWP
                If lotQty > 0 And planWP > 0 Then
                    planQty = -Int(Int(-RemProdPlan / planWP) / lotQty) * lotQty
                Else
                    planQty = 0
                End If
                'NEXT PLAN
                For i = iCol + 1 To .Cols - 1
                    If lotQty > 0 And Val(.TextMatrix(1, i)) > 0 And planWP > 0 Then
                        .TextMatrix(rowHeader + posList + 5, i) = RemProdPlan
                        If RemProdPlan > planQty Then
                            .TextMatrix(rowHeader + posList + 1, i) = planQty
                            .TextMatrix(rowHeader + posList + 4, i) = planQty 'temp qty
                            .TextMatrix(rowHeader + posList + 1, 9) = Val(.TextMatrix(rowHeader + posList + 1, 9)) + planQty
                            RemProdPlan = RemProdPlan - planQty
                        Else
                            .TextMatrix(rowHeader + posList + 1, i) = RemProdPlan
                            .TextMatrix(rowHeader + posList + 4, i) = RemProdPlan 'temp qty
                            .TextMatrix(rowHeader + posList + 1, 9) = Val(.TextMatrix(rowHeader + posList + 1, 9)) + RemProdPlan
                            RemProdPlan = 0
                        End If
                    End If
                Next
                If Val(.TextMatrix(rowHeader + posList + 1, 9)) > Val(.TextMatrix(rowHeader + posList + 1, 7)) Then
                    .Row = rowHeader + posList + 1
                    .Col = 9
                    .CellBackColor = RGB(225, 100, 100)
                End If
            Case Else 're-generate
                .Row = rowHeader + posList + 1
                .Col = 6
                .CellBackColor = RGB(165, 240, 165)
                For i = 10 To .Cols - 1
                    rWPid = Format(.TextMatrix(1, i), "00")
                    If Val(.TextMatrix(1, i)) > 0 Then
                        Set RsBantu = Con.Execute("select serial_mpp, qty, temp_woc_id from wip_trx_mpp where periode = '" & sPeriod & "' and assy_no = '" & RTrim(RsGet!assy_no) & "' and line = '" & sCdLine & "' and wp_id = '" & rWPid & "'")
                        If Not RsBantu.EOF Then
                            .TextMatrix(rowHeader + posList + 1, i) = RsBantu!qty
                            .TextMatrix(rowHeader + posList + 4, i) = RsBantu!qty 'temp qty
                            .TextMatrix(rowHeader + posList + 5, i) = RsBantu!serial_mpp
                            .TextMatrix(rowHeader + posList + 6, i) = RsBantu!temp_woc_id
                            .TextMatrix(rowHeader + posList + 1, 9) = Val(.TextMatrix(rowHeader + posList + 1, 9)) + Val(RsBantu!qty)
                        Else
                            .TextMatrix(rowHeader + posList + 1, i) = "0"
                            .TextMatrix(rowHeader + posList + 4, i) = "0" 'temp qty
                            .TextMatrix(rowHeader + posList + 5, i) = "-"
                            .TextMatrix(rowHeader + posList + 6, i) = "-"
                        End If
                        RsBantu.Close
                    End If
                Next
                If Val(.TextMatrix(rowHeader + posList + 1, 9)) > Val(.TextMatrix(rowHeader + posList + 1, 7)) Then
                    .Col = 9
                    .CellBackColor = RGB(225, 100, 100)
                End If
            End Select
        End With
    Else
        queryTime = False
        footerMPP queOffset
        stGenerate = "GENERATED"
    End If
    RsGet.Close
errQuery:
    If Err.Number <> 0 Then
        queryTime = False
        lblStatusGenerate.ForeColor = vbRed
        cmdGenerate.Enabled = True
        lblStatusGenerate = "Error: " & Err.Number
    End If
End Sub

Private Sub footerMPP(ByVal countData As Integer)
    rowPosFooter = rowHeader + (countData * 7) + 1
    With MSFlexGridMPP
        .rows = rowPosFooter + 8
        .TextMatrix(rowPosFooter, 1) = "TOTAL MPP"
        .TextMatrix(rowPosFooter + 1, 1) = "TOTAL CUTTING"
        .TextMatrix(rowPosFooter + 2, 1) = "TOTAL CRIMPING"
        .TextMatrix(rowPosFooter + 3, 1) = "TOTAL JOINT"
        
        .TextMatrix(rowPosFooter + 5, 1) = "TOTAL CAPACITY"
        .TextMatrix(rowPosFooter + 6, 1) = "TOTAL CAPACITY "
        .TextMatrix(rowPosFooter + 7, 1) = "TOTAL CAPACITY  "
        .TextMatrix(rowPosFooter + 5, 2) = "CUTTING PROCESS"
        .TextMatrix(rowPosFooter + 6, 2) = "CRIMPING PROCESS"
        .TextMatrix(rowPosFooter + 7, 2) = "JOINTING PROCESS"
        
        'CAPACITY MACHINE
        Set RsBantu = Con.Execute("select target from wip_mst_target_mk where mk_to = (select max(mk_to) from wip_mst_target_mk)")
        If Not RsBantu.EOF Then
            sCapCut = Val(RsBantu!Target) * 8
        Else
            sCapCut = 0
        End If
        RsBantu.Close
        Set RsBantu = Con.Execute("select coalesce(sum(" & sCapCut & " * shift), 0) cap_cut from wip_mst_machine where cd_line = '" & sCdLine & "' and proses = '1'")
        If Not RsBantu.EOF Then
            sCapCut = RsBantu!cap_cut
        End If
        RsBantu.Close
        
        'TOTAL MPP
        For iLoop = 10 To .Cols - 1
            If .TextMatrix(1, iLoop) <> "" Then
                For i = 0 To countData - 1
                    .TextMatrix(rowPosFooter, iLoop) = Val(.TextMatrix(rowPosFooter, iLoop)) + Val(.TextMatrix(rowHeader + (i * 7) + 1, iLoop))
                    .TextMatrix(rowPosFooter + 1, iLoop) = Val(.TextMatrix(rowPosFooter + 1, iLoop)) + (Val(.TextMatrix(rowHeader + (i * 7) + 1, iLoop)) * Val(.TextMatrix(rowHeader + (i * 7) + 1, 3)))
                Next
                .TextMatrix(rowPosFooter, 9) = Val(.TextMatrix(rowPosFooter, 9)) + Val(.TextMatrix(rowPosFooter, iLoop))
                .TextMatrix(rowPosFooter + 1, 9) = Val(.TextMatrix(rowPosFooter + 1, 9)) + Val(.TextMatrix(rowPosFooter + 1, iLoop))
                If Val(sCapCut) > 0 Then
                    .TextMatrix(rowPosFooter + 5, iLoop) = Round(Val(.TextMatrix(rowPosFooter + 1, iLoop)) * 100 / Val(sCapCut), 2) & "%"
                Else
                    .TextMatrix(rowPosFooter + 5, iLoop) = "0%"
                End If
            End If
        Next
    End With
    stEditor = True
    cmdGenerate.Enabled = True
    cmdSave.Enabled = True
    cmdExcel.Enabled = True
    cmdprint.Enabled = True
End Sub

Private Sub clearSummary()
    FrameLog.Visible = False
    cmdLog.Visible = False
    sDocno = ""
    sPeriod = ""
    sWPFr = ""
    sWPTo = ""
    sRev = ""
    sLine = ""
    sCdLine = ""
    sTotalWP = 0
    sLoadData = 0
    sCapCut = 0
    sCapCrimp = 0
    sCapJoint = 0
    lblStatusGenerate.ForeColor = vbBlack
    lblStatusGenerate = ""
    txtDocNo = ""
    txtRev = ""
End Sub

