VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form_GenerateLTPP 
   Caption         =   "Generate LTPP"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19620
   Icon            =   "Form_GenerateLTPP.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   19620
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picunprocess 
      Height          =   3615
      Left            =   6120
      ScaleHeight     =   3555
      ScaleWidth      =   6255
      TabIndex        =   113
      Top             =   4440
      Visible         =   0   'False
      Width           =   6320
      Begin MSComctlLib.ListView lv1 
         Height          =   2775
         Left            =   0
         TabIndex        =   114
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label40 
         Caption         =   "Please revise data in the template file"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   116
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Inactive Item List"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   115
         Top             =   0
         Width           =   5775
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   2400
   End
   Begin VB.PictureBox pic0 
      BackColor       =   &H0000FF00&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   15195
      TabIndex        =   108
      Top             =   4800
      Visible         =   0   'False
      Width           =   15255
      Begin MSFlexGridLib.MSFlexGrid grid0 
         Height          =   2655
         Left            =   120
         TabIndex        =   111
         Top             =   480
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   4683
         _Version        =   393216
         BackColorBkg    =   8438015
         AllowUserResizing=   1
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
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14760
         TabIndex        =   110
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   109
         Top             =   0
         Width           =   14775
      End
   End
   Begin VB.PictureBox FrameBSet 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8760
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   105
      Top             =   240
      Width           =   735
      Begin VB.OptionButton opSA 
         Caption         =   "SA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton opUL 
         Caption         =   "UL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   0
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.PictureBox frameGenerate 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.PictureBox picIndicator 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   360
         ScaleHeight     =   735
         ScaleWidth      =   855
         TabIndex        =   112
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox chkReGenerate 
         Caption         =   "RE-GENERATE"
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
         Left            =   4560
         TabIndex        =   73
         Top             =   3360
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
         TabIndex        =   46
         Top             =   840
         Width           =   2175
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
         Left            =   6240
         TabIndex        =   9
         Top             =   3240
         Width           =   1815
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
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   1800
         Width           =   6015
      End
      Begin VB.ComboBox cmbRev 
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
         TabIndex        =   6
         Top             =   1320
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
         Top             =   840
         Width           =   3735
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
         TabIndex        =   1
         Top             =   360
         Width           =   6015
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
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  Revision"
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
         TabIndex        =   5
         Top             =   1320
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
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  LTPP Period"
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
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "Form_GenerateLTPP.frx":000C
      Top             =   0
   End
   Begin VB.PictureBox frameFORM 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin MSComDlg.CommonDialog comSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog comDialogUpload 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox gridFrame 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5175
      ScaleWidth      =   19455
      TabIndex        =   47
      Top             =   4320
      Width           =   19455
      Begin VB.CommandButton cmdFindAssy 
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   94
         Top             =   0
         Width           =   855
      End
      Begin VB.PictureBox FrameGenerateHeader 
         Height          =   495
         Left            =   9360
         ScaleHeight     =   435
         ScaleWidth      =   7875
         TabIndex        =   75
         Top             =   0
         Width           =   7935
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "FC M4(%)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6360
            TabIndex        =   93
            Top             =   0
            Width           =   855
         End
         Begin VB.Label g_fc4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7320
            TabIndex        =   92
            Top             =   0
            Width           =   495
         End
         Begin VB.Label g_hkw 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   5640
            TabIndex        =   91
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label36 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "HKW 4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4920
            TabIndex        =   90
            Top             =   0
            Width           =   615
         End
         Begin VB.Label g_hkw 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   5640
            TabIndex        =   89
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label34 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "HKW 5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4920
            TabIndex        =   88
            Top             =   240
            Width           =   615
         End
         Begin VB.Label g_hkw 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   87
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "HKW 2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3240
            TabIndex        =   86
            Top             =   0
            Width           =   615
         End
         Begin VB.Label g_hkw 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   3960
            TabIndex        =   85
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "HKW 3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3240
            TabIndex        =   84
            Top             =   240
            Width           =   615
         End
         Begin VB.Label g_lt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   83
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "L/T"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   82
            Top             =   0
            Width           =   615
         End
         Begin VB.Label g_hkw 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   81
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "HKW 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   80
            Top             =   240
            Width           =   615
         End
         Begin VB.Label g_period 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   79
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "PERIOD"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   735
         End
         Begin VB.Label g_rev 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   77
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "REV"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   76
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdEditMode 
         Appearance      =   0  'Flat
         Caption         =   "EDIT MODE"
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
         Height          =   495
         Left            =   17280
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   0
         Width           =   2055
      End
      Begin VB.OptionButton optA3 
         Caption         =   "A3"
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
         TabIndex        =   72
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optA4 
         Caption         =   "A4"
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
         TabIndex        =   71
         Top             =   0
         Width           =   735
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
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   120
         Width           =   4215
      End
      Begin VB.CommandButton cmdPrintLTPP 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   49
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdExcelLTPP 
         Caption         =   "EXCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   48
         Top             =   0
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridLTPP 
         Height          =   4455
         Left            =   0
         TabIndex        =   50
         Top             =   480
         Width           =   19455
         _ExtentX        =   34316
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483631
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtFindAssy 
         Height          =   195
         Left            =   8400
         TabIndex        =   95
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   8355
      TabIndex        =   67
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton cmdOKHeader 
         Caption         =   "CHANGE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   68
         Top             =   3480
         Width           =   1815
      End
   End
   Begin VB.PictureBox pTriangle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      Picture         =   "Form_GenerateLTPP.frx":0240
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   70
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox logoBEI32 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "Form_GenerateLTPP.frx":05C6
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox FrameUpload 
      ForeColor       =   &H00800000&
      Height          =   4095
      Left            =   8640
      ScaleHeight     =   4035
      ScaleWidth      =   10875
      TabIndex        =   10
      Top             =   120
      Width           =   10935
      Begin VB.CommandButton cmdCancelUpload 
         Caption         =   "CANCEL UPLOAD"
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
         Left            =   1200
         TabIndex        =   63
         Top             =   3240
         Width           =   2295
      End
      Begin TabDlg.SSTab TabLOG 
         Height          =   3135
         Left            =   6120
         TabIndex        =   59
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5530
         _Version        =   393216
         Tab             =   1
         TabHeight       =   423
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "WIP"
         TabPicture(0)   =   "Form_GenerateLTPP.frx":0C03
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "logUploadWIP"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "SO"
         TabPicture(1)   =   "Form_GenerateLTPP.frx":0C1F
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "logUploadSO"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "FC"
         TabPicture(2)   =   "Form_GenerateLTPP.frx":0C3B
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "logUploadFC"
         Tab(2).ControlCount=   1
         Begin VB.TextBox logUploadFC 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   2685
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   62
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox logUploadSO 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   2685
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   61
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox logUploadWIP 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   2685
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   60
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "UPLOAD"
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
         Left            =   3600
         TabIndex        =   11
         Top             =   3240
         Width           =   2295
      End
      Begin MSComctlLib.ProgressBar progressBarWIP 
         Height          =   495
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar progressBarSO 
         Height          =   495
         Left            =   1200
         TabIndex        =   15
         Top             =   1560
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar progressBarFC 
         Height          =   495
         Left            =   1200
         TabIndex        =   17
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label l_hkw 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   9480
         TabIndex        =   65
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HKW5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8760
         TabIndex        =   64
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HKW4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7320
         TabIndex        =   58
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HKW3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7320
         TabIndex        =   57
         Top             =   120
         Width           =   615
      End
      Begin VB.Label l_hkw 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   56
         Top             =   360
         Width           =   615
      End
      Begin VB.Label l_hkw 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   8040
         TabIndex        =   55
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HKW2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   54
         Top             =   360
         Width           =   615
      End
      Begin VB.Label l_hkw 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   6480
         TabIndex        =   53
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "FC M4 (%)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label l_fc4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4680
         TabIndex        =   51
         Top             =   360
         Width           =   615
      End
      Begin VB.Label uploadStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   8760
         TabIndex        =   45
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label countFC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5040
         TabIndex        =   44
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label countSO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label countWIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5040
         TabIndex        =   42
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label l_lt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4680
         TabIndex        =   41
         Top             =   120
         Width           =   615
      End
      Begin VB.Label l_hkw 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   40
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L/T"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   39
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HKW1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   38
         Top             =   120
         Width           =   615
      End
      Begin VB.Label l_period 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label l_rev 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "REVISION"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PERIOD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label prcFC 
         Alignment       =   1  'Right Justify
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5160
         TabIndex        =   33
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label prcSO 
         Alignment       =   1  'Right Justify
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label fFC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label sFC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Success: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Failed: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label fSO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label sSO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Success: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Failed: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label fWIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label sWIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Failed: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Success: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label prcWIP 
         Alignment       =   1  'Right Justify
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5160
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Forecast"
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
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ost SO"
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
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "WIP"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.PictureBox FrameApproval 
      Height          =   4095
      Left            =   8640
      ScaleHeight     =   4035
      ScaleWidth      =   10875
      TabIndex        =   96
      Top             =   120
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton cmdSApproval 
         Caption         =   "SAVE APPROVAL"
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
         Left            =   8760
         TabIndex        =   103
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtDibuat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   101
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtDiperiksa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   99
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtDiketahui 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   97
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "SETTING APPROVAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   104
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label31 
         Caption         =   "Dibuat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   102
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "Diperiksa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   100
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "Diketahui"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   98
         Top             =   1920
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_GenerateLTPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long


Dim bal_1 As Long, ito_1 As Long, del_rate_1 As Long, del_rate_2 As Long, del_rate_3 As Long, _
    del_rate_4 As Long, s_stock_1 As Long, s_stock_2 As Long, s_stock_3 As Long, s_stock_4 As Long, _
    need_1 As Long, prod_plan_1 As Long, bal_2 As Long, ito_2 As Long, need_2 As Long, prod_plan_2 As Long, _
    bal_3 As Long, ito_3 As Long, need_3 As Long, prod_plan_3 As Long, bal_4 As Long, ito_4 As Long, _
    need_4 As Long, prod_plan_4 As Long, bal_end As Long

Dim arrMM(4)    As String
Dim pYear       As String
Dim prev        As String
Dim doc         As String
Dim injLine     As String
Dim holdLine    As String
Dim iRev        As Integer
Dim dtLTPP      As Date
Dim iLoop       As Integer
Dim ltppEDIT    As Boolean
Dim rsAssy      As ADODB.Recordset
Dim rsIter         As ADODB.Recordset
Dim cacheRowAssy As String
Dim stSPart As Integer
Dim stExcel As Boolean
Dim c_cap_p_day As Long
Dim t_Mult As Integer
Dim t_pp As Double


Private Const xlUp As Long = -4162
Private Const xlCenter As Long = -4108
Private Const xlContinuous = 1
Private Const xlEdgeLeft = 7
Private Const xlEdgeBottom = 9
Private Const msoShapeIsoscelesTriangle = 7
Private Const HORZRES = 8
Private Const VERTRES = 10



Private Sub activeFrameSet(st As Boolean)
If st = False Then
    FrameApproval.Visible = False
    FrameUpload.Visible = True
Else
    FrameApproval.Visible = True
    FrameUpload.Visible = False
End If
End Sub

Private Sub cmbLine_Click()
On Error GoTo errGenerateLine
    Dim i As Integer
    Dim iCol As Integer
    If cmbLine.Text = "-ALL-" Then
        injLine = ""
    ElseIf cmbLine.Text = "***NO LINE***" Then
        injLine = "and d.nm_line isnull"
    Else
        injLine = "and upper(d.nm_line) = '" & cmbLine.Text & "'"
    End If
'    Set RsGet = Con.Execute("select a.*, upper(coalesce(c.cust_name, '***NO CUSTOMER***')) cust_name from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
'        & "left join r_customer c on b.cust_id = c.cust_id where a.ltpp_doc = '" & txtDocNo & "' " & injLine & " order by cust_name, a.assy_no")
    Set RsGet = Con.Execute("select a.*, upper(coalesce(d.nm_line, '***NO LINE***')) nm_line, b.st_sparepart, b.prc_safetystock, coalesce(b.item_muloq, 0) item_muloq, upper(coalesce(e.cust_name, '-')) cust_name from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
        & "left join mst_item_line c on b.item_id = c.item_id left join wip_mst_line d on c.cd_line_1 = d.cd_line left join r_customer e on b.cust_id = e.cust_id where a.ltpp_doc = '" & txtDocNo & "' " & injLine & " order by d.nm_line, a.assy_no")
    If Not RsGet.EOF Then
        i = 0
        holdLine = ""
        txtNote = RsGet!Notes
        Call LoadHeaderGrid(RsGet.RecordCount * 2, arrMM(0), arrMM(1), arrMM(2), arrMM(3), arrMM(4))
        dtLTPP = RsGet!ltpp_date
        g_period = RsGet!period
        g_rev = RsGet!rev
        g_lt = RsGet!lt
        g_hkw(1) = RsGet!hkw_1
        g_hkw(2) = RsGet!hkw_2
        g_hkw(3) = RsGet!hkw_3
        g_hkw(4) = RsGet!hkw_4
        g_hkw(5) = RsGet!hkw_5
        g_fc4 = RsGet!fc_m4
        RsGet.MoveFirst
        Do Until RsGet.EOF
            On Error Resume Next
            With MSFlexGridLTPP
                i = i + 1
                If holdLine <> RsGet!nm_line Then
                    holdLine = RsGet!nm_line
                    .TextMatrix(1 + (i * 2), 0) = " " & holdLine
                    .TextMatrix(1 + (i * 2), 1) = " " & holdLine
                    .TextMatrix(1 + (i * 2), 2) = " " & holdLine
                    .TextMatrix(1 + (i * 2), 3) = " " & holdLine
                    .MergeRow(1 + (i * 2)) = True
                    For iCol = 0 To 50
                        .Row = 1 + (i * 2)
                        .Col = iCol
                        .CellBackColor = vbWhite
                        .CellAlignment = flexAlignCenterCenter
                        If iCol > 3 Then
                            .TextMatrix(1 + (i * 2), iCol) = " "
                        End If
                    Next
                Else
                    .RowHeight(1 + (i * 2)) = 0
                End If
                .TextMatrix(2 + (i * 2), 0) = i
                .TextMatrix(2 + (i * 2), 1) = RsGet!assy_no
                .TextMatrix(2 + (i * 2), 2) = RsGet!item_name
                .TextMatrix(2 + (i * 2), 3) = RsGet!cust_name
                .TextMatrix(2 + (i * 2), 4) = RsGet!item_muloq
                .TextMatrix(2 + (i * 2), 5) = RsGet!pp
                .TextMatrix(2 + (i * 2), 6) = RsGet!p1
                .TextMatrix(2 + (i * 2), 7) = RsGet!p2
                .TextMatrix(2 + (i * 2), 8) = RsGet!p3
                .TextMatrix(2 + (i * 2), 9) = RsGet!p4
                .TextMatrix(2 + (i * 2), 10) = RsGet!p5
                .TextMatrix(2 + (i * 2), 11) = RsGet!p6
                .TextMatrix(2 + (i * 2), 12) = RsGet!p7
                .TextMatrix(2 + (i * 2), 13) = RsGet!p8
                .TextMatrix(2 + (i * 2), 14) = RsGet!p9
                .TextMatrix(2 + (i * 2), 15) = RsGet!p10
                .TextMatrix(2 + (i * 2), 16) = RsGet!rep_qa
                .TextMatrix(2 + (i * 2), 17) = RsGet!hav
                .TextMatrix(2 + (i * 2), 18) = RsGet!fg
                .TextMatrix(2 + (i * 2), 19) = RsGet!ost_mpp
                .TextMatrix(2 + (i * 2), 20) = RsGet!t_stock
                .TextMatrix(2 + (i * 2), 21) = RsGet!so_0
                .TextMatrix(2 + (i * 2), 22) = RsGet!bal_1
                .TextMatrix(2 + (i * 2), 23) = RsGet!ito_1
                .TextMatrix(2 + (i * 2), 24) = RsGet!so_1
                .TextMatrix(2 + (i * 2), 25) = RsGet!fc1
                .TextMatrix(2 + (i * 2), 26) = RsGet!del_rate_1
                .TextMatrix(2 + (i * 2), 27) = RsGet!s_stock_1
                .TextMatrix(2 + (i * 2), 28) = RsGet!need_1
                .TextMatrix(2 + (i * 2), 29) = RsGet!prod_plan_1
                .TextMatrix(2 + (i * 2), 30) = RsGet!bal_2
                .TextMatrix(2 + (i * 2), 31) = RsGet!ito_2
                .TextMatrix(2 + (i * 2), 32) = RsGet!so_2
                .TextMatrix(2 + (i * 2), 33) = RsGet!fc2
                .TextMatrix(2 + (i * 2), 34) = RsGet!del_rate_2
                .TextMatrix(2 + (i * 2), 35) = RsGet!s_stock_2
                .TextMatrix(2 + (i * 2), 36) = RsGet!need_2
                .TextMatrix(2 + (i * 2), 37) = RsGet!prod_plan_2
                .TextMatrix(2 + (i * 2), 38) = RsGet!bal_3
                .TextMatrix(2 + (i * 2), 39) = RsGet!ito_3
                .TextMatrix(2 + (i * 2), 40) = RsGet!so_3
                .TextMatrix(2 + (i * 2), 41) = RsGet!fc3
                .TextMatrix(2 + (i * 2), 42) = RsGet!del_rate_3
                .TextMatrix(2 + (i * 2), 43) = RsGet!s_stock_3
                .TextMatrix(2 + (i * 2), 44) = RsGet!need_3
                .TextMatrix(2 + (i * 2), 45) = RsGet!prod_plan_3
                .TextMatrix(2 + (i * 2), 46) = RsGet!bal_4
                .TextMatrix(2 + (i * 2), 47) = RsGet!ito_4
                .TextMatrix(2 + (i * 2), 48) = RsGet!so_4
                .TextMatrix(2 + (i * 2), 49) = RsGet!fc4
                .TextMatrix(2 + (i * 2), 50) = RsGet!del_rate_4
                .TextMatrix(2 + (i * 2), 51) = RsGet!s_stock_4
                .TextMatrix(2 + (i * 2), 52) = RsGet!need_4
                .TextMatrix(2 + (i * 2), 53) = RsGet!prod_plan_4
                .TextMatrix(2 + (i * 2), 54) = RsGet!bal_end
                .TextMatrix(2 + (i * 2), 55) = RsGet!st_sparepart
                .TextMatrix(2 + (i * 2), 56) = "-"
                .TextMatrix(2 + (i * 2), 57) = RsGet!prc_safetystock
            End With
            RsGet.MoveNext
        Loop
        cmdEditMode.Enabled = True
    Else
        Call LoadHeaderGrid(1, "M+0", "M+1", "M+2", "M+3", "M+4")
        cmdEditMode.Enabled = False
        MsgBox "No Data Found!", vbExclamation, "Warning..."
    End If
    RsGet.Close
Exit Sub
errGenerateLine:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error : " & Err.Number
    End If
End Sub

Private Sub cmbMM_Click()
    getRev txtYear & Format(cmbMM.ListIndex + 1, "00")
End Sub

Private Sub getRev(ByVal period As String)
On Error GoTo errRev
    Set RsGet = Con.Execute("select distinct rev from ltpp_header where period = '" & period & "' order by rev")
    cmbRev.Clear
    If Not RsGet.EOF Then
        Do Until RsGet.EOF
            cmbRev.AddItem RsGet!rev
            RsGet.MoveNext
        Loop
    End If
    RsGet.Close
Exit Sub
errRev:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Get Rev: " & Err.Number
    End If
End Sub
    
Private Sub cmdCancelUpload_Click()
On Error GoTo errCancel
    Dim confCancel As String
    confCancel = MsgBox("Period: " & l_period & vbCrLf & "Rev. : " & l_rev & vbCrLf & "Akan Dihapus. Lanjutkan?", vbExclamation + vbYesNo, "Cancel Upload: " & doc)
    If confCancel = vbYes Then
        Me.MousePointer = vbHourglass
        Con.Execute "delete from ltpp_wip where ltpp_doc = '" & doc & "'"
        Con.Execute "delete from ltpp_so where ltpp_doc = '" & doc & "'"
        Con.Execute "delete from ltpp_fc where ltpp_doc = '" & doc & "'"
        Con.Execute "delete from ltpp_header where ltpp_doc = '" & doc & "'"
        cmdCancelUpload.Enabled = False
        clearUploadForm
    End If
    Me.MousePointer = vbDefault
Exit Sub
errCancel:
    Me.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Cancel: " & Err.Number
    End If
End Sub

Private Sub cmdEditMode_Click()
On Error GoTo errEdit
    Select Case cmdEditMode.Caption
        Case "EDIT MODE"
            ltppEDIT = True
            cmdEditMode.Caption = "UPDATE LTPP"
            cmdEditMode.BackColor = vbGreen
        Case "UPDATE LTPP"
            updateLTPP MSFlexGridLTPP
            ltppEDIT = False
            cmdEditMode.BackColor = vbButtonFace
            cmdEditMode.Caption = "EDIT MODE"
    End Select
Exit Sub
errEdit:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Edit: " & Err.Number
    End If
End Sub

Private Sub updateLTPP(ByVal fgrid As MSFlexGrid)
    Dim rUpdate As Integer
    Dim affUpdate As Integer
    rUpdate = 0
    With fgrid
        For iLoop = 4 To fgrid.rows - 1
            If Val(.TextMatrix(iLoop, 0)) <> 0 And .TextMatrix(iLoop, 56) = "edit" Then
                s_stock_1 = RoundNumber(Val(.TextMatrix(iLoop, 55)) * Val(.TextMatrix(iLoop, 25)) / 100)
                s_stock_2 = RoundNumber(Val(.TextMatrix(iLoop, 55)) * Val(.TextMatrix(iLoop, 33)) / 100)
                s_stock_3 = RoundNumber(Val(.TextMatrix(iLoop, 55)) * Val(.TextMatrix(iLoop, 41)) / 100)
                s_stock_4 = RoundNumber(Val(.TextMatrix(iLoop, 55)) * Val(.TextMatrix(iLoop, 49)) / 100)
                Con.Execute "update ltpp_generate set " & "pp = " & Val(.TextMatrix(iLoop, 5)) & ", p1 = " & Val(.TextMatrix(iLoop, 6)) _
                & ", p2 = " & Val(.TextMatrix(iLoop, 7)) & ", p3 = " & Val(.TextMatrix(iLoop, 8)) & ", p4 = " & Val(.TextMatrix(iLoop, 9)) _
                & ", p5 = " & Val(.TextMatrix(iLoop, 10)) & ", p6 = " & Val(.TextMatrix(iLoop, 11)) & ", p7 = " & Val(.TextMatrix(iLoop, 12)) _
                & ", p8 = " & Val(.TextMatrix(iLoop, 13)) & ", p9 = " & Val(.TextMatrix(iLoop, 14)) & ", p10 = " & Val(.TextMatrix(iLoop, 15)) _
                & ", rep_qa = " & Val(.TextMatrix(iLoop, 16)) & ", hav = " & Val(.TextMatrix(iLoop, 17)) & ", fg = " & Val(.TextMatrix(iLoop, 18)) _
                & ", ost_mpp = " & Val(.TextMatrix(iLoop, 19)) & ", t_stock = " & Val(.TextMatrix(iLoop, 20)) _
                & ", fc1 = " & Val(.TextMatrix(iLoop, 25)) & ", fc2 = " & Val(.TextMatrix(iLoop, 33)) & ", fc3 = " & Val(.TextMatrix(iLoop, 41)) _
                & ", fc4 = " & Val(.TextMatrix(iLoop, 49)) & ", t_fc = " & Val(.TextMatrix(iLoop, 25)) + Val(.TextMatrix(iLoop, 33)) + Val(.TextMatrix(iLoop, 41)) + Val(.TextMatrix(iLoop, 49)) _
                & ", so_0 = " & Val(.TextMatrix(iLoop, 21)) & ", bal_1 = " & Val(.TextMatrix(iLoop, 22)) & ", ito_1 = " & Val(.TextMatrix(iLoop, 23)) _
                & ", so_1 = " & Val(.TextMatrix(iLoop, 24)) & ", so_2 = " & Val(.TextMatrix(iLoop, 32)) & ", so_3 = " & Val(.TextMatrix(iLoop, 40)) _
                & ", so_4 = " & Val(.TextMatrix(iLoop, 48)) & ", del_rate_1 = " & Val(.TextMatrix(iLoop, 26)) & ", del_rate_2 = " & Val(.TextMatrix(iLoop, 34)) _
                & ", del_rate_3 = " & Val(.TextMatrix(iLoop, 42)) & ", del_rate_4 = " & Val(.TextMatrix(iLoop, 50)) & ", s_stock_1 = " & s_stock_1 _
                & ", s_stock_2 = " & s_stock_2 & ", s_stock_3 = " & s_stock_3 & ", s_stock_4 = " & s_stock_4 & ", need_1 = " & Val(.TextMatrix(iLoop, 28)) _
                & ", prod_plan_1 = " & Val(.TextMatrix(iLoop, 29)) & ", bal_2 = " & Val(.TextMatrix(iLoop, 30)) & ", ito_2 = " & Val(.TextMatrix(iLoop, 31)) _
                & ", need_2 = " & Val(.TextMatrix(iLoop, 36)) & ", prod_plan_2 = " & Val(.TextMatrix(iLoop, 37)) & ", bal_3 = " & Val(.TextMatrix(iLoop, 38)) _
                & ", ito_3 = " & Val(.TextMatrix(iLoop, 39)) & ", need_3 = " & Val(.TextMatrix(iLoop, 44)) & ", prod_plan_3 = " & Val(.TextMatrix(iLoop, 45)) _
                & ", bal_4 = " & Val(.TextMatrix(iLoop, 46)) & ", ito_4 = " & Val(.TextMatrix(iLoop, 47)) & ", need_4 = " & Val(.TextMatrix(iLoop, 52)) _
                & ", prod_plan_4 = " & Val(.TextMatrix(iLoop, 53)) & ", bal_end = " & Val(.TextMatrix(iLoop, 54)) _
                & " where ltpp_doc = '" & txtDocNo & "' and period = '" & g_period & "' and rev = " & g_rev & " and assy_no = '" & RTrim(.TextMatrix(iLoop, 1)) & "'", affUpdate
                
                rUpdate = rUpdate + affUpdate
                cmdEditMode.Caption = "UPDATE (" & rUpdate & ")"
            End If
        Next
    End With
    Call cmbLine_Click
End Sub

Private Sub cmdExcelLTPP_Click()
On Error GoTo errExcel
'    Set RsGet = Con.Execute("select a.*, upper(coalesce(c.cust_name, '***NO CUSTOMER***')) cust_name, coalesce(x.assy_no, '-') mark from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
'        & "left join v_ltpp_generate_all x on a.period = x.period and a.assy_no = x.assy_no and x.rev = a.rev " _
'        & "left join r_customer c on b.cust_id = c.cust_id where a.ltpp_doc = '" & txtDocNo & "' " & injLine & " order by cust_name, a.assy_no")
    Set RsGet = Con.Execute("select a.*, upper(coalesce(d.nm_line, '***NO LINE***')) nm_line, coalesce(x.assy_no, '-') mark, upper(coalesce(e.cust_name, '-')) cust_name, coalesce(b.item_muloq, 0) item_muloq from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
        & "left join v_ltpp_generate_all x on a.period = x.period and a.assy_no = x.assy_no and x.rev = a.rev " _
        & "left join mst_item_line c on b.item_id = c.item_id left join wip_mst_line d on c.cd_line_1 = d.cd_line left join r_customer e on b.cust_id = e.cust_id where a.ltpp_doc = '" & txtDocNo & "' " & injLine & " order by d.nm_line, a.assy_no")
    If Not RsGet.EOF Then
        Screen.MousePointer = vbHourglass
        Dim oExcel As Object 'Excel.Application 'Object
        Dim oBook As Object 'Excel.Workbook 'Object
        Dim oSheet As Object 'Excel.Worksheet 'Object
        Dim getSheet As Object
        Dim DataArray() As Variant
        Dim SumArray() As Variant
        Dim DataRev() As Variant
        Dim MarkRev() As Boolean
        Dim totalRow As Integer
        Dim totalRev As Integer
        Dim r As Integer
        Dim exNo As Integer
        
        Dim ColPos As Long
        Dim rowPos As Long
        
        With comSave
            .DefaultExt = ".xls"
            .Filter = "Excel Workbook (*.xls)|*.xls"
            .ShowSave
        End With
        
        Set RsDB = Con.Execute("select distinct rev, notes from ltpp_generate where period = '" & RsGet!period & "' and rev <= " & RsGet!rev & " order by rev")
        If Not RsDB.EOF Then
            totalRev = RsDB.RecordCount
            ReDim DataRev(1 To totalRev, 0 To 1) As Variant
            RsDB.MoveFirst
            iRev = 0
            Do Until RsDB.EOF
                iRev = iRev + 1
                DataRev(iRev, 0) = RsDB!rev
                DataRev(iRev, 1) = RsDB!Notes
                RsDB.MoveNext
            Loop
        Else
            ReDim DataRev(1 To 1, 0 To 1)
        End If
        RsDB.Close
        
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open pTemplateLTPP
        
        Set oBook = oExcel.Workbooks(1)
        Set oSheet = oBook.Worksheets(1)
        
        If cmbLine.Text = "-ALL-" Then
            totalRow = RsGet.RecordCount + cmbLine.ListCount - 1
        Else
            totalRow = RsGet.RecordCount + 1
        End If
        
        ReDim DataArray(1 To totalRow, 0 To 44) As Variant
        ReDim SumArray(0 To 39) As Variant
        r = 1
        exNo = 0
        holdLine = ""
        
        RsGet.MoveFirst
        Do Until RsGet.EOF
            exNo = exNo + 1
            If holdLine <> RsGet!nm_line Then
                holdLine = RsGet!nm_line
                DataArray(r, 1) = holdLine
                r = r + 1
            End If
            DataArray(r, 0) = exNo
            DataArray(r, 1) = RsGet!assy_no
            DataArray(r, 2) = RsGet!item_name
            DataArray(r, 3) = RsGet!cust_name
            DataArray(r, 4) = RsGet!item_muloq
            DataArray(r, 5) = RsGet!p1
            DataArray(r, 6) = RsGet!p2
            DataArray(r, 7) = RsGet!p3
            DataArray(r, 8) = RsGet!fg
            DataArray(r, 9) = RsGet!ost_mpp
            DataArray(r, 10) = RsGet!t_stock
            DataArray(r, 11) = RsGet!so_0
            DataArray(r, 12) = RsGet!bal_1
            DataArray(r, 13) = RsGet!ito_1
            DataArray(r, 14) = RsGet!so_1
            DataArray(r, 15) = RsGet!fc1
            DataArray(r, 16) = RsGet!del_rate_1
            DataArray(r, 17) = RsGet!s_stock_1
            DataArray(r, 18) = RsGet!need_1
            DataArray(r, 19) = RsGet!prod_plan_1
            DataArray(r, 20) = RsGet!bal_2
            DataArray(r, 21) = RsGet!ito_2
            DataArray(r, 22) = RsGet!so_2
            DataArray(r, 23) = RsGet!fc2
            DataArray(r, 24) = RsGet!del_rate_2
            DataArray(r, 25) = RsGet!s_stock_2
            DataArray(r, 26) = RsGet!need_2
            DataArray(r, 27) = RsGet!prod_plan_2
            DataArray(r, 28) = RsGet!bal_3
            DataArray(r, 29) = RsGet!ito_3
            DataArray(r, 30) = RsGet!so_3
            DataArray(r, 31) = RsGet!fc3
            DataArray(r, 32) = RsGet!del_rate_3
            DataArray(r, 33) = RsGet!s_stock_3
            DataArray(r, 34) = RsGet!need_3
            DataArray(r, 35) = RsGet!prod_plan_3
            DataArray(r, 36) = RsGet!bal_4
            DataArray(r, 37) = RsGet!ito_4
            DataArray(r, 38) = RsGet!so_4
            DataArray(r, 39) = RsGet!fc4
            DataArray(r, 40) = RsGet!del_rate_4
            DataArray(r, 41) = RsGet!s_stock_4
            DataArray(r, 42) = RsGet!need_4
            DataArray(r, 43) = RsGet!prod_plan_4
            DataArray(r, 44) = RsGet!bal_end
            
            SumArray(0) = SumArray(0) + RsGet!p1
            SumArray(1) = SumArray(1) + RsGet!p2
            SumArray(2) = SumArray(2) + RsGet!p3
            SumArray(3) = SumArray(3) + RsGet!fg
            SumArray(4) = SumArray(4) + RsGet!ost_mpp
            SumArray(5) = SumArray(5) + RsGet!t_stock
            SumArray(6) = SumArray(6) + RsGet!so_0
            SumArray(7) = SumArray(7) + RsGet!bal_1
            SumArray(8) = SumArray(8) + RsGet!ito_1
            SumArray(9) = SumArray(9) + RsGet!so_1
            SumArray(10) = SumArray(10) + RsGet!fc1
            SumArray(11) = SumArray(11) + RsGet!del_rate_1
            SumArray(12) = SumArray(12) + RsGet!s_stock_1
            SumArray(13) = SumArray(13) + RsGet!need_1
            SumArray(14) = SumArray(14) + RsGet!prod_plan_1
            SumArray(15) = SumArray(15) + RsGet!bal_2
            SumArray(16) = SumArray(16) + RsGet!ito_2
            SumArray(17) = SumArray(17) + RsGet!so_2
            SumArray(18) = SumArray(18) + RsGet!fc2
            SumArray(19) = SumArray(19) + RsGet!del_rate_2
            SumArray(20) = SumArray(20) + RsGet!s_stock_2
            SumArray(21) = SumArray(21) + RsGet!need_2
            SumArray(22) = SumArray(22) + RsGet!prod_plan_2
            SumArray(23) = SumArray(23) + RsGet!bal_3
            SumArray(24) = SumArray(24) + RsGet!ito_3
            SumArray(25) = SumArray(25) + RsGet!so_3
            SumArray(26) = SumArray(26) + RsGet!fc3
            SumArray(27) = SumArray(27) + RsGet!del_rate_3
            SumArray(28) = SumArray(28) + RsGet!s_stock_3
            SumArray(29) = SumArray(29) + RsGet!need_3
            SumArray(30) = SumArray(30) + RsGet!prod_plan_3
            SumArray(31) = SumArray(31) + RsGet!bal_4
            SumArray(32) = SumArray(32) + RsGet!ito_4
            SumArray(33) = SumArray(33) + RsGet!so_4
            SumArray(34) = SumArray(34) + RsGet!fc4
            SumArray(35) = SumArray(35) + RsGet!del_rate_4
            SumArray(36) = SumArray(36) + RsGet!s_stock_4
            SumArray(37) = SumArray(37) + RsGet!need_4
            SumArray(38) = SumArray(38) + RsGet!prod_plan_4
            SumArray(39) = SumArray(39) + RsGet!bal_end
            
            If RsGet!mark <> "-" And cmbRev.Text <> "0" Then
            '---------
                'ColPos = oSheet.Range("C" & r + 17).Left - 10
                'rowPos = oSheet.Range("C" & r + 17).Top + 10
                'MarkingRev oSheet, ColPos, rowPos
            '---------
                oSheet.Range("A" & r + 17).Resize(1, 54).Font.Bold = True
                oSheet.Range("A" & r + 17).Interior.Color = RGB(255, 255, 150)
            End If
            
            r = r + 1
            RsGet.MoveNext
        Loop
        
        oSheet.Shapes("rectDiketahui").TextFrame.Characters.Text = txtDiketahui
        oSheet.Shapes("rectDiperiksa").TextFrame.Characters.Text = txtDiperiksa
        oSheet.Shapes("rectDibuat").TextFrame.Characters.Text = txtDibuat
        
        oSheet.Range("A5") = "PERIODE : " & arrMM(1)
        'oSheet.Range("D3") = "REV. " & pRev
        oSheet.Range("F16") = arrMM(0)
        oSheet.Range("M16") = arrMM(1)
        oSheet.Range("U16") = arrMM(2)
        oSheet.Range("AC16") = arrMM(3)
        oSheet.Range("AK16") = arrMM(4)
        oSheet.Range("C8") = ": " & txtDocNo
        oSheet.Range("C9") = ": " & Format(dtLTPP, "DD MMMM YYYY")
        oSheet.Range("D8").Resize(totalRev, 2).Value = DataRev
        oSheet.Range("B19").Resize(totalRow, 1).NumberFormat = "@"
        oSheet.Range("A19").Resize(totalRow, 45).Value = DataArray
        oSheet.Range("A19").Resize(totalRow + 1, 45).VerticalAlignment = xlCenter
        oSheet.Range("F" & totalRow + 19).Resize(1, 40) = SumArray
        oSheet.Range("B" & totalRow + 19 & ":" & "D" & totalRow + 19).Merge
        oSheet.Range("B" & totalRow + 19).Value = "TOTAL"
        oSheet.Range("B" & totalRow + 19).HorizontalAlignment = xlCenter
        oSheet.Range("A19").Resize(totalRow + 1, 45).Borders.LineStyle = xlContinuous

        'oSheet.Columns("A:AW").EntireColumn.AutoFit
        oBook.SaveAs comSave.FileName
        oExcel.Quit
        Set oExcel = Nothing
        RsGet.Close
        
        MsgBox "Excel Saved...", vbInformation, "Exported..."
    Else
        MsgBox "No Data Found!", vbExclamation, "Exporting..."
    End If
    Screen.MousePointer = vbDefault
Exit Sub
errExcel:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Export: " & Err.Number
    End If
End Sub

Private Sub cmdFindAssy_Click()
On Error GoTo errFind
    txtFindAssy = ""
    popUp_LTPPFindAssy.Show 1
    If txtFindAssy <> "" Then
        With MSFlexGridLTPP
            For iLoop = 3 To .rows - 1
                If InStr(1, .TextMatrix(iLoop, 1), txtFindAssy) > 0 Then
                    .TopRow = iLoop
                    .LeftCol = 3
                    .Row = iLoop
                    .Col = 3
                    .SetFocus
                    Exit For
                End If
            Next
        End With
    End If
Exit Sub
errFind:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Find: " & Err.Number
    End If
End Sub

Private Function cekDataEng() As Boolean
    Dim qry As String
    Dim i As Byte
    Dim period As String
    Dim adaBaris As Boolean
    period = txtYear & Right("00" & cmbMM.ListIndex + 1, 2)
    
    qry = "select assy_no,item_name,fc1,fc2,fc3,fc4,so_qty,cap_p_day,case when bc_type='0' then item_muloq else item_perbox end itemmul, " _
    & "item_regdate from (select a.assy_no,fc1,fc2,fc3,fc4,so_qty from ltpp_fc a " _
    & " inner join ltpp_header b on a.ltpp_doc=b.ltpp_doc " _
    & " full outer join " _
    & " (select  a.ltpp_doc,a.assy_no ,so_qty from ltpp_so a inner join ltpp_header b on a.ltpp_doc=b.ltpp_doc " _
        & " where period='" & period & "' and rev=" & cmbRev & " and so_qty>0 " _
    & " ) v1 on a.ltpp_doc=v1.ltpp_doc and a.assy_no=v1.assy_no " _
    & " where period='" & period & "' and rev=" & cmbRev & " and (fc1>0 or fc2>0 or fc3>0 or fc4>0 or coalesce(so_qty,0)>0 ) " _
    & " group by a.assy_no,fc1,fc2,fc3,fc4,so_qty " _
    & " ) v1 left join (select partno,max(cap_p_day) cap_p_day from vlc_itemhd group by partno) a on v1.assy_no=a.partno " _
    & " left join mst_item v on v1.assy_no= v.item_id " _
    & " left join loadcap_mst_product_r r on v1.assy_no= r.partno " _
    & " where catgory='Injection' AND coalesce(cap_p_day, 0) <= 0 and stscode_id='01'" _
    & " order by assy_no asc"
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        adaBaris = True
    Else
        adaBaris = False
    End If
    grid0.rows = 1
    If adaBaris Then
        With grid0
            .rows = 1 + RsBantu.RecordCount
            i = 1
            .Col = .Cols - 1
            While Not RsBantu.EOF
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Trim$(RsBantu("assy_no"))
                .TextMatrix(i, 2) = Trim$(RsBantu("item_name"))
                .TextMatrix(i, 3) = IIf(IsNull(RsBantu("fc1")), "", FormatNumber(RsBantu("fc1"), 0))
                .TextMatrix(i, 4) = IIf(IsNull(RsBantu("fc2")), "", FormatNumber(RsBantu("fc2"), 0))
                .TextMatrix(i, 5) = IIf(IsNull(RsBantu("fc3")), "", FormatNumber(RsBantu("fc3"), 0))
                .TextMatrix(i, 6) = IIf(IsNull(RsBantu("fc4")), "", FormatNumber(RsBantu("fc4"), 0))
                .TextMatrix(i, 7) = IIf(IsNull(RsBantu("so_qty")), "", FormatNumber(RsBantu("so_qty"), 0))
                .TextMatrix(i, 8) = IIf(IsNull(RsBantu("cap_p_day")), "", FormatNumber(RsBantu("cap_p_day"), 0))
                .TextMatrix(i, 9) = IIf(IsNull(RsBantu("itemmul")), "", FormatNumber(RsBantu("itemmul"), 0))
                .TextMatrix(i, 10) = IIf(IsNull(RsBantu("item_regdate")), "", Format(RsBantu("item_regdate"), "dd-MMM-yyyy"))
                i = i + 1
                RsBantu.MoveNext
            Wend
        End With
        cekDataEng = True
    Else
        cekDataEng = False
    End If
    qry = "select assy_no,b.item_name,fc1,fc2,fc3,fc4,so_0,coalesce(cap_p_day,0) cap_p_day," _
    & "case when bc_type='0' then item_muloq else item_perbox end itemmul,item_regdate from v_ltpp_generate a inner join mst_item b on a.assy_no = b.item_id left join (select partno,max(cap_p_day) cap_p_day from vlc_itemhd " _
    & " group by partno ) v1 on a.assy_no=v1.partno " _
    & " where (fc1>0 or fc2>0 or fc3 >0 or fc4 >0 or so_0>0) and stscode_id='01' and ltpp_doc = '" & txtDocNo & "' and  case when bc_type='0' then item_muloq else item_perbox end =0 "
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        With grid0
            While Not RsBantu.EOF
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = .rows - 1
                .TextMatrix(.rows - 1, 1) = Trim$(RsBantu("assy_no"))
                .TextMatrix(.rows - 1, 2) = Trim$(RsBantu("item_name"))
                .TextMatrix(.rows - 1, 3) = IIf(IsNull(RsBantu("fc1")), "", FormatNumber(RsBantu("fc1"), 0))
                .TextMatrix(.rows - 1, 4) = IIf(IsNull(RsBantu("fc2")), "", FormatNumber(RsBantu("fc2"), 0))
                .TextMatrix(.rows - 1, 5) = IIf(IsNull(RsBantu("fc3")), "", FormatNumber(RsBantu("fc3"), 0))
                .TextMatrix(.rows - 1, 6) = IIf(IsNull(RsBantu("fc4")), "", FormatNumber(RsBantu("fc4"), 0))
                .TextMatrix(.rows - 1, 7) = IIf(IsNull(RsBantu("so_0")), "", FormatNumber(RsBantu("so_0"), 0))
                .TextMatrix(.rows - 1, 8) = IIf(IsNull(RsBantu("cap_p_day")), "", FormatNumber(RsBantu("cap_p_day"), 0))
                .TextMatrix(.rows - 1, 9) = IIf(IsNull(RsBantu("itemmul")), "", FormatNumber(RsBantu("itemmul"), 0))
                .TextMatrix(.rows - 1, 10) = IIf(IsNull(RsBantu("item_regdate")), "", Format(RsBantu("item_regdate"), "dd-MMM-yyyy"))
                RsBantu.MoveNext
            Wend
        End With
        cekDataEng = True
    End If
    
    Set RsBantu = Nothing
End Function

Private Sub cmdGenerate_Click()
'On Error GoTo errGenerate
    Dim i As Integer
    Dim iArr As Integer
    Dim setPeriod As String
    Dim setDoc As String
        
    cacheRowAssy = ""
    stSPart = 0
    cmdEditMode.Enabled = False
    cmdEditMode.Caption = "EDIT MODE"
    cmdEditMode.BackColor = vbButtonFace
    ltppEDIT = False
    setPeriod = txtYear & Format(cmbMM.ListIndex + 1, "00")
    pYear = txtYear
    prev = cmbRev.Text
    If cmbRev.Text <> "" Then
        Set RsGet = Con.Execute("select * from ltpp_header where period = '" & setPeriod & "' and rev = " & cmbRev.Text)
        If Not RsGet.EOF Then
            setDoc = RsGet!ltpp_doc
            For iArr = 0 To 4
                arrMM(iArr) = UCase(Format(DateSerial(Year(RsGet!date_period), Month(RsGet!date_period) + iArr - 1, 1), "MMMM YYYY"))
            Next
        Else
            setDoc = ""
        End If
        
        txtDocNo = setDoc
        
        RsGet.Close
    Else
        MsgBox "Input Rev!", vbExclamation, "Rev NULL"
        Exit Sub
    End If
    
    Call selectDB
    RsDB.Open "select * from ltpp_generate where ltpp_doc = '" & setDoc & "'", Con, adOpenDynamic, adLockOptimistic
    If RsDB.EOF Then ' jika kosong
        If cekDataEng Then
            pic0.Visible = True
            grid0.SetFocus
            Exit Sub
        Else
            pic0.Visible = False
        End If
        
        'from upload
'        Con.Execute "insert into ltpp_generate select *, '" & txtNote & "' notes, now()::date ltpp_date from v_ltpp_generate_all where ltpp_doc = '" & setDoc & "'"
        Set RsBantu = Con.Execute("select a.*, b.st_sparepart, b.prc_safetystock, coalesce(b.item_muloq, 0) item_muloq, now()::date ltpp_date,coalesce(cap_p_day,1) cap_p_day,item_perbox,coalesce(bc_type,'0') bc_type " _
        & " from v_ltpp_generate a inner join mst_item b on a.assy_no = b.item_id left join (select partno,max(cap_p_day) cap_p_day from vlc_itemhd " _
        & " group by partno ) v1 on a.assy_no=v1.partno " _
        & " where ltpp_doc = '" & setDoc & "'")
        
        Do Until RsBantu.EOF ' lakukan sampai akhir
            With RsDB
                .AddNew
                !ltpp_doc = RsBantu!ltpp_doc
                !period = RsBantu!period
                !date_period = RsBantu!date_period
                !lt = RsBantu!lt
                !hkw_1 = RsBantu!hkw_1
                !hkw_2 = RsBantu!hkw_2
                !hkw_3 = RsBantu!hkw_3
                !hkw_4 = RsBantu!hkw_4
                !hkw_5 = RsBantu!hkw_5
                !fc_m4 = RsBantu!fc_m4
                !rev = RsBantu!rev
                !assy_no = RsBantu!assy_no
                !item_name = RsBantu!item_name
                !cct = RsBantu!cct
                !pp = RsBantu!pp
                !p1 = RsBantu!p1
                !p2 = RsBantu!p2
                !p3 = RsBantu!p3
                !p4 = RsBantu!p4
                !p5 = RsBantu!p5
                !p6 = RsBantu!p6
                !p7 = RsBantu!p7
                !p8 = RsBantu!p8
                !p9 = RsBantu!p9
                !p10 = RsBantu!p10
                !rep_qa = RsBantu!rep_qa
                !hav = RsBantu!hav
                !fg = RsBantu!fg
                !ost_mpp = RsBantu!ost_mpp
                !t_stock = RsBantu!t_stock
                !fc1 = RsBantu!fc1
                !fc2 = RsBantu!fc2
                !fc3 = RsBantu!fc3
                !fc4 = RsBantu!fc4
                !t_fc = RsBantu!t_fc
                !so_0 = RsBantu!so_0
                !so_1 = RsBantu!so_1
                !so_2 = RsBantu!so_2
                !so_3 = RsBantu!so_3
                !so_4 = RsBantu!so_4
                                
                c_cap_p_day = FormatNumber(RsBantu("cap_p_day"), 0)
                
                
                del_rate_1 = RsBantu!fc1 / RsBantu!hkw_1
                del_rate_2 = RsBantu!fc2 / RsBantu!hkw_2
                del_rate_3 = RsBantu!fc3 / RsBantu!hkw_3
                del_rate_4 = RsBantu!fc4 / RsBantu!hkw_4
                
                s_stock_1 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc1 / 100)
                s_stock_2 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc2 / 100)
                s_stock_3 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc3 / 100)
                s_stock_4 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc4 / 100)
                
                
                'M+0
                bal_1 = RsBantu!t_stock - RsBantu!so_0
                If del_rate_1 > 0 Then
                    ito_1 = bal_1 / del_rate_1
                Else
                    ito_1 = 0
                End If
                If RsBantu!fc1 > RsBantu!so_1 Then
                    need_1 = RsBantu!fc1 + s_stock_1 - bal_1
                Else
                    need_1 = RsBantu!so_1 + s_stock_1 - bal_1
                End If
                
                prod_plan_1 = 0
                                
                If RsBantu!item_muloq > 0 Then
                    prod_plan_1 = RoundNumber(-Int(-need_1 / RsBantu!item_muloq) * RsBantu!item_muloq)
                ElseIf RsBantu!item_perbox > 0 Then
                    prod_plan_1 = RoundNumber(-Int(-need_1 / RsBantu!item_perbox) * RsBantu!item_perbox)
                End If
                
                If need_1 < 0 Then
                    need_1 = 0
                    prod_plan_1 = 0
                End If
                                                                           
                If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_1 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_1), "a")
                    prod_plan_1 = FKelipatan(RsBantu("item_muloq"), t_pp, "a") '  t_Mult * RsBantu("item_muloq")
                ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_1 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_1), "a")
                    prod_plan_1 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
                End If
                
                'M+1
                If RsBantu!fc1 > RsBantu!so_1 Then
                    bal_2 = bal_1 + prod_plan_1 - RsBantu!fc1
                Else
                    bal_2 = bal_1 + prod_plan_1 - RsBantu!so_1
                End If
                If del_rate_2 > 0 Then
                    ito_2 = bal_2 / del_rate_2
                Else
                    ito_2 = 0
                End If
                If RsBantu!fc2 > RsBantu!so_2 Then
                    need_2 = RsBantu!fc2 + s_stock_2 - bal_2
                Else
                    need_2 = RsBantu!so_2 + s_stock_2 - bal_2
                End If
                
                prod_plan_2 = 0
                If RsBantu!item_muloq > 0 Then
                    prod_plan_2 = RoundNumber(-Int(-need_2 / RsBantu!item_muloq) * RsBantu!item_muloq)
                ElseIf RsBantu!item_perbox > 0 Then
                    prod_plan_2 = RoundNumber(-Int(-need_2 / RsBantu!item_perbox) * RsBantu!item_perbox)
                End If
                
                If need_2 < 0 Then
                    need_2 = 0
                    prod_plan_2 = 0
                End If
                
                If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_2 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_2), "a")  't_Mult * c_cap_p_day
                    prod_plan_2 = FKelipatan(RsBantu("item_muloq"), t_pp, "a") 't_Mult * RsBantu("item_muloq")
                ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_2 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_2), "a")
                    prod_plan_2 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
                End If
                
                'M+2
                If RsBantu!fc2 > RsBantu!so_2 Then
                    bal_3 = bal_2 + prod_plan_2 - RsBantu!fc2
                Else
                    bal_3 = bal_2 + prod_plan_2 - RsBantu!so_2
                End If
                If del_rate_3 > 0 Then
                    ito_3 = bal_3 / del_rate_3
                Else
                    ito_3 = 0
                End If
                If RsBantu!fc3 > RsBantu!so_3 Then
                    need_3 = RsBantu!fc3 + s_stock_3 - bal_3
                Else
                    need_3 = RsBantu!so_3 + s_stock_3 - bal_3
                End If
                
                prod_plan_3 = 0
                If RsBantu!item_muloq > 0 Then
                    prod_plan_3 = RoundNumber(-Int(-need_3 / RsBantu!item_muloq) * RsBantu!item_muloq)
                ElseIf RsBantu!item_perbox > 0 Then
                    prod_plan_3 = RoundNumber(-Int(-need_3 / RsBantu!item_perbox) * RsBantu!item_perbox)
                End If
                    
                If need_3 < 0 Then
                    need_3 = 0
                    prod_plan_3 = 0
                End If
                
                If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_3 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_3), "a")
                    prod_plan_3 = FKelipatan(RsBantu("item_muloq"), t_pp, "a")
                ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_3 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_3), "a")
                    prod_plan_3 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
                End If
                
                'M+3
                If RsBantu!fc3 > RsBantu!so_3 Then
                    bal_4 = bal_3 + prod_plan_3 - RsBantu!fc3
                Else
                    bal_4 = bal_3 + prod_plan_3 - RsBantu!so_3
                End If
                If del_rate_4 > 0 Then
                    ito_4 = bal_4 / del_rate_4
                Else
                    ito_4 = 0
                End If
                If RsBantu!fc4 > RsBantu!so_4 Then
                    need_4 = RsBantu!fc4 + s_stock_4 - bal_4
                Else
                    need_4 = RsBantu!so_4 + s_stock_4 - bal_4
                End If
                
                prod_plan_4 = 0
                If RsBantu!item_muloq > 0 Then
                    prod_plan_4 = RoundNumber(-Int(-need_4 / RsBantu!item_muloq) * RsBantu!item_muloq)
                ElseIf RsBantu!item_perbox > 0 Then
                    prod_plan_4 = RoundNumber(-Int(-need_4 / RsBantu!item_perbox) * RsBantu!item_perbox)
                End If
                
                If need_4 < 0 Then
                    need_4 = 0
                    prod_plan_4 = 0
                End If
                
                If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_4 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_4), "a")
                    prod_plan_4 = FKelipatan(RsBantu("item_muloq"), t_pp, "a")
                ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_4 > 0 Then
                    t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_4), "a")
                    prod_plan_4 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
                End If
                
                'bal end
                If RsBantu!fc4 > RsBantu!so_4 Then
                    bal_end = bal_4 + prod_plan_4 - RsBantu!fc4
                Else
                    bal_end = bal_4 + prod_plan_4 - RsBantu!so_4
                End If
                
                !bal_1 = bal_1
                !ito_1 = ito_1
                !del_rate_1 = del_rate_1
                !del_rate_2 = del_rate_2
                !del_rate_3 = del_rate_3
                !del_rate_4 = del_rate_4
                !s_stock_1 = s_stock_1
                !s_stock_2 = s_stock_2
                !s_stock_3 = s_stock_3
                !s_stock_4 = s_stock_4
                !need_1 = need_1
                !prod_plan_1 = prod_plan_1
                !bal_2 = bal_2
                !ito_2 = ito_2
                !need_2 = need_2
                !prod_plan_2 = prod_plan_2
                !bal_3 = bal_3
                !ito_3 = ito_3
                !need_3 = need_3
                !prod_plan_3 = prod_plan_3
                !bal_4 = bal_4
                !ito_4 = ito_4
                !need_4 = need_4
                !prod_plan_4 = prod_plan_4
                !bal_end = bal_end
                !Notes = txtNote
                !ltpp_date = RsBantu!ltpp_date
                .Update
            End With
            RsBantu.MoveNext
        Loop
        RsBantu.Close
        
        'from last rev
        Con.Execute "insert into ltpp_generate select '" & setDoc & "' ltpp_doc, a.period, a.date_period, a.lt, a.hkw_1, a.hkw_2, a.hkw_3, a.hkw_4, a.hkw_5, a.fc_m4, " & cmbRev.Text & ", a.assy_no, a.item_name, a.cct, " _
            & "a.pp, a.p1, a.p2, a.p3, a.p4, a.p5, a.p6, a.p7, a.p8, a.p9, a.p10, a.rep_qa, a.hav, a.fg, a.ost_mpp, a.t_stock, a.fc1, a.fc2, a.fc3, " _
            & "a.fc4, a.t_fc, a.so_0, a.so_1, a.so_2, a.so_3, a.so_4, a.bal_1, a.ito_1, a.del_rate_1, a.del_rate_2, a.del_rate_3, a.del_rate_4, " _
            & "a.s_stock_1, a.s_stock_2, a.s_stock_3, a.s_stock_4, a.need_1, a.prod_plan_1, a.bal_2, a.ito_2, a.need_2, a.prod_plan_2, a.bal_3, a.ito_3, " _
            & "a.need_3, a.prod_plan_3, a.bal_4, a.ito_4, a.need_4, a.prod_plan_4, a.bal_end, '" & txtNote & "' notes, now()::date ltpp_date " _
            & "from ltpp_generate a left join v_ltpp_generate b on a.assy_no = b.assy_no and a.period = b.period and b.rev = " & cmbRev.Text & " " _
            & "where a.period = '" & setPeriod & "' and b.assy_no isnull and a.rev = " & cmbRev.Text & " - 1"
        InsertAssy setDoc
    Else
        If chkReGenerate.Value = 1 Then
            RsDB.MoveFirst
            If cekDataEng Then
                pic0.Visible = True
                grid0.SetFocus
                Exit Sub
            Else
                pic0.Visible = False
            End If
            Do Until RsDB.EOF
                reGenerateLTPP setDoc, RTrim(RsDB!assy_no)
                RsDB.MoveNext
            Loop
            InsertAssy setDoc
        End If
    End If
    RsDB.Close
    
    
    
    Set RsDB = Con.Execute("select distinct upper(coalesce(d.nm_line, '***NO LINE***')) nm_line from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
        & "left join mst_item_line c on b.item_id = c.item_id " _
        & "left join wip_mst_line d on c.cd_line_1 = d.cd_line " _
        & "where a.ltpp_doc = '" & setDoc & "'")
    cmbLine.Clear
    cmbLine.AddItem "-ALL-"
    Do Until RsDB.EOF
        cmbLine.AddItem RsDB!nm_line
        RsDB.MoveNext
    Loop
    RsDB.Close
    cmbLine.Text = "-ALL-"
'Exit Sub
'errGenerate:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description, vbCritical, "Error Generate: " & Err.Number
'    End If
End Sub

Private Sub InsertAssy(dokumenNo As String)
    Dim rsTempAssy As ADODB.Recordset
    Set rsTempAssy = Con.Execute("select assy_no from ltpp_wip where ltpp_doc='" & dokumenNo & "'")
    'WIP
    While Not rsTempAssy.EOF
        Set rsAssy = Con.Execute("select bom_par_item,bom_com_item,item_name from mst_bom a inner join mst_item b on a.bom_com_item=b.item_id " _
            & " where pfm_id='10' and bom_par_item='" & Trim$(rsTempAssy("assy_no")) & "'")
        If rsAssy.RecordCount > 0 Then
            While Not rsAssy.EOF
                If checkPrimary(dokumenNo, rsAssy("bom_com_item"), "ltpp_wip") = 0 Then
                    Con.Execute "insert into ltpp_wip values(" _
                    & "'" & dokumenNo & "','" & rsAssy("bom_com_item") & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)"
                    
                End If
                rsAssy.MoveNext
'                MsgBox "lanjut"
            Wend
        End If
        rsTempAssy.MoveNext
    Wend
    'SO
    Set rsTempAssy = Con.Execute("select * from ltpp_so where ltpp_doc='" & dokumenNo & "'")
    While Not rsTempAssy.EOF
        Set rsAssy = Con.Execute("select bom_par_item,bom_com_item,item_name from mst_bom a inner join mst_item b on a.bom_com_item=b.item_id " _
            & " where pfm_id='10' and bom_par_item='" & Trim$(rsTempAssy("assy_no")) & "'")
        If rsAssy.RecordCount > 0 Then
            While Not rsAssy.EOF
'                MsgBox checkPrimary(dokumenNo, rsAssy("bom_com_item"), "ltpp_so")
                If checkPrimary(dokumenNo, rsAssy("bom_com_item"), "ltpp_so") = 0 Then
                    Con.Execute "insert into ltpp_so values(" _
                    & "'" & dokumenNo & "','" & rsTempAssy("so_id") & "',NULL" _
                    & ",'" & rsAssy("bom_com_item") & "',0,'" & rsTempAssy("so_reqdate") & "')"
                End If
                rsAssy.MoveNext
            Wend
        End If
        rsTempAssy.MoveNext
    Wend
    'FC
    Set rsTempAssy = Con.Execute("select * from ltpp_fc where ltpp_doc='" & dokumenNo & "'")
    While Not rsTempAssy.EOF
        Set rsAssy = Con.Execute("select bom_par_item,bom_com_item,item_name from mst_bom a inner join mst_item b on a.bom_com_item=b.item_id " _
            & " where pfm_id='10' and bom_par_item='" & Trim$(rsTempAssy("assy_no")) & "'")
        If rsAssy.RecordCount > 0 Then
            While Not rsAssy.EOF
                If checkPrimary(dokumenNo, rsAssy("bom_com_item"), "ltpp_fc") = 0 Then
                    Con.Execute "insert into ltpp_fc values(" _
                    & "'" & dokumenNo & "','" & rsAssy("bom_com_item") & "',0,0,0,0)"
                End If
                rsAssy.MoveNext
            Wend
        End If
        rsTempAssy.MoveNext
    Wend
    'Generate
    Set rsTempAssy = Con.Execute("select * from ltpp_generate where ltpp_doc='" & dokumenNo & "'")
    While Not rsTempAssy.EOF
        Set rsAssy = Con.Execute("select bom_par_item,bom_com_item,item_name from mst_bom a inner join mst_item b on a.bom_com_item=b.item_id " _
            & " where pfm_id='10' and bom_par_item='" & Trim$(rsTempAssy("assy_no")) & "'")
        If rsAssy.RecordCount > 0 Then
            While Not rsAssy.EOF
                If checkPrimary(dokumenNo, rsAssy("bom_com_item"), "ltpp_generate") = 0 Then
                    Con.Execute "insert into ltpp_generate(ltpp_doc,period,date_period,lt,hkw_1,hkw_2,hkw_3,hkw_4,hkw_5,fc_m4,rev,assy_no,item_name,fg,fc1,fc2,fc3,fc4," _
                    & "prod_plan_1,prod_plan_2,prod_plan_3,prod_plan_4,p1,p2,p3) values(" _
                    & "'" & dokumenNo & "','" & rsTempAssy("period") & "'," _
                    & "'" & rsTempAssy("date_period") & "'," & rsTempAssy("lt") & "," _
                    & rsTempAssy("hkw_1") & "," & rsTempAssy("hkw_2") & "," & rsTempAssy("hkw_3") & "," _
                    & rsTempAssy("hkw_4") & "," & rsTempAssy("hkw_5") & ",0," & rsTempAssy("rev") & "" _
                    & ",'" & rsAssy("bom_com_item") & "','" & rsAssy("item_name") & "',0," & rsTempAssy("fc1") & "," & rsTempAssy("fc2") & "," _
                    & rsTempAssy("fc3") & "," & rsTempAssy("fc4") & "," & rsTempAssy("prod_plan_1") & "," & rsTempAssy("prod_plan_2") & "" _
                    & "," & rsTempAssy("prod_plan_3") & "," & rsTempAssy("prod_plan_4") & "," _
                    & rsTempAssy("p1") & "," & rsTempAssy("p2") & "," & rsTempAssy("p3") & ")"
                Else
                    Con.Execute "update ltpp_generate set prod_plan_1=" & rsTempAssy("prod_plan_1") & ",prod_plan_2=" & rsTempAssy("prod_plan_2") & ",prod_plan_3=" & rsTempAssy("prod_plan_3") & ",prod_plan_4 =" & rsTempAssy("prod_plan_4") _
                    & ",p1=" & rsTempAssy("p1") & ",p2=" & rsTempAssy("p2") & ",p3=" & rsTempAssy("p3") _
                    & " where ltpp_doc='" & dokumenNo & "' and assy_no='" & rsAssy("bom_com_item") & "'"
                End If
                rsAssy.MoveNext
            Wend
        End If
        rsTempAssy.MoveNext
    Wend
    Set rsTempAssy = Nothing
End Sub

Private Function checkPrimary(dokumen As String, part As String, tabel As String) As Byte
    Set rsIter = Con.Execute("select count(*) from " & tabel & " where ltpp_doc='" & dokumen & "' and assy_no='" & part & "'")
    checkPrimary = rsIter(0)
    Set rsIter = Nothing
End Function

Private Sub reGenerateLTPP(ByVal DocNo As String, ByVal AssyNo As String)
    Set RsBantu = Con.Execute("select a.*, b.st_sparepart, b.prc_safetystock, coalesce(b.item_muloq, 0) item_muloq, now()::date ltpp_date " _
    & " ,coalesce(cap_p_day,1) cap_p_day,item_perbox,coalesce(bc_type,'0') bc_type from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
    & " left join (select partno,max(cap_p_day) cap_p_day from vlc_itemhd " _
    & " group by partno ) v1 on a.assy_no=v1.partno " _
    & " where ltpp_doc = '" & DocNo & "' and assy_no = '" & AssyNo & "'")
    If Not RsBantu.EOF Then
        c_cap_p_day = FormatNumber(RsBantu("cap_p_day"), 0)
        
        del_rate_1 = RsBantu!fc1 / RsBantu!hkw_1
        del_rate_2 = RsBantu!fc2 / RsBantu!hkw_2
        del_rate_3 = RsBantu!fc3 / RsBantu!hkw_3
        del_rate_4 = RsBantu!fc4 / RsBantu!hkw_4
        
        s_stock_1 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc1 / 100)
        s_stock_2 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc2 / 100)
        s_stock_3 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc3 / 100)
        s_stock_4 = RoundNumber(RsBantu!prc_safetystock * RsBantu!fc4 / 100)
        
        'M+0
        bal_1 = RsBantu!t_stock - RsBantu!so_0
        If del_rate_1 > 0 Then
            ito_1 = bal_1 / del_rate_1
        Else
            ito_1 = 0
        End If
        If RsBantu!fc1 > RsBantu!so_1 Then
            need_1 = RsBantu!fc1 + s_stock_1 - bal_1
        Else
            need_1 = RsBantu!so_1 + s_stock_1 - bal_1
        End If
        debugitem = AssyNo
        prod_plan_1 = 0
        If RsBantu!item_muloq > 0 Then
            prod_plan_1 = RoundNumber(-Int(-need_1 / RsBantu!item_muloq) * RsBantu!item_muloq)
        ElseIf RsBantu!item_perbox > 0 Then
            prod_plan_1 = RoundNumber(-Int(-need_1 / RsBantu!item_perbox) * RsBantu!item_perbox)
        End If
        If debugitem = "11101-B630-002" Then
            MsgBox need_1 & vbNewLine & "fc1=" & RsBantu!fc1
        End If
        If need_1 < 0 Then
            need_1 = 0
            prod_plan_1 = 0
        End If
        
        If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_1 > 0 Then
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_1), "a") 't_Mult * c_cap_p_day
            prod_plan_1 = FKelipatan(RsBantu("item_muloq"), t_pp, "a") '  t_Mult * RsBantu("item_muloq")
        ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_1 > 0 Then
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_1), "a")
            prod_plan_1 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
        End If
        
        'M+1
        If RsBantu!fc1 > RsBantu!so_1 Then
            bal_2 = bal_1 + prod_plan_1 - RsBantu!fc1
        Else
            bal_2 = bal_1 + prod_plan_1 - RsBantu!so_1
        End If
        If del_rate_2 > 0 Then
            ito_2 = bal_2 / del_rate_2
        Else
            ito_2 = 0
        End If
        If RsBantu!fc2 > RsBantu!so_2 Then
            need_2 = RsBantu!fc2 + s_stock_2 - bal_2
        Else
            need_2 = RsBantu!so_2 + s_stock_2 - bal_2
        End If
        
        prod_plan_2 = 0
        If RsBantu!item_muloq > 0 Then
            prod_plan_2 = RoundNumber(-Int(-need_2 / RsBantu!item_muloq) * RsBantu!item_muloq)
        ElseIf RsBantu!item_perbox > 0 Then
            prod_plan_2 = RoundNumber(-Int(-need_2 / RsBantu!item_perbox) * RsBantu!item_perbox)
        End If
        
        If need_2 < 0 Then
            need_2 = 0
            prod_plan_2 = 0
        End If
        
        If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_2 > 0 Then
            't_Mult = (need_2 / c_cap_p_day) + 1
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_2), "a")  't_Mult * c_cap_p_day
            't_Mult = (t_pp / RsBantu("item_muloq"))
            prod_plan_2 = FKelipatan(RsBantu("item_muloq"), t_pp, "a") 't_Mult * RsBantu("item_muloq")
        ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_2 > 0 Then
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_2), "a")
            prod_plan_2 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
        End If

        'M+2
        If RsBantu!fc2 > RsBantu!so_2 Then
            bal_3 = bal_2 + prod_plan_2 - RsBantu!fc2
        Else
            bal_3 = bal_2 + prod_plan_2 - RsBantu!so_2
        End If
        If del_rate_3 > 0 Then
            ito_3 = bal_3 / del_rate_3
        Else
            ito_3 = 0
        End If
        If RsBantu!fc3 > RsBantu!so_3 Then
            need_3 = RsBantu!fc3 + s_stock_3 - bal_3
        Else
            need_3 = RsBantu!so_3 + s_stock_3 - bal_3
        End If
        
        prod_plan_3 = 0
        If RsBantu!item_muloq > 0 Then
            prod_plan_3 = RoundNumber(-Int(-need_3 / RsBantu!item_muloq) * RsBantu!item_muloq)
        ElseIf RsBantu!item_perbox > 0 Then
            prod_plan_3 = RoundNumber(-Int(-need_3 / RsBantu!item_perbox) * RsBantu!item_perbox)
        End If
        
        If need_3 < 0 Then
            need_3 = 0
            prod_plan_3 = 0
        End If
        
        If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_3 > 0 Then
           ' t_Mult = (need_3 / c_cap_p_day) + 0
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_3), "a")  ' t_Mult * c_cap_p_day
            't_Mult = (t_pp / RsBantu("item_muloq"))
            prod_plan_3 = FKelipatan(RsBantu("item_muloq"), t_pp, "a") 't_Mult * RsBantu("item_muloq")
        ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_3 > 0 Then
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_3), "a")
            prod_plan_3 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
        End If
        
        'M+3
        If RsBantu!fc3 > RsBantu!so_3 Then
            bal_4 = bal_3 + prod_plan_3 - RsBantu!fc3
        Else
            bal_4 = bal_3 + prod_plan_3 - RsBantu!so_3
        End If
        If del_rate_4 > 0 Then
            ito_4 = bal_4 / del_rate_4
        Else
            ito_4 = 0
        End If
        If RsBantu!fc4 > RsBantu!so_4 Then
            need_4 = RsBantu!fc4 + s_stock_4 - bal_4
        Else
            need_4 = RsBantu!so_4 + s_stock_4 - bal_4
        End If
        
        prod_plan_4 = 0
        If RsBantu!item_muloq > 0 Then
            prod_plan_4 = RoundNumber(-Int(-need_4 / RsBantu!item_muloq) * RsBantu!item_muloq)
        ElseIf RsBantu!item_perbox > 0 Then
            prod_plan_4 = RoundNumber(-Int(-need_4 / RsBantu!item_perbox) * RsBantu!item_perbox)
        End If
            
        If need_4 < 0 Then
            need_4 = 0
            prod_plan_4 = 0
        End If
        
        If RsBantu("item_muloq") > 0 And c_cap_p_day > 0 And need_4 > 0 Then
            't_Mult = (need_4 / c_cap_p_day) + 1
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_4), "a") 't_Mult * c_cap_p_day
            't_Mult = (t_pp / RsBantu("item_muloq"))
            prod_plan_4 = FKelipatan(RsBantu("item_muloq"), t_pp, "a") 't_Mult * RsBantu("item_muloq")
        ElseIf RsBantu("item_perbox") > 0 And c_cap_p_day > 0 And need_4 > 0 Then
            t_pp = FKelipatan(CDbl(c_cap_p_day), CDbl(need_4), "a")
            prod_plan_4 = FKelipatan(RsBantu("item_perbox"), t_pp, "a")
        End If
        
        'bal end
        If RsBantu!fc4 > RsBantu!so_4 Then
            bal_end = bal_4 + prod_plan_4 - RsBantu!fc4
        Else
            bal_end = bal_4 + prod_plan_4 - RsBantu!so_4
        End If
        
        Con.Execute "update ltpp_generate set " _
        & "bal_1 = " & bal_1 & ", ito_1 = " & ito_1 & ", del_rate_1 = " & del_rate_1 & ", del_rate_2 = " & del_rate_2 _
        & ", del_rate_3 = " & del_rate_3 & ", del_rate_4 = " & del_rate_4 & ", s_stock_1 = " & s_stock_1 _
        & ", s_stock_2 = " & s_stock_2 & ", s_stock_3 = " & s_stock_3 & ", s_stock_4 = " & s_stock_4 & ", need_1 = " & need_1 _
        & ", prod_plan_1 = " & prod_plan_1 & ", bal_2 = " & bal_2 & ", ito_2 = " & ito_2 & ", need_2 = " & need_2 _
        & ", prod_plan_2 = " & prod_plan_2 & ", bal_3 = " & bal_3 & ", ito_3 = " & ito_3 & ", need_3 = " & need_3 _
        & ", prod_plan_3 = " & prod_plan_3 & ", bal_4 = " & bal_4 & ", ito_4 = " & ito_4 & ", need_4 = " & need_4 _
        & ", prod_plan_4 = " & prod_plan_4 & ", bal_end = " & bal_end & ", notes = '" & txtNote & "',ltpp_date = '" & RsBantu!ltpp_date & "' " _
        & " where ltpp_doc = '" & DocNo & "' and assy_no = '" & AssyNo & "'"
    End If
    RsBantu.Close
End Sub

Private Sub cmdPrintLTPP_Click()
On Error GoTo errPrint
    pStPrinter = False
    PopUp_PrinterPrint.Show 1
    If pStPrinter = True Then
'        Set RsGet = Con.Execute("select a.*, upper(coalesce(c.cust_name, '***NO CUSTOMER***')) cust_name, coalesce(x.assy_no, '-') mark from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
'            & "left join v_ltpp_generate_all x on a.period = x.period and a.assy_no = x.assy_no and x.rev = a.rev " _
'            & "left join r_customer c on b.cust_id = c.cust_id where a.ltpp_doc = '" & txtDocNo & "' " & injLine & " order by cust_name, a.assy_no")
    Set RsGet = Con.Execute("select a.*, upper(coalesce(d.nm_line, '***NO LINE***')) nm_line, coalesce(x.assy_no, '-') mark, upper(coalesce(e.cust_name, '-')) cust_name, coalesce(b.item_muloq, 0) item_muloq from ltpp_generate a inner join mst_item b on a.assy_no = b.item_id " _
        & "left join v_ltpp_generate_all x on a.period = x.period and a.assy_no = x.assy_no and x.rev = a.rev " _
        & "left join mst_item_line c on b.item_id = c.item_id left join wip_mst_line d on c.cd_line_1 = d.cd_line left join r_customer e on b.cust_id = e.cust_id where a.ltpp_doc = '" & txtDocNo & "' " & injLine & " order by d.nm_line, a.assy_no")
        If Not RsGet.EOF Then
            Printer.ScaleMode = vbTwips
            If optA3.Value = True Then
                Printer.PaperSize = vbPRPSA3
            Else
                Printer.PaperSize = vbPRPSA4
            End If
            Printer.Orientation = vbPRORLandscape
            Printer.Font = "Arial"
            printLTPP Printer, Printer.Width, Printer.Height
            Printer.EndDoc
        Else
            MsgBox "No Data Found!", vbExclamation, "Warning..."
        End If
    End If
Exit Sub
errPrint:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Print: " & Err.Number
    End If
End Sub

Private Sub cmdSApproval_Click()
    SaveINI "LTPP", "diketahui", txtDiketahui
    SaveINI "LTPP", "diperiksa", txtDiperiksa
    SaveINI "LTPP", "dibuat", txtDibuat
    MsgBox "Approval Saved!", vbInformation, "Setting Approval"
End Sub

Private Sub cmdUpload_Click()
On Error GoTo errUpload
    comDialogUpload.ShowOpen
    Me.MousePointer = vbHourglass
    cmdGenerate.Enabled = False
    cmdPrintLTPP.Enabled = False
    cmdExcelLTPP.Enabled = False
    cmdCancelUpload = False
    Call uploadExcel(comDialogUpload.FileName)
    cmdGenerate.Enabled = True
    cmdPrintLTPP.Enabled = True
    cmdExcelLTPP.Enabled = True
    Call cmbMM_Click
    Me.MousePointer = vbDefault
Exit Sub
errUpload:
    Me.MousePointer = vbDefault
    uploadStatus = ""
    cmdUpload.Enabled = True
    cmdGenerate.Enabled = True
    cmdPrintLTPP.Enabled = True
    cmdExcelLTPP.Enabled = True
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Upload: " & Err.Number
    End If
End Sub

Private Sub Form_Activate()
FocusTab Me
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call LoadHeaderGrid(1, "M+0", "M+1", "M+2", "M+3", "M+4")
    Call getListComboBox
    Call getSA
'    Call WheelHook(Me.hwnd)
    txtYear = Year(Now)
    settingGrid
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub LoadHeaderGrid(ByVal rows As Integer, month0 As String, month1 As String, month2 As String, month3 As String, month4 As String)
    rows = rows + 3
    With MSFlexGridLTPP
        .Clear
        .Cols = 58
        .rows = rows
        .FixedRows = 3
        .FixedCols = 4
        
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = False
        
        .RowHeightMin = 300
        .Col = 0
        .Row = 0
        .ColWidth(0) = 700
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignRightCenter
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .TextMatrix(0, 0) = "NO"
        .TextMatrix(1, 0) = "NO"
        .TextMatrix(2, 0) = "NO"
         
        .Col = 1
        .Row = 0
        .ColWidth(1) = 2500
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .MergeCol(1) = True
        .TextMatrix(0, 1) = "ASSY NO"
        .TextMatrix(1, 1) = "ASSY NO"
        .TextMatrix(2, 1) = "ASSY NO"
        
        .Col = 2
        .Row = 0
        .ColWidth(2) = 2500
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .MergeCol(2) = True
        .TextMatrix(0, 2) = "ASSY NAME"
        .TextMatrix(1, 2) = "ASSY NAME"
        .TextMatrix(2, 2) = "ASSY NAME"
        
        .Col = 3
        .Row = 0
        .ColWidth(3) = 3000
        .CellAlignment = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .MergeCol(3) = True
        .TextMatrix(0, 3) = "CUSTOMER"
        .TextMatrix(1, 3) = "CUSTOMER"
        .TextMatrix(2, 3) = "CUSTOMER"
         
        .Col = 4
        .Row = 0
        .ColWidth(4) = 700
        .CellAlignment = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .MergeCol(4) = True
        .TextMatrix(0, 4) = "MPQ"
        .TextMatrix(1, 4) = "MPQ"
        .TextMatrix(2, 4) = "MPQ"
        
        .Col = 5
        .Row = 0
        .Text = month0
        .TextMatrix(1, 5) = "WIP"
        .ColWidth(5) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(2, 5) = "PP"
         
        .Col = 6
        .Row = 0
        .Text = month0
        .TextMatrix(1, 6) = "WIP"
        .ColWidth(6) = 800
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(2, 6) = "INJ" 'P1
        
        .Col = 7
        .Row = 0
        .Text = month0
        .TextMatrix(1, 7) = "WIP"
        .ColWidth(7) = 1000
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(2, 7) = "NON ASSY" 'P2

        .Col = 8
        .Row = 0
        .Text = month0
        .TextMatrix(1, 8) = "WIP"
        .ColWidth(8) = 800
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(2, 8) = "ASSY" 'P3
        
        .Col = 9
        .Row = 0
        .Text = month0
        .TextMatrix(1, 9) = "WIP"
        .ColWidth(9) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(2, 9) = "P4"
        
        .Col = 10
        .Row = 0
        .Text = month0
        .TextMatrix(1, 10) = "WIP"
        .ColWidth(10) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(2, 10) = "P5"
        
        .Col = 11
        .Row = 0
        .Text = month0
        .TextMatrix(1, 11) = "WIP"
        .ColWidth(11) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignRightCenter
        .TextMatrix(2, 11) = "P6"
        
        .Col = 12
        .Row = 0
        .Text = month0
        .TextMatrix(1, 12) = "WIP"
        .ColWidth(12) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(2, 12) = "P7"
        
        .Col = 13
        .Row = 0
        .Text = month0
        .TextMatrix(1, 13) = "WIP"
        .ColWidth(13) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignRightCenter
        .TextMatrix(2, 13) = "P8"
        
        .Col = 14
        .Row = 0
        .Text = month0
        .TextMatrix(1, 14) = "WIP"
        .ColWidth(14) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(14) = flexAlignRightCenter
        .TextMatrix(2, 14) = "P9"
        
        .Col = 15
        .Row = 0
        .Text = month0
        .TextMatrix(1, 15) = "WIP"
        .ColWidth(15) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(15) = flexAlignRightCenter
        .TextMatrix(2, 15) = "P10"
        
        .Col = 16
        .Row = 0
        .Text = month0
        .TextMatrix(1, 16) = "WIP"
        .ColWidth(16) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(16) = flexAlignRightCenter
        .TextMatrix(2, 16) = "REP. QA"
        
        .Col = 17
        .Row = 0
        .Text = month0
        .TextMatrix(1, 17) = "WIP"
        .ColWidth(17) = 0 '700
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(17) = flexAlignRightCenter
        .TextMatrix(2, 17) = "HAV"
        
        .Col = 18
        .Row = 0
        .Text = month0
        
        .ColWidth(18) = 700
        .CellAlignment = flexAlignRightCenter
        .ColAlignment(18) = flexAlignRightCenter
        .MergeCol(18) = True
        .TextMatrix(1, 18) = "F/G"
        .TextMatrix(2, 18) = "F/G"
        
        .Col = 19
        .Row = 0
        .ColWidth(19) = 1000
        .CellAlignment = flexAlignRightCenter
        .ColAlignment(19) = flexAlignRightCenter
        .MergeCol(19) = True
        .TextMatrix(0, 19) = month0
        .TextMatrix(1, 19) = "OST MPP"
        .TextMatrix(2, 19) = "OST MPP"
        
        .Col = 20
        .Row = 0
        .ColWidth(20) = 1000
        .CellAlignment = flexAlignRightCenter
        .ColAlignment(20) = flexAlignRightCenter
        .MergeCol(20) = True
        .TextMatrix(0, 20) = month0
        .TextMatrix(1, 20) = "TOTAL STOCK"
        .TextMatrix(2, 20) = "TOTAL STOCK"
        
        .Col = 21
        .Row = 0
        .ColWidth(21) = 1000
        .CellAlignment = flexAlignRightCenter
        .ColAlignment(21) = flexAlignRightCenter
        .MergeCol(21) = True
        .TextMatrix(0, 21) = month0
        .TextMatrix(1, 21) = "OST SO"
        .TextMatrix(2, 21) = "OST SO"
        
        .Col = 22
        .Row = 0
        .ColWidth(22) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(22) = flexAlignRightCenter
        .MergeCol(22) = True
        .TextMatrix(0, 22) = month1
        .TextMatrix(1, 22) = "BAL. AWAL"
        .TextMatrix(2, 22) = "BAL. AWAL"
        
        .Col = 23
        .Row = 0
        .ColWidth(23) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(23) = flexAlignRightCenter
        .MergeCol(23) = True
        .TextMatrix(0, 23) = month1
        .TextMatrix(1, 23) = "ITO"
        .TextMatrix(2, 23) = "ITO"
        
        .Col = 24
        .Row = 0
        .ColWidth(24) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(24) = flexAlignRightCenter
        .MergeCol(24) = True
        .TextMatrix(0, 24) = month1
        .TextMatrix(1, 24) = "SO"
        .TextMatrix(2, 24) = "SO"
        
        .Col = 25
        .Row = 0
        .ColWidth(25) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(25) = flexAlignRightCenter
        .MergeCol(25) = True
        .TextMatrix(0, 25) = month1
        .TextMatrix(1, 25) = "FC"
        .TextMatrix(2, 25) = "FC"
        
        .Col = 26
        .Row = 0
        .ColWidth(26) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(26) = flexAlignRightCenter
        .MergeCol(26) = True
        .TextMatrix(0, 26) = month1
        .TextMatrix(1, 26) = "DEL. RATE"
        .TextMatrix(2, 26) = "DEL. RATE"
        
        .Col = 27
        .Row = 0
        .ColWidth(27) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(27) = flexAlignRightCenter
        .MergeCol(27) = True
        .TextMatrix(0, 27) = month1
        .TextMatrix(1, 27) = "S. STOCK"
        .TextMatrix(2, 27) = "S. STOCK"
        
        .Col = 28
        .Row = 0
        .ColWidth(28) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(28) = flexAlignRightCenter
        .MergeCol(28) = True
        .TextMatrix(0, 28) = month1
        .TextMatrix(1, 28) = "NEED"
        .TextMatrix(2, 28) = "NEED"
        
        .Col = 29
        .Row = 0
        .ColWidth(29) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(29) = flexAlignRightCenter
        .MergeCol(29) = True
        .TextMatrix(0, 29) = month1
        .TextMatrix(1, 29) = "PROD. PLAN"
        .TextMatrix(2, 29) = "PROD. PLAN"
        
        .Col = 30
        .Row = 0
        .ColWidth(30) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(30) = flexAlignRightCenter
        .MergeCol(30) = True
        .TextMatrix(0, 30) = month2
        .TextMatrix(1, 30) = "BAL. AWAL"
        .TextMatrix(2, 30) = "BAL. AWAL"
        
        .Col = 31
        .Row = 0
        .ColWidth(31) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(31) = flexAlignRightCenter
        .MergeCol(31) = True
        .TextMatrix(0, 31) = month2
        .TextMatrix(1, 31) = "ITO"
        .TextMatrix(2, 31) = "ITO"
        
        .Col = 32
        .Row = 0
        .ColWidth(32) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(32) = flexAlignRightCenter
        .MergeCol(32) = True
        .TextMatrix(0, 32) = month2
        .TextMatrix(1, 32) = "SO"
        .TextMatrix(2, 32) = "SO"
        
        .Col = 33
        .Row = 0
        .ColWidth(33) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(33) = flexAlignRightCenter
        .MergeCol(33) = True
        .TextMatrix(0, 33) = month2
        .TextMatrix(1, 33) = "FC"
        .TextMatrix(2, 33) = "FC"
        
        .Col = 34
        .Row = 0
        .ColWidth(34) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(34) = flexAlignRightCenter
        .MergeCol(34) = True
        .TextMatrix(0, 34) = month2
        .TextMatrix(1, 34) = "DEL. RATE"
        .TextMatrix(2, 34) = "DEL. RATE"

        .Col = 35
        .Row = 0
        .ColWidth(35) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(35) = flexAlignRightCenter
        .MergeCol(35) = True
        .TextMatrix(0, 35) = month2
        .TextMatrix(1, 35) = "S. STOCK"
        .TextMatrix(2, 35) = "S. STOCK"
        
        .Col = 36
        .Row = 0
        .ColWidth(36) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(36) = flexAlignRightCenter
        .MergeCol(36) = True
        .TextMatrix(0, 36) = month2
        .TextMatrix(1, 36) = "NEED"
        .TextMatrix(2, 36) = "NEED"

        .Col = 37
        .Row = 0
        .ColWidth(37) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(37) = flexAlignRightCenter
        .MergeCol(37) = True
        .TextMatrix(0, 37) = month2
        .TextMatrix(1, 37) = "PROD. PLAN"
        .TextMatrix(2, 37) = "PROD. PLAN"
        
        .Col = 38
        .Row = 0
        .ColWidth(38) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(38) = flexAlignRightCenter
        .MergeCol(38) = True
        .TextMatrix(0, 38) = month3
        .TextMatrix(1, 38) = "BAL. AWAL"
        .TextMatrix(2, 38) = "BAL. AWAL"
        
        .Col = 39
        .Row = 0
        .ColWidth(39) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(39) = flexAlignRightCenter
        .MergeCol(39) = True
        .TextMatrix(0, 39) = month3
        .TextMatrix(1, 39) = "ITO"
        .TextMatrix(2, 39) = "ITO"
        
        .Col = 40
        .Row = 0
        .ColWidth(40) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(40) = flexAlignRightCenter
        .MergeCol(40) = True
        .TextMatrix(0, 40) = month3
        .TextMatrix(1, 40) = "SO"
        .TextMatrix(2, 40) = "SO"
        
        .Col = 41
        .Row = 0
        .ColWidth(41) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(41) = flexAlignRightCenter
        .MergeCol(41) = True
        .TextMatrix(0, 41) = month3
        .TextMatrix(1, 41) = "FC"
        .TextMatrix(2, 41) = "FC"
        
        .Col = 42
        .Row = 0
        .ColWidth(42) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(42) = flexAlignRightCenter
        .MergeCol(42) = True
        .TextMatrix(0, 42) = month3
        .TextMatrix(1, 42) = "DEL. RATE"
        .TextMatrix(2, 42) = "DEL. RATE"

        .Col = 43
        .Row = 0
        .ColWidth(43) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(43) = flexAlignRightCenter
        .MergeCol(43) = True
        .TextMatrix(0, 43) = month3
        .TextMatrix(1, 43) = "S. STOCK"
        .TextMatrix(2, 43) = "S. STOCK"
        
        .Col = 44
        .Row = 0
        .ColWidth(44) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(44) = flexAlignRightCenter
        .MergeCol(44) = True
        .TextMatrix(0, 44) = month3
        .TextMatrix(1, 44) = "NEED"
        .TextMatrix(2, 44) = "NEED"

        .Col = 45
        .Row = 0
        .ColWidth(45) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(45) = flexAlignRightCenter
        .MergeCol(45) = True
        .TextMatrix(0, 45) = month3
        .TextMatrix(1, 45) = "PROD. PLAN"
        .TextMatrix(2, 45) = "PROD. PLAN"
        
        .Col = 46
        .Row = 0
        .ColWidth(46) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(46) = flexAlignRightCenter
        .MergeCol(46) = True
        .TextMatrix(0, 46) = month4
        .TextMatrix(1, 46) = "BAL. AWAL"
        .TextMatrix(2, 46) = "BAL. AWAL"
        
        .Col = 47
        .Row = 0
        .ColWidth(47) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(47) = flexAlignRightCenter
        .MergeCol(47) = True
        .TextMatrix(0, 47) = month4
        .TextMatrix(1, 47) = "ITO"
        .TextMatrix(2, 47) = "ITO"
        
        .Col = 48
        .Row = 0
        .ColWidth(48) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(48) = flexAlignRightCenter
        .MergeCol(48) = True
        .TextMatrix(0, 48) = month4
        .TextMatrix(1, 48) = "SO"
        .TextMatrix(2, 48) = "SO"
        
        .Col = 49
        .Row = 0
        .ColWidth(49) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(49) = flexAlignRightCenter
        .MergeCol(49) = True
        .TextMatrix(0, 49) = month4
        .TextMatrix(1, 49) = "FC"
        .TextMatrix(2, 49) = "FC"
        
        .Col = 50
        .Row = 0
        .ColWidth(50) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(50) = flexAlignRightCenter
        .MergeCol(50) = True
        .TextMatrix(0, 50) = month4
        .TextMatrix(1, 50) = "DEL. RATE"
        .TextMatrix(2, 50) = "DEL. RATE"

        .Col = 51
        .Row = 0
        .ColWidth(51) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(51) = flexAlignRightCenter
        .MergeCol(51) = True
        .TextMatrix(0, 51) = month4
        .TextMatrix(1, 51) = "S. STOCK"
        .TextMatrix(2, 51) = "S. STOCK"
        
        .Col = 52
        .Row = 0
        .ColWidth(52) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(52) = flexAlignRightCenter
        .MergeCol(52) = True
        .TextMatrix(0, 52) = month4
        .TextMatrix(1, 52) = "NEED"
        .TextMatrix(2, 52) = "NEED"

        .Col = 53
        .Row = 0
        .ColWidth(53) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(53) = flexAlignRightCenter
        .MergeCol(53) = True
        .TextMatrix(0, 53) = month4
        .TextMatrix(1, 53) = "PROD. PLAN"
        .TextMatrix(2, 53) = "PROD. PLAN"
        
        .Col = 54
        .Row = 0
        .ColWidth(54) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(54) = flexAlignRightCenter
        .MergeCol(54) = True
        .TextMatrix(0, 54) = "BAL"
        .TextMatrix(1, 54) = "BAL"
        .TextMatrix(2, 54) = "BAL"
        
        .ColWidth(55) = 0
        .TextMatrix(0, 55) = "ST SPAREPART"
        .TextMatrix(1, 55) = "ST SPAREPART"
        .TextMatrix(2, 55) = "ST SPAREPART"
        .ColWidth(56) = 0
        .TextMatrix(0, 56) = "ST EDIT"
        .TextMatrix(1, 56) = "ST EDIT"
        .TextMatrix(2, 56) = "ST EDIT"
        .ColWidth(57) = 0
        .TextMatrix(0, 57) = "(%)Safetystock"
        .TextMatrix(1, 57) = "(%)Safetystock"
        .TextMatrix(2, 57) = "(%)Safetystock"
    End With
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

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Call WheelUnHook(Me.hwnd)
        DelTab Me
    End If
End Sub

Private Sub uploadExcel(file As String)
On Error GoTo errUpExcel
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet1 As Object
    Dim ExcelSheet2 As Object
    Dim ExcelSheet3 As Object
    Dim i As Double
    Dim ttlinactive As Integer
    Dim affQue As Integer
    Dim qry As String
    Dim fixSO As String * 30
    Dim chkUpload As Integer
    Dim li As ListItem
    
    lv1.ListItems.Clear
    
    stExcel = False
    
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet1 = CreateObject("Excel.Sheet")
    
    ExcelObj.Workbooks.Open file
    
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet1 = ExcelBook.Worksheets(1)
    Set ExcelSheet2 = ExcelBook.Worksheets(2)
    Set ExcelSheet3 = ExcelBook.Worksheets(3)
    
    clearUploadForm
    
    stExcel = True
    qry = "select item_id from mst_item where stscode_id='02'"
    Set RsGet = Con.Execute(qry)
    
    With ExcelSheet1
        l_period = Val(.Cells(2, 2))
        l_lt = Val(.Cells(3, 2))
        l_hkw(1) = Val(.Cells(4, 2))
        l_hkw(2) = Val(.Cells(4, 4))
        l_hkw(3) = Val(.Cells(4, 5))
        l_hkw(4) = Val(.Cells(4, 6))
        l_hkw(5) = Val(.Cells(4, 7))
        l_rev = Val(.Cells(5, 2))
        l_fc4 = Val(.Cells(6, 2))
    End With
    If Len(l_period) = 6 Then
        cmdUpload.Enabled = False
        Set RsBantu = Con.Execute("select count(*) c_data from ltpp_header where period = '" & l_period & "' and rev = " & l_rev)
        chkUpload = RsBantu!c_data
        RsBantu.Close
        If Val(chkUpload) = 0 Then
            'WIP
            With ExcelSheet1
                uploadStatus = "Please Wait..."
                'Set RsDoc = Con.Execute("select '" & l_period & "/LTPP/' || coalesce(lpad(cast(cast(max(substr(ltpp_doc, 13, 3)) as integer) + 1 as varchar(3)), 3, '0'), '001') doc " _
                '    & "from ltpp_header where substr(ltpp_doc, 1, 6) = '" & l_period & "'")
                Set RsDoc = Con.Execute("select coalesce(lpad(cast(cast(max(substr(ltpp_doc, 1, 2)) as integer) + 1 as varchar(2)), 2, '0'), '01') || '/BPI-INJ/LTPP/" & Right(l_period, 2) & "/" & Left(l_period, 4) & "' doc " _
                    & "from ltpp_header where period = '" & l_period & "'")
                doc = RsDoc!doc
                RsDoc.Close
                'countWIP = (.UsedRange.rows.Count - 7)
                countWIP = .Range("A" & .rows.Count).End(xlUp).Row - 7
                Con.Execute "insert into ltpp_header (ltpp_doc, period, date_period, lt, hkw_1, hkw_2, hkw_3, hkw_4, hkw_5, fc_m4, rev, upload_user, upload_time) values('" & doc & "', '" & l_period & "', '" & Left(l_period, 4) & "-" & Right(l_period, 2) & "-01', " & l_lt & ", " & l_hkw(1) & ", " & l_hkw(2) & ", " & l_hkw(3) & ", " & l_hkw(4) & ", " & l_hkw(5) & ", " & l_fc4 & ", " & l_rev & ", '" & pUserName & "', now())", affQue
                
                ttlinactive = 0
                i = 8
                Do Until .Cells(i, 1) & "" = ""
                    RsGet.Filter = "item_id='" & Trim(.Cells(i, 1)) & "'"
                    If RsGet.RecordCount > 0 Then
                        Set li = lv1.ListItems.Add(, , Trim(.Cells(i, 1)))
                    End If
                    i = i + 1
                Loop
                
                If lv1.ListItems.Count > 0 Then
                    
                End If
                
                i = 8
                Do Until .Cells(i, 1) & "" = ""
                    On Error Resume Next
                    affQue = 0
                    Con.Execute "insert into ltpp_wip (ltpp_doc, assy_no, pp, hav, ost_mpp, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, rep_qa, fg) " _
                        & "values ('" & doc & "', '" & Trim(.Cells(i, 1)) & "', " & Val(.Cells(i, 3)) & ", " & Val(.Cells(i, 15)) & ", " & Val(.Cells(i, 2)) & ", " & Val(.Cells(i, 4)) & ", " & Val(.Cells(i, 5)) & ", " & Val(.Cells(i, 6)) & ", " & Val(.Cells(i, 7)) & ", " & Val(.Cells(i, 8)) & ", " & Val(.Cells(i, 9)) & ", " & Val(.Cells(i, 10)) & ", " & Val(.Cells(i, 11)) & ", " & Val(.Cells(i, 12)) & ", " & Val(.Cells(i, 13)) & ", " & Val(.Cells(i, 14)) & ", " & Val(.Cells(i, 16)) & ")", affQue
                    If affQue = 1 Then
                        sWIP = Val(sWIP) + 1
                        .Cells(i, 17) = "SUCCESS"
                        logUploadWIP = logUploadWIP & "SUCCESS ->" & RTrim(.Cells(i, 1)) & vbCrLf
                    Else
                        fWIP = Val(fWIP) + 1
                        .Cells(i, 17) = "FAILED"
                        logUploadWIP = logUploadWIP & "FAILED  ->" & RTrim(.Cells(i, 1)) & vbCrLf
                    End If
                    progressBarWIP.Value = FormatNumber(((i - 7) * 100) / Val(countWIP), 0)
                    prcWIP = progressBarWIP.Value & "%"
                    i = i + 1
                Loop
            End With
            'SO
            With ExcelSheet2
                i = 8
                'countSO = (.UsedRange.rows.Count - 7)
                countSO = .Range("A" & .rows.Count).End(xlUp).Row - 7
                Do Until .Cells(i, 1) & "" = ""
                    On Error Resume Next
                    affQue = 0
                    fixSO = RTrim(.Cells(i, 1))
'                    Con.Execute "insert into ltpp_so (ltpp_doc, so_id, cust_id, assy_no, so_qty, so_reqdate) " _
'                        & "values ('" & doc & "', '" & Trim(.Cells(i, 1)) & "', '" & Trim(.Cells(i, 5)) & "','" & Trim(.Cells(i, 2)) & "', " & Val(.Cells(i, 3)) & ", '" & .Cells(i, 4) & "')", affQue
                    Con.Execute "insert into ltpp_so (ltpp_doc, so_id, cust_id, assy_no, so_qty, so_reqdate) " _
                        & "values ('" & doc & "', '" & Trim(.Cells(i, 1)) & "', null,'" & Trim(.Cells(i, 2)) & "', " & Val(.Cells(i, 3)) & ", '" & .Cells(i, 4) & "')", affQue
                    If affQue = 1 Then
                        sSO = Val(sSO) + 1
                        .Cells(i, 6) = "SUCCESS"
                        logUploadSO = logUploadSO & "SUCCESS ->" & fixSO & RTrim(.Cells(i, 2)) & vbCrLf
                    Else
                        fSO = Val(fSO) + 1
                        .Cells(i, 6) = "FAILED"
                        logUploadSO = logUploadSO & "FAILED  ->" & fixSO & RTrim(.Cells(i, 2)) & vbCrLf
                    End If
                    progressBarSO.Value = FormatNumber(((i - 7) * 100) / Val(countSO), 0)
                    prcSO = progressBarSO.Value & "%"
                    i = i + 1
                Loop
            End With
            'FC
            With ExcelSheet3
                i = 8
                'countFC = (.UsedRange.rows.Count - 7)
                countFC = .Range("A" & .rows.Count).End(xlUp).Row - 7
                Do Until .Cells(i, 1) & "" = ""
                    On Error Resume Next
                    affQue = 0
                    Con.Execute "insert into ltpp_fc (ltpp_doc, assy_no, fc1, fc2, fc3, fc4) " _
                        & "values ('" & doc & "', '" & .Cells(i, 1) & "', " & Val(.Cells(i, 2)) & ", " & Val(.Cells(i, 3)) & ", " & Val(.Cells(i, 4)) & ", " & RoundNumber(.Cells(i, 5)) & ")", affQue
                    If affQue = 1 Then
                        sFC = Val(sFC) + 1
                        .Cells(i, 6) = "SUCCESS"
                        logUploadFC = logUploadFC & "SUCCESS ->" & RTrim(.Cells(i, 1)) & vbCrLf
                    Else
                        fFC = Val(fFC) + 1
                        .Cells(i, 6) = "FAILED"
                        logUploadFC = logUploadFC & "FAILED  ->" & RTrim(.Cells(i, 1)) & vbCrLf
                    End If
                    progressBarFC.Value = FormatNumber(((i - 7) * 100) / Val(countFC), 0)
                    prcFC = progressBarFC.Value & "%"
                    i = i + 1
                Loop
            End With
            uploadStatus = "UPLOADED"
            cmdCancelUpload.Enabled = True
            ExcelBook.SaveAs FileName:=Replace(file, ".xlsx", "_UL" & Format(Now, "yymmdd-hhmmss") & ".xlsx")
        Else
            MsgBox "Data Sudah Ada!", vbExclamation, "Warning..."
            clearUploadForm
        End If
        
        cmdUpload.Enabled = True
    Else
        MsgBox "Periksa Periode!"
    End If
    ExcelObj.Workbooks.Close
    
    Set ExcelSheet1 = Nothing
    Set ExcelSheet2 = Nothing
    Set ExcelSheet3 = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
Exit Sub
errUpExcel:
    If Err.Number <> 0 Then
        MsgBox Err.Description & "dengan error numb =" & Err.Number
        If stExcel = True Then
            ExcelObj.Workbooks.Close
        End If
        Set ExcelSheet1 = Nothing
        Set ExcelSheet2 = Nothing
        Set ExcelSheet3 = Nothing
        Set ExcelBook = Nothing
        Set ExcelObj = Nothing
    End If
End Sub

Private Sub getListComboBox()
    Dim iMonth As Integer
    cmbMM.Clear
    For iMonth = 1 To 12
        cmbMM.AddItem Format(DateSerial(Year(Now), iMonth, 1), "MMMM")
    Next
    cmbMM.ListIndex = Month(Now) - 1
    cmbRev.Clear
End Sub

Private Sub clearUploadForm()
    doc = ""
    l_period = "-"
    l_lt = "-"
    l_hkw(1) = "-"
    l_hkw(2) = "-"
    l_hkw(3) = "-"
    l_hkw(4) = "-"
    l_hkw(5) = "-"
    l_fc4 = "-"
    l_rev = "-"
    sWIP = 0
    sSO = 0
    sFC = 0
    fWIP = 0
    fSO = 0
    fFC = 0
    countWIP = 0
    countSO = 0
    countFC = 0
    prcWIP = "0%"
    prcSO = "0%"
    prcFC = "0%"
    uploadStatus = ""
    progressBarWIP.Value = 0
    progressBarSO.Value = 0
    progressBarFC.Value = 0
    logUploadWIP = ""
    logUploadSO = ""
    logUploadFC = ""
End Sub

Private Sub MarkingRev(ByVal objSheet As Object, ByVal ColPos, ByVal rowPos)
    With objSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, ColPos, rowPos, 9, 9)
        .Fill.ForeColor.RGB = RGB(220, 220, 220)
        .line.Weight = 1
        .line.ForeColor.RGB = RGB(255, 50, 0)
    End With
End Sub

Private Sub printLTPPGroupHeader(ByVal objPrint As Printer, ByVal canvasWidth As Long, ByVal xpos As Long, ByVal Ypos As Long, ByVal cellWidth As Double, ByVal cellHeight As Double, ByVal cellTop As Double, ByVal paperScale As Long, ByVal fontScale As Double)
    objPrint.FontBold = True
    PrintCell objPrint, xpos, Ypos + (cellTop - (cellHeight * 3)), 235 * paperScale, cellHeight * 3, "NO", 3 * fontScale, True, , 2, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (250 * paperScale), Ypos + (cellTop - (cellHeight * 3)), 1100 * paperScale, cellHeight * 3, "ASSY NO", 3 * fontScale, True, , 2, 2, 30 * paperScale
    PrintCell objPrint, xpos + (1380 * paperScale), Ypos + (cellTop - (cellHeight * 3)), 1100 * paperScale, cellHeight * 3, "ASSY NAME", 3 * fontScale, True, , 2, 2, 30 * paperScale
    PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (cellTop - (cellHeight * 3)), 1300 * paperScale, cellHeight * 3, "CUSTOMER", 3 * fontScale, True, , 2, 2, 30 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale), Ypos + (cellTop - (cellHeight * 3)), cellWidth, cellHeight * 3, "MPQ", 3 * fontScale, True, , 2, 2
    
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 1), Ypos + (cellTop - (cellHeight * 3)), cellWidth * 3, cellHeight, arrMM(0), 3 * fontScale, True, , 2, 2
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 1), Ypos + (cellTop - (cellHeight * 2)), cellWidth * 3, cellHeight, "WIP", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 1), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "INJ", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 2), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "NON ASSY", 2.5 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 3), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "ASSY", 3 * fontScale, True, , 2, 2
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 4), Ypos + (cellTop - (cellHeight * 3)), cellWidth, cellHeight * 3, "FG", 3 * fontScale, True, , 2, 2
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 5), Ypos + (cellTop - (cellHeight * 3)), cellWidth, cellHeight * 3, "OST MPP", 3 * fontScale, True, , 2, 2
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 6), Ypos + (cellTop - (cellHeight * 3)), cellWidth, cellHeight * 3, , , True
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 6), Ypos + (cellTop - (cellHeight * 3)), cellWidth, cellHeight * 2, "TOTAL", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 6), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "STOCK", 3 * fontScale, False, , 2, 2
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 7), Ypos + (cellTop - (cellHeight * 3)), cellWidth, cellHeight * 3, "OST SO", 3 * fontScale, True, , 2, 2
    
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 8), Ypos + (cellTop - (cellHeight * 3)), cellWidth * 8, cellHeight, arrMM(1), 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 8), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 8), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "BALANCE", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 8), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "AWAL", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 9), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "ITO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 10), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "SO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 11), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "FC", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 12), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 12), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "DELIVERY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 12), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "RATE", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 13), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 13), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "SAFETY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 13), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "STOCK", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 14), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "NEED", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 15), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 15), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "PROD.", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 15), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "PLAN", 3 * fontScale, False, , 2, 2

    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 16), Ypos + (cellTop - (cellHeight * 3)), cellWidth * 8, cellHeight, arrMM(2), 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 16), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 16), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "BALANCE", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 16), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "AWAL", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 17), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "ITO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 18), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "SO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 19), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "FC", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 20), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 20), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "DELIVERY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 20), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "RATE", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 21), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 21), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "SAFETY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 21), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "STOCK", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 22), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "NEED", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 23), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 23), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "PROD.", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 23), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "PLAN", 3 * fontScale, False, , 2, 2

    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 24), Ypos + (cellTop - (cellHeight * 3)), cellWidth * 8, cellHeight, arrMM(3), 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 24), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 24), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "BALANCE", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 24), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "AWAL", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 25), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "ITO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 26), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "SO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 27), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "FC", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 28), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 28), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "DELIVERY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 28), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "RATE", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 29), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 29), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "SAFETY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 29), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "STOCK", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 30), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "NEED", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 31), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 31), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "PROD.", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 31), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "PLAN", 3 * fontScale, False, , 2, 2

    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 32), Ypos + (cellTop - (cellHeight * 3)), cellWidth * 8, cellHeight, arrMM(4), 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 32), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 32), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "BALANCE", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 32), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "AWAL", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 33), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "ITO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 34), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "SO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 35), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "FC", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 36), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 36), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "DELIVERY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 36), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "RATE", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 37), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 37), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "SAFETY", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 37), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "STOCK", 3 * fontScale, False, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 38), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, "NEED", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 39), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight * 2, , , True
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 39), Ypos + (cellTop - (cellHeight * 2)), cellWidth, cellHeight, "PROD.", 3 * fontScale, False, , 2, 2
            PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 39), Ypos + (cellTop - cellHeight), cellWidth, cellHeight, "PLAN", 3 * fontScale, False, , 2, 2
    
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 40), Ypos + (cellTop - (cellHeight * 3)), (canvasWidth - ((3840 * paperScale) + (cellWidth * 40))), cellHeight * 3, "BAL", 3 * fontScale, True, , 2, 2
End Sub

Private Sub printLTPP(ByVal objPrint As Printer, ByVal printWidth As Long, ByVal printHeight As Long)
    Dim xpos As Long
    Dim Ypos As Long
    Dim canvasWidth As Long
    Dim canvasHeight As Long
    Dim cellWidth As Double
    Dim cellHeight As Long
    Dim cellTop As Long
    Dim paperScale As Long
    Dim fontScale As Double
    Dim iData As Integer
    Dim iROW As Integer
    Dim posTop As Long
    Dim SumArray(0 To 39) As Variant
    
    iROW = 0
    xpos = 1 * Printer.TwipsPerPixelX
    Ypos = 1 * Printer.TwipsPerPixelY
    canvasWidth = (GetDeviceCaps(Printer.hdc, HORZRES) - 4) * Printer.TwipsPerPixelX
    canvasHeight = (GetDeviceCaps(Printer.hdc, VERTRES) - 4) * Printer.TwipsPerPixelY
    paperScale = canvasWidth / 16538 '16238 = canvas A4
    fontScale = canvasWidth / 16238 '16238 = canvas A4
    cellWidth = ((canvasWidth - 3840) / 41)
    cellHeight = 200 * fontScale
    cellTop = 2750 * fontScale
    
    'Scale Twips
    objPrint.ScaleMode = vbTwips
    PrintCell objPrint, xpos, Ypos, canvasWidth, canvasHeight, , , True
    objPrint.PaintPicture logoBEI32.Image, xpos + (800 * paperScale), Ypos + (200 * fontScale), 375 * fontScale, 375 * fontScale
    objPrint.FontBold = True
    PrintCell objPrint, xpos + canvasWidth - (930 * 3 * paperScale), Ypos, 930 * 3 * paperScale, 300 * fontScale, "FM-PPC-001-REV-02", 9 * fontScale, True, , 2, 2
    PrintCell objPrint, xpos + (1300 * paperScale), Ypos + (300 * fontScale), 3000 * fontScale, 200 * fontScale, "PT. BANSHU PLASTIC INDONESIA", 7 * fontScale, False
    objPrint.FontBold = False
    objPrint.FontItalic = True
    PrintCell objPrint, xpos + (1300 * paperScale), Ypos + (450 * fontScale), 3000 * fontScale, 200 * fontScale, "PPC SECTION", 6 * fontScale, False
    objPrint.FontBold = True
    objPrint.FontItalic = False
    PrintCell objPrint, xpos, Ypos + (800 * fontScale), canvasWidth, 150 * fontScale, "LONGTERM PRODUCTION PLANNING", 14 * fontScale, False, , 2
    PrintCell objPrint, xpos, Ypos + (1100 * fontScale), canvasWidth, 150 * fontScale, "PERIODE : " & arrMM(1), 8 * fontScale, False, , 2
    objPrint.FontBold = False
    
    
    PrintCell objPrint, xpos, Ypos + (1400 * fontScale), canvasWidth, 800 * fontScale, , , True
    
    'DIKETAHUI
    PrintCell objPrint, xpos + canvasWidth - (930 * 3 * paperScale), Ypos + (1400 * fontScale), 930 * paperScale, 800 * fontScale, , , True
        PrintCell objPrint, xpos + canvasWidth - (930 * 3 * paperScale), Ypos + (1400 * fontScale), 930 * paperScale, 130 * fontScale, "DIKETAHUI", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + canvasWidth - (930 * 3 * paperScale), Ypos + (2000 * fontScale), 930 * paperScale, 185 * fontScale, txtDiketahui, 3 * fontScale, True, , 2, 2, , , , 15 * fontScale
    'DIPERIKSA
    PrintCell objPrint, xpos + canvasWidth - (930 * 2 * paperScale), Ypos + (1400 * fontScale), 930 * paperScale, 800 * fontScale, , , True
        PrintCell objPrint, xpos + canvasWidth - (930 * 2 * paperScale), Ypos + (1400 * fontScale), 930 * paperScale, 130 * fontScale, "DIPERIKSA", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + canvasWidth - (930 * 2 * paperScale), Ypos + (2000 * fontScale), 930 * paperScale, 185 * fontScale, txtDiperiksa, 3 * fontScale, True, , 2, 2, , , , 15 * fontScale
    'DIBUAT
    PrintCell objPrint, xpos + canvasWidth - (930 * paperScale), Ypos + (1400 * fontScale), 930 * paperScale, 800 * fontScale, , , True
        PrintCell objPrint, xpos + canvasWidth - (930 * paperScale), Ypos + (1400 * fontScale), 930 * paperScale, 130 * fontScale, "DIBUAT", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + canvasWidth - (930 * paperScale), Ypos + (2000 * fontScale), 930 * paperScale, 185 * fontScale, txtDibuat, 3 * fontScale, True, , 2, 2, , , , 15 * fontScale
    
    'DOC
    PrintCell objPrint, xpos, Ypos + (1400 * fontScale), (2510 * paperScale), 800 * fontScale, , , True
        PrintCell objPrint, xpos, Ypos + (1450 * fontScale), (cellWidth * 2) - 30, 100 * fontScale, "DOC NO", 3 * fontScale, False, , , 2, 30
        PrintCell objPrint, xpos, Ypos + (1550 * fontScale), (cellWidth * 2) - 30, 100 * fontScale, "DATE", 3 * fontScale, False, , , 2, 30
        PrintCell objPrint, xpos, Ypos + (1650 * fontScale), (cellWidth * 2) - 30, 100 * fontScale, "SECTION", 3 * fontScale, False, , , 2, 30
        PrintCell objPrint, xpos, Ypos + (1750 * fontScale), (cellWidth * 2) - 30, 100 * fontScale, "DISTRIBUTION", 3 * fontScale, False, , , 2, 30
        objPrint.FontBold = True
        'DOC NO
        PrintCell objPrint, xpos + (cellWidth * 2), Ypos + (1450 * fontScale), (cellWidth * 4), 100 * fontScale, ": " & txtDocNo, 3 * fontScale, False, , , 2
        'DATE
        PrintCell objPrint, xpos + (cellWidth * 2), Ypos + (1550 * fontScale), (cellWidth * 4), 100 * fontScale, ": " & Format(dtLTPP, "DD MMMM YYYY"), 3 * fontScale, False, , , 2
        'SECTION
        PrintCell objPrint, xpos + (cellWidth * 2), Ypos + (1650 * fontScale), (cellWidth * 4), 100 * fontScale, ": INJECTION", 3 * fontScale, False, , , 2
        'DISTRIBUTION
        PrintCell objPrint, xpos + (cellWidth * 2), Ypos + (1750 * fontScale), (cellWidth * 4), 100 * fontScale, ":", 3 * fontScale, False, , , 2
        PrintCell objPrint, xpos, Ypos + (1850 * fontScale), (cellWidth * 2) - 30, 100 * fontScale, "1. MCL", 3 * fontScale, False, , , 2, 30
        PrintCell objPrint, xpos, Ypos + (1950 * fontScale), (cellWidth * 2) - 30, 100 * fontScale, "2. PRODUKSI", 3 * fontScale, False, , , 2, 30
        
    'REV
    PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (1400 * fontScale), (cellWidth * 8) + (1300 * paperScale), 800 * fontScale, , , True
        'FM
        PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (1400 * fontScale), cellWidth, 100 * fontScale, "REV NO", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (2510 * paperScale) + cellWidth, Ypos + (1400 * fontScale), (cellWidth * 7) + (1300 * paperScale), 100 * fontScale, "REMARK", 3 * fontScale, True, , 2, 2
        PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (1600 * fontScale), (cellWidth * 8) + (1300 * paperScale), 500 * fontScale, , , True
        PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (1700 * fontScale), (cellWidth * 8) + (1300 * paperScale), 300 * fontScale, , , True
        PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (1800 * fontScale), (cellWidth * 8) + (1300 * paperScale), 100 * fontScale, , , True
        PrintCell objPrint, xpos + (2510 * paperScale) + cellWidth, Ypos + (1500 * fontScale), (cellWidth * 7) + (1300 * paperScale), 700 * fontScale, , , True
        
        objPrint.FontBold = False
        Set RsDB = Con.Execute("select distinct rev, notes from ltpp_generate where period = '" & RsGet!period & "' and rev <= " & RsGet!rev & " order by rev")
        If RsDB.RecordCount <= 7 Then
            PrintCell objPrint, xpos, Ypos, canvasWidth, 2200 * fontScale, , , True
        Else
            PrintCell objPrint, xpos, Ypos, canvasWidth, (1700 + (RsDB.RecordCount * 100)) * fontScale, , , True
            cellTop = (2050 + (RsDB.RecordCount * 100)) * fontScale
        End If
        RsDB.MoveFirst
        If Not RsDB.EOF Then
            iRev = 0
            Do Until RsDB.EOF
                iRev = iRev + 1
                PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (1400 * fontScale) + (100 * fontScale * iRev), cellWidth, 100 * fontScale, RsDB!rev, 3 * fontScale, True, , 2, 2
                PrintCell objPrint, xpos + (2510 * paperScale) + cellWidth, Ypos + (1400 * fontScale) + (100 * fontScale * iRev), (cellWidth * 7) + (1300 * paperScale), 100 * fontScale, RsDB!Notes, 3 * fontScale, True, , 2, 30
                RsDB.MoveNext
            Loop
        End If
        RsDB.Close
    
    'GROUP HEADER
    printLTPPGroupHeader objPrint, canvasWidth, xpos, Ypos, cellWidth, 150, cellTop, paperScale, fontScale
    
    'DETAIL
    iData = 0
    RsGet.MoveFirst
    Do Until RsGet.EOF
        iData = iData + 1
        posTop = Ypos + (cellTop) + (cellHeight * iROW)
        objPrint.FontBold = False
        If posTop + cellHeight + 30 >= Ypos + canvasHeight Then
            objPrint.NewPage
            cellTop = Ypos + (450 * fontScale)
            iROW = 0
            PrintCell objPrint, xpos, Ypos, canvasWidth, canvasHeight, , , True
            printLTPPGroupHeader objPrint, canvasWidth, xpos, Ypos, cellWidth, 150, cellTop, paperScale, fontScale
        End If
        If holdLine <> RsGet!nm_line Then
            holdLine = RsGet!nm_line
            PrintCell objPrint, xpos, Ypos + (cellTop) + (cellHeight * iROW), 3840 * paperScale, cellHeight, holdLine, 3 * fontScale, True, , 2, 2
            iROW = iROW + 1
        End If
        If RsGet!mark <> "-" And cmbRev.Text <> "0" Then
            objPrint.PaintPicture pTriangle, xpos + (1350 * paperScale) - (120 * paperScale), Ypos + (cellTop) + (cellHeight * iROW) + (cellHeight / 4), 100 * paperScale, 100 * fontScale
            objPrint.FontBold = True
        End If
        posTop = Ypos + (cellTop) + (cellHeight * iROW)
        PrintCell objPrint, xpos, Ypos + (cellTop) + (cellHeight * iROW), 235 * paperScale, cellHeight, Str(iData), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (250 * paperScale), Ypos + (cellTop) + (cellHeight * iROW), 1100 * paperScale, cellHeight, RsGet!assy_no, 3 * fontScale, True, , , 2, 30 * paperScale
        PrintCell objPrint, xpos + (1380 * paperScale), Ypos + (cellTop) + (cellHeight * iROW), 1100 * paperScale, cellHeight, RsGet!item_name, 3 * fontScale, True, , , 2, 30 * paperScale
        PrintCell objPrint, xpos + (2510 * paperScale), Ypos + (cellTop) + (cellHeight * iROW), 1300 * paperScale, cellHeight, RsGet!cust_name, 2.5 * fontScale, True, , , 2, 30 * paperScale
        'cct ~ bal 'mpq
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 0), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!item_muloq, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 1), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!p1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 2), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!p2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 3), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!p3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 4), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!fg, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 5), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!ost_mpp, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 6), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!t_stock, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 7), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!so_0, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 8), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!bal_1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 9), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!ito_1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 10), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!so_1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 11), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!fc1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 12), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!del_rate_1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 13), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!s_stock_1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 14), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!need_1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 15), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!prod_plan_1, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 16), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!bal_2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 17), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!ito_2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 18), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!so_2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 19), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!fc2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 20), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!del_rate_2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 21), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!s_stock_2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 22), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!need_2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 23), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!prod_plan_2, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 24), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!bal_3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 25), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!ito_3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 26), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!so_3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 27), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!fc3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 28), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!del_rate_3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 29), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!s_stock_3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 30), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!need_3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 31), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!prod_plan_3, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 32), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!bal_4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 33), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!ito_4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 34), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!so_4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 35), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!fc4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 36), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!del_rate_4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 37), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!s_stock_4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 38), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!need_4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 39), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(RsGet!prod_plan_4, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 40), posTop, (canvasWidth - ((3840 * paperScale) + (cellWidth * 40))) - (15 * paperScale), cellHeight, FormatNumber(RsGet!bal_end, 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
        
        SumArray(0) = SumArray(0) + RsGet!p1
        SumArray(1) = SumArray(1) + RsGet!p2
        SumArray(2) = SumArray(2) + RsGet!p3
        SumArray(3) = SumArray(3) + RsGet!fg
        SumArray(4) = SumArray(4) + RsGet!ost_mpp
        SumArray(5) = SumArray(5) + RsGet!t_stock
        SumArray(6) = SumArray(6) + RsGet!so_0
        SumArray(7) = SumArray(7) + RsGet!bal_1
        SumArray(8) = SumArray(8) + RsGet!ito_1
        SumArray(9) = SumArray(9) + RsGet!so_1
        SumArray(10) = SumArray(10) + RsGet!fc1
        SumArray(11) = SumArray(11) + RsGet!del_rate_1
        SumArray(12) = SumArray(12) + RsGet!s_stock_1
        SumArray(13) = SumArray(13) + RsGet!need_1
        SumArray(14) = SumArray(14) + RsGet!prod_plan_1
        SumArray(15) = SumArray(15) + RsGet!bal_2
        SumArray(16) = SumArray(16) + RsGet!ito_2
        SumArray(17) = SumArray(17) + RsGet!so_2
        SumArray(18) = SumArray(18) + RsGet!fc2
        SumArray(19) = SumArray(19) + RsGet!del_rate_2
        SumArray(20) = SumArray(20) + RsGet!s_stock_2
        SumArray(21) = SumArray(21) + RsGet!need_2
        SumArray(22) = SumArray(22) + RsGet!prod_plan_2
        SumArray(23) = SumArray(23) + RsGet!bal_3
        SumArray(24) = SumArray(24) + RsGet!ito_3
        SumArray(25) = SumArray(25) + RsGet!so_3
        SumArray(26) = SumArray(26) + RsGet!fc3
        SumArray(27) = SumArray(27) + RsGet!del_rate_3
        SumArray(28) = SumArray(28) + RsGet!s_stock_3
        SumArray(29) = SumArray(29) + RsGet!need_3
        SumArray(30) = SumArray(30) + RsGet!prod_plan_3
        SumArray(31) = SumArray(31) + RsGet!bal_4
        SumArray(32) = SumArray(32) + RsGet!ito_4
        SumArray(33) = SumArray(33) + RsGet!so_4
        SumArray(34) = SumArray(34) + RsGet!fc4
        SumArray(35) = SumArray(35) + RsGet!del_rate_4
        SumArray(36) = SumArray(36) + RsGet!s_stock_4
        SumArray(37) = SumArray(37) + RsGet!need_4
        SumArray(38) = SumArray(38) + RsGet!prod_plan_4
        SumArray(39) = SumArray(39) + RsGet!bal_end
        
        iROW = iROW + 1
        RsGet.MoveNext
    Loop
    
    'FOOTER
    posTop = Ypos + (cellTop) + (cellHeight * iROW)
    objPrint.FontBold = False
    If posTop + cellHeight + 30 >= Ypos + canvasHeight Then
        objPrint.NewPage
        cellTop = Ypos + (450 * fontScale)
        iROW = 0
        PrintCell objPrint, xpos, Ypos, canvasWidth, canvasHeight, , , True
        printLTPPGroupHeader objPrint, canvasWidth, xpos, Ypos, cellWidth, 150, cellTop, paperScale, fontScale
    End If
    posTop = Ypos + (cellTop) + (cellHeight * iROW)
    PrintCell objPrint, xpos, Ypos + (cellTop) + (cellHeight * iROW), 235 * paperScale, cellHeight, "", 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (250 * paperScale), Ypos + (cellTop) + (cellHeight * iROW), 3560 * paperScale, cellHeight, "TOTAL", 3 * fontScale, True, , 2, 2, 30 * paperScale
    '-
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 0), posTop, cellWidth - (15 * paperScale), cellHeight, "", 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 1), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(0), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 2), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(1), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 3), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(2), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 4), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(3), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 5), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(4), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 6), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(5), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 7), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(6), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 8), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(7), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 9), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(8), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 10), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(9), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 11), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(10), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 12), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(11), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 13), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(12), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 14), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(13), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 15), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(14), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 16), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(15), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 17), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(16), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 18), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(17), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 19), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(18), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 20), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(19), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 21), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(20), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 22), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(21), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 23), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(22), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 24), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(23), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 25), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(24), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 26), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(25), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 27), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(26), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 28), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(27), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 29), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(28), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 30), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(29), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 31), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(30), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 32), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(31), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 33), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(32), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 34), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(33), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 35), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(34), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 36), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(35), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 37), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(36), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 38), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(37), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 39), posTop, cellWidth - (15 * paperScale), cellHeight, FormatNumber(SumArray(38), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    PrintCell objPrint, xpos + (3840 * paperScale) + (cellWidth * 40), posTop, (canvasWidth - ((3840 * paperScale) + (cellWidth * 40))) - (15 * paperScale), cellHeight, FormatNumber(SumArray(39), 0, True, True, True), 3 * fontScale, True, , 1, 2, , , 15 * paperScale
    
End Sub

Public Sub flexEditor(argFlexGrid As MSFlexGrid, KeyCode As Integer)
On Error GoTo errEditor
Dim textHolder As String
    textHolder = argFlexGrid.Text
    With argFlexGrid
        If (.Col >= 0 And .Col <= 4) Then Exit Sub
        If .Col = 20 Then Exit Sub
        If (.Col >= 22 And .Col <= 23) Then Exit Sub
        If (.Col >= 26 And .Col <= 31) Then Exit Sub
        If (.Col >= 34 And .Col <= 39) Then Exit Sub
        If (.Col >= 42 And .Col <= 47) Then Exit Sub
        If (.Col >= 50 And .Col <= 54) Then Exit Sub
        If KeyCode = 13 Or KeyCode = 9 Then Exit Sub
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
            .CellBackColor = RGB(220, 220, 75)
            .TextMatrix(.Row, 56) = "edit"
        End If
        GenerateByRow .Row
    End With
Exit Sub
errEditor:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Editor: " & Err.Number
    End If
End Sub



Private Sub grid0_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText grid0.Clip
    End If
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MousePointer = 15
End Sub

Private Sub Label35_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lX As Integer, lY As Single
    If Button = vbLeftButton Then
        pic0.Left = pic0.Left + (x / 15 - lX)
        pic0.Top = pic0.Top + (Y / 15 - lY)
    Else
        lX = x / 15: lY = Y / 15
    End If
End Sub

Private Sub Label35_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MousePointer = 0
End Sub

Private Sub Label37_Click()
    pic0.Visible = False
End Sub

Private Sub Label39_Click()
    picunprocess.Visible = False
End Sub

Private Sub MSFlexGridLTPP_KeyDown(KeyCode As Integer, Shift As Integer)
    If (ltppEDIT = True And ((KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Or KeyCode = 46 Or KeyCode = 8 Or KeyCode = 13)) Then
        If (KeyCode >= 96 And KeyCode <= 105) Then
            KeyCode = KeyCode - 48
        End If
        flexEditor MSFlexGridLTPP, KeyCode
    End If
End Sub

Private Sub GenerateByRow(ByVal rowId As Integer)
On Error GoTo errGRow
    Dim tStock As Double
    
    tStock = 0
    
    With MSFlexGridLTPP
    
'        If cacheRowAssy <> RTrim(.TextMatrix(rowId, 1)) Then
'            cacheRowAssy = RTrim(.TextMatrix(rowId, 1))
'            Set RsBantu = Con.Execute("select st_sparepart from mst_item where item_id = '" & cacheRowAssy & "'")
'            If Not RsBantu.EOF Then
'                stSPart = RsBantu!st_sparepart
'            Else
'                stSPart = 0
'            End If
'            RsBantu.Close
'        End If
        
        'generate total stock
        For iLoop = 5 To 19
            tStock = tStock + Val(.TextMatrix(rowId, iLoop))
        Next
        .TextMatrix(rowId, 20) = tStock 'total stock
        
        bal_1 = tStock - Val(.TextMatrix(rowId, 21)) 'bal M+0
        
        del_rate_1 = Round(Val(.TextMatrix(rowId, 25)) / g_hkw(1))  'delivery rate 1
        del_rate_2 = Round(Val(.TextMatrix(rowId, 33)) / g_hkw(2)) 'delivery rate 2
        del_rate_3 = Round(Val(.TextMatrix(rowId, 41)) / g_hkw(3)) 'delivery rate 3
        del_rate_4 = Round(Val(.TextMatrix(rowId, 49)) / g_hkw(4)) 'delivery rate 4
        
        s_stock_1 = RoundNumber(Val(.TextMatrix(rowId, 57)) * Val(.TextMatrix(rowId, 25)) / 100)
        s_stock_2 = RoundNumber(Val(.TextMatrix(rowId, 57)) * Val(.TextMatrix(rowId, 33)) / 100)
        s_stock_3 = RoundNumber(Val(.TextMatrix(rowId, 57)) * Val(.TextMatrix(rowId, 41)) / 100)
        s_stock_4 = RoundNumber(Val(.TextMatrix(rowId, 57)) * Val(.TextMatrix(rowId, 49)) / 100)
        
        
        'M+0
        If del_rate_1 > 0 Then
            ito_1 = bal_1 / del_rate_1
        Else
            ito_1 = 0
        End If
        If Val(.TextMatrix(rowId, 25)) > Val(.TextMatrix(rowId, 24)) Then 'fc1 > so1 = t or f
            need_1 = Val(.TextMatrix(rowId, 25)) + s_stock_1 - bal_1 ' w/ FC
        Else
            need_1 = Val(.TextMatrix(rowId, 24)) + s_stock_1 - bal_1 ' w/ SO
        End If
        
        prod_plan_1 = 0
        If Val(.TextMatrix(rowId, 4)) > 0 Then prod_plan_1 = RoundNumber(-Int(-need_1 / Val(.TextMatrix(rowId, 4))) * Val(.TextMatrix(rowId, 4)))
        
        If need_1 < 0 Then
            need_1 = 0
            prod_plan_1 = 0
        End If
            
        'M+1
        If Val(.TextMatrix(rowId, 25)) > Val(.TextMatrix(rowId, 24)) Then 'fc1 > so1 = t or f
            bal_2 = bal_1 + prod_plan_1 - Val(.TextMatrix(rowId, 25)) 'w/ FC
        Else
            bal_2 = bal_1 + prod_plan_1 - Val(.TextMatrix(rowId, 24)) 'w/ SO
        End If
        If del_rate_2 > 0 Then
            ito_2 = bal_2 / del_rate_2
        Else
            ito_2 = 0
        End If
        If Val(.TextMatrix(rowId, 33)) > Val(.TextMatrix(rowId, 32)) Then 'fc2 > so2 = t or f
            need_2 = Val(.TextMatrix(rowId, 33)) + s_stock_2 - bal_2 'w/ FC
        Else
            need_2 = Val(.TextMatrix(rowId, 32)) + s_stock_2 - bal_2 'w/ SO
        End If
        
        prod_plan_2 = 0
        If Val(.TextMatrix(rowId, 4)) > 0 Then prod_plan_2 = RoundNumber(-Int(-need_2 / Val(.TextMatrix(rowId, 4))) * Val(.TextMatrix(rowId, 4)))
        
        If need_2 < 0 Then
            need_2 = 0
            prod_plan_2 = 0
        End If

        'M+2
        If Val(.TextMatrix(rowId, 33)) > Val(.TextMatrix(rowId, 32)) Then 'fc2 > so2 = t or f
            bal_3 = bal_2 + prod_plan_2 - Val(.TextMatrix(rowId, 33)) 'w/ FC
        Else
            bal_3 = bal_2 + prod_plan_2 - Val(.TextMatrix(rowId, 32)) 'w/ SO
        End If
        If del_rate_3 > 0 Then
            ito_3 = bal_3 / del_rate_3
        Else
            ito_3 = 0
        End If
        If Val(.TextMatrix(rowId, 41)) > Val(.TextMatrix(rowId, 40)) Then 'fc3 > so3 = t or f
            need_3 = Val(.TextMatrix(rowId, 41)) + s_stock_3 - bal_3 'w/ FC
        Else
            need_3 = Val(.TextMatrix(rowId, 40)) + s_stock_3 - bal_3 'w/ SO
        End If
        
        prod_plan_3 = 0
        If Val(.TextMatrix(rowId, 4)) > 0 Then prod_plan_3 = RoundNumber(-Int(-need_3 / Val(.TextMatrix(rowId, 4))) * Val(.TextMatrix(rowId, 4)))
        
        If need_3 < 0 Then
            need_3 = 0
            prod_plan_3 = 0
        End If

        'M+3
        If Val(.TextMatrix(rowId, 41)) > Val(.TextMatrix(rowId, 40)) Then 'fc3 > so3 = t or f
            bal_4 = bal_3 + prod_plan_3 - Val(.TextMatrix(rowId, 41)) 'w/ FC
        Else
            bal_4 = bal_3 + prod_plan_3 - Val(.TextMatrix(rowId, 40)) 'w/ SO
        End If
        If del_rate_4 > 0 Then
            ito_4 = bal_4 / del_rate_4
        Else
            ito_4 = 0
        End If
        If Val(.TextMatrix(rowId, 49)) > Val(.TextMatrix(rowId, 48)) Then 'fc4 > so4 = t or f
            need_4 = Val(.TextMatrix(rowId, 49)) + s_stock_4 - bal_4 'w/ FC
        Else
            need_4 = Val(.TextMatrix(rowId, 48)) + s_stock_4 - bal_4 'w/ SO
        End If
        
        prod_plan_4 = 0
        If Val(.TextMatrix(rowId, 4)) > 0 Then prod_plan_4 = RoundNumber(-Int(-need_4 / Val(.TextMatrix(rowId, 4))) * Val(.TextMatrix(rowId, 4)))
        
        If need_4 < 0 Then
            need_4 = 0
            prod_plan_4 = 0
        End If
        
        'bal end
        If Val(.TextMatrix(rowId, 49)) > Val(.TextMatrix(rowId, 48)) Then 'fc4 > so4 = t or f
            bal_end = bal_4 + prod_plan_4 - Val(.TextMatrix(rowId, 49)) 'w/ FC
        Else
            bal_end = bal_4 + prod_plan_4 - Val(.TextMatrix(rowId, 48)) 'w/ SO
        End If
        
        .TextMatrix(rowId, 22) = bal_1
        .TextMatrix(rowId, 23) = ito_1
        .TextMatrix(rowId, 26) = del_rate_1
        .TextMatrix(rowId, 27) = s_stock_1
        .TextMatrix(rowId, 28) = need_1
        .TextMatrix(rowId, 29) = prod_plan_1
        
        .TextMatrix(rowId, 30) = bal_2
        .TextMatrix(rowId, 31) = ito_2
        .TextMatrix(rowId, 34) = del_rate_2
        .TextMatrix(rowId, 35) = s_stock_2
        .TextMatrix(rowId, 36) = need_2
        .TextMatrix(rowId, 37) = prod_plan_2
        
        .TextMatrix(rowId, 38) = bal_3
        .TextMatrix(rowId, 39) = ito_3
        .TextMatrix(rowId, 42) = del_rate_3
        .TextMatrix(rowId, 43) = s_stock_3
        .TextMatrix(rowId, 44) = need_3
        .TextMatrix(rowId, 45) = prod_plan_3
        
        .TextMatrix(rowId, 46) = bal_4
        .TextMatrix(rowId, 47) = ito_4
        .TextMatrix(rowId, 50) = del_rate_4
        .TextMatrix(rowId, 51) = s_stock_4
        .TextMatrix(rowId, 52) = need_4
        .TextMatrix(rowId, 53) = prod_plan_4
        
        .TextMatrix(rowId, 54) = bal_end
    End With
Exit Sub
errGRow:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error By Row: " & Err.Number
    End If
End Sub

Private Sub opSA_Click()
activeFrameSet True
End Sub

Private Sub opUL_Click()
    activeFrameSet False
End Sub

Private Sub getSA()
    txtDiketahui = GetINI("LTPP", "diketahui", vbNullString)
    txtDiperiksa = GetINI("LTPP", "diperiksa", vbNullString)
    txtDibuat = GetINI("LTPP", "dibuat", vbNullString)
End Sub

Private Sub picIndicator_Click()
    If pic0.Visible Then
        pic0.Visible = False
    Else
        pic0.Visible = True
    End If
End Sub

Private Sub Timer1_Timer()
    If grid0.rows > 1 Then
        If grid0.TextMatrix(1, 0) = "" Then
            picIndicator.BackColor = RGB(240, 240, 240)
        Else
            If picIndicator.BackColor = RGB(255, 212, 0) Then
                picIndicator.BackColor = RGB(255, 42, 42)
            Else
                picIndicator.BackColor = RGB(255, 212, 0)
            End If
        End If
    Else
        picIndicator.BackColor = RGB(240, 240, 240)
    End If
End Sub

Private Sub txtYear_Change()
    getRev txtYear & Format(cmbMM.ListIndex + 1, "00")
End Sub

Private Sub settingGrid()
    With grid0
        .Cols = 11
        .FixedCols = 3
        .TextMatrix(0, 0) = "No"
        .ColWidth(0) = 500
        .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(0, 1) = "Item Id"
        .ColWidth(1) = 2900
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Item Name"
        .ColWidth(2) = 3000
        .ColAlignment(0) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "FC1"
        .TextMatrix(0, 4) = "FC2"
        .TextMatrix(0, 5) = "FC3"
        .TextMatrix(0, 6) = "FC4"
        .TextMatrix(0, 7) = "SO"
        .TextMatrix(0, 8) = "Cap/Day"
        .TextMatrix(0, 9) = "MPQ"
        .TextMatrix(0, 10) = "Reg Date"
        .GridLinesFixed = flexGridFlat
    End With
End Sub
