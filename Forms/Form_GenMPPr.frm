VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_GenMPPr 
   Caption         =   "Generate MPS"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14385
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
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   Begin MSComctlLib.ListView lvprintp 
      Height          =   2415
      Left            =   1800
      TabIndex        =   70
      Top             =   2880
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4260
      SortKey         =   4
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Part No"
         Object.Width           =   3810
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Machine"
         Object.Width           =   3810
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mold"
         Object.Width           =   3810
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Issue Date"
         Object.Width           =   3810
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "x"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "y"
         Object.Width           =   529
      EndProperty
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   225
      Left            =   4680
      TabIndex        =   69
      Top             =   0
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicFIND 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   7320
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   56
      Top             =   4080
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Find Next"
         Height          =   375
         Left            =   3600
         TabIndex        =   60
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtFindNext 
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   3375
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
         TabIndex        =   58
         Top             =   0
         Width           =   4215
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
         TabIndex        =   57
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picSuggest_detail 
      Height          =   2655
      Left            =   4200
      ScaleHeight     =   2595
      ScaleWidth      =   6315
      TabIndex        =   35
      Top             =   2640
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtSugDetail 
         Height          =   2055
         Left            =   50
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Text            =   "Form_GenMPPr.frx":0000
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label lblCLOSE 
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
         Left            =   5880
         TabIndex        =   37
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Suggestion"
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
         TabIndex        =   36
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.PictureBox pic_pp_or_p 
      BackColor       =   &H0080FF80&
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   14115
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   14175
      Begin VB.CommandButton cuemd_print 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7320
         TabIndex        =   26
         Top             =   360
         Width           =   2475
      End
      Begin VB.CommandButton cuemd_printprev 
         Caption         =   "Print Preview"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   25
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label Label13 
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
         Left            =   13680
         TabIndex        =   55
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Options"
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
         TabIndex        =   54
         Top             =   0
         Width           =   13695
      End
      Begin VB.Label lblPlease 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Please wait . . ."
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox PicTrial 
      BackColor       =   &H00C0FFC0&
      Height          =   2295
      Left            =   5520
      ScaleHeight     =   2235
      ScaleWidth      =   7275
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox lblTrial_partno 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblStart 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1560
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Part No "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Trial Information"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   11040
      TabIndex        =   68
      Top             =   5040
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.PictureBox PicListMPP 
      BackColor       =   &H00C0FFC0&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   14115
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   14175
      Begin VB.CheckBox ckprodplan 
         BackColor       =   &H0080FF80&
         Caption         =   "Show data only Prod. Plan > 0"
         Height          =   375
         Left            =   9120
         TabIndex        =   65
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         TabIndex        =   47
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid fgmpp 
         Height          =   3015
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "double click or press Enter to load the data"
         Top             =   840
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   5318
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   13680
         TabIndex        =   49
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "MPS Data List"
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
         TabIndex        =   48
         Top             =   0
         Width           =   13695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxsh 
      Height          =   735
      Left            =   4200
      TabIndex        =   67
      Top             =   780
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show or hide column"
      Height          =   375
      Left            =   5160
      TabIndex        =   66
      Top             =   360
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel PGCheckLotLabel 
      Height          =   255
      Left            =   9360
      OleObjectBlob   =   "Form_GenMPPr.frx":0005
      TabIndex        =   62
      Top             =   1095
      Visible         =   0   'False
      Width           =   4695
   End
   Begin MSComctlLib.ProgressBar PGCheckLot 
      Height          =   255
      Left            =   9000
      TabIndex        =   61
      Top             =   840
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox CmbDocument 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdlu_findDoc 
      Caption         =   "..."
      Height          =   375
      Left            =   4125
      TabIndex        =   52
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6360
      Top             =   360
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
      TabIndex        =   51
      Top             =   480
      Visible         =   0   'False
      Width           =   375
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
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6840
      Top             =   360
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblstate 
      Height          =   705
      Left            =   9840
      OleObjectBlob   =   "Form_GenMPPr.frx":0079
      TabIndex        =   33
      Top             =   795
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   735
      Left            =   7440
      TabIndex        =   27
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cmdFileType 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Form_GenMPPr.frx":00E9
         Left            =   5160
         List            =   "Form_GenMPPr.frx":00F3
         Style           =   2  'Dropdown List
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Normal Mode"
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "EM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Edit Mode"
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox picSign_ovr 
         Height          =   375
         Left            =   3840
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
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
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
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
         Left            =   1320
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
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
         Left            =   4320
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
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
         Left            =   2040
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Load saved MPS"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   255
         Left            =   5160
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin MSComctlLib.ProgressBar PG1 
      Height          =   135
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form_GenMPPr.frx":010C
      Left            =   1560
      List            =   "Form_GenMPPr.frx":011C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox cmbPeriod 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox txtRevision 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid anaGrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9763
      _Version        =   393216
      BackColor       =   16777215
      FocusRect       =   2
      MergeCells      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   1560
      OleObjectBlob   =   "Form_GenMPPr.frx":0157
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_GenMPPr.frx":038B
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_GenMPPr.frx":03F1
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   2520
      OleObjectBlob   =   "Form_GenMPPr.frx":0457
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   1560
      OleObjectBlob   =   "Form_GenMPPr.frx":04B3
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form_GenMPPr.frx":051D
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid anaSubcont 
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9763
      _Version        =   393216
      BackColorBkg    =   12632256
      MergeCells      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid anaUnproc 
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9763
      _Version        =   393216
      BackColorBkg    =   12648384
      MergeCells      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid anaAssy 
      Height          =   5535
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9763
      _Version        =   393216
      BackColorBkg    =   12640511
      MergeCells      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicEditedList 
      BackColor       =   &H0080FF80&
      Height          =   3015
      Left            =   5400
      ScaleHeight     =   2955
      ScaleWidth      =   7755
      TabIndex        =   39
      Top             =   3600
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdCommitEdit 
         Caption         =   "Commit"
         Height          =   375
         Left            =   5880
         TabIndex        =   43
         Top             =   480
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid fge 
         Height          =   1935
         Left            =   45
         TabIndex        =   42
         Top             =   960
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3413
         _Version        =   393216
      End
      Begin VB.Label lblTtl_rev 
         BackColor       =   &H0080FF80&
         Caption         =   "..."
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label9 
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
         Left            =   7320
         TabIndex        =   45
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Edited List"
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
         TabIndex        =   44
         Top             =   0
         Width           =   7335
      End
   End
End
Attribute VB_Name = "Form_GenMPPr"
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
Private RsA As ADODB.Recordset
Private rsA_aks As ADODB.Recordset
Private rsAsubct As ADODB.Recordset
Private rsAunpro As ADODB.Recordset
Private rsMCtrial As ADODB.Recordset
Private rsBOM As ADODB.Recordset
Private rsB As ADODB.Recordset
Private rsWO As ADODB.Recordset
Private rsMstSub As ADODB.Recordset
Private rsMPPprev As ADODB.Recordset
Private rsa_rev As ADODB.Recordset
Private rsPurg As ADODB.Recordset
Private sd_mpsrev() As String
Private timeupdate As String



Private bDrag As Boolean
Private bT1 As Boolean
Private ttlSelCells As Long

Private temp_mch As String
Private noItemPerMesin As Integer
Private aMesinInj() As Variant
Private aMesinInj_r() As Variant
Private aMesinSubc() As Variant
Private ar_totallc() As Variant
Private nmbulan() As String
Private i As Long
Private k As Long
Private r As Long
Private c As Integer
Private tanggal As Integer
Private bhulan As Date


Private aDelv() As Variant
Private aPart() As Variant
Private aPart_sub() As Variant
Private aPart_unpc() As Variant
Private aJadwal() As Variant
Private aDayOFF() As Variant
Private aDayOvr() As Variant
Private aSuggest() As Variant

Private vTerkecil As Variant
Private vItemTerkecil As String
Private vItemTerkecil2 As String
Private indexTerkecil As Integer
Private vMesin      As String
Private oExcel      As Object 'Excel.Application
Private oBook       As Object 'Excel.Workbook
Private oSheet      As Object 'Excel.Worksheet
Private spreasheet  As String


Const formatDY As String = "yy"
Const COLSDATE As Byte = 22

Private fSO As FileSystemObject
Private st_MPP As Variant
Private st_CapDay As Long
Private st_ReqDay As Double
Private st_ReqHour As Double
Private st_OSpo As Long
Private st_PP As Long
Private st_FC As Long
Private st_ttlmpp As Long
Private st_lc As Double

Private ttlMPP As Variant


Private HKWs As Integer
Private totalHari As Integer
Private hLibur_gak As Boolean
Private vBalanc As Long, vBalancH As Long, vBalancHf As Long
Private mesinb As String
Private moldb As String
Private NoDocMPS As String
Private rev_MPS As String
Private belumSimpan As Boolean
Const WO_R As Integer = 95
Const WO_G As Integer = 186
Const WO_B As Integer = 84

Private Type POINTAPI
    x As Long
    Y As Long
End Type
Private Type oBOM
    bom_nm As String
    bom_cd As String
    bom_qty As Single
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
    lpPoint As POINTAPI) As Long
    
Dim pa As POINTAPI
Dim coBom  As oBOM
Dim coMba As oBOM
Dim coMsa As oBOM

Private emateri As String

Private em_x As Byte
Private em_y As Integer
Private em_qtyBuf As Long
Private mosX As Long
Private mosY As Long
Private posisisFind As Double
Dim in_PartOL() As String
Dim in_PartOLQTY() As Long
Dim in_partLTPP() As Long
Dim in_partFC() As Long

'TWEAKS
Dim faMachine As ADODB.Field
Dim faColor As ADODB.Field
Dim faCustomer As ADODB.Field
Dim faItemId As ADODB.Field
Dim faItemName As ADODB.Field
Dim faMold As ADODB.Field
Dim faCavity As ADODB.Field
Dim faCT As ADODB.Field
Dim faNeedMP As ADODB.Field
Dim faCapPHour As ADODB.Field
Dim faCapPShift As ADODB.Field
Dim faCapPDay As ADODB.Field
Dim faProdPlan As ADODB.Field
Dim faMPSProdPlan As ADODB.Field
Dim faFC As ADODB.Field
Dim ttlActMCH As Byte


Private Sub inSetValueZero(partNo As String)
    If (Not in_PartOL) <> -1 Then
        For i = 1 To UBound(in_PartOL)
            If partNo = in_PartOL(i) Then
                in_PartOLQTY(i) = 0
                in_partLTPP(i) = 0
                in_partFC(i) = 0
            End If
        Next
    End If
End Sub

Private Sub inSetValue(partNo As String, prodplanLtpp As Long, FC As Long)
    Dim iLoop As Double
    Dim tempTotal As Long
    If (Not in_PartOL) <> -1 Then
        
        For i = 2 To anaGrid.rows - 1
            If anaGrid.TextMatrix(i, 3) = partNo Then
                tempTotal = tempTotal + anaGrid.TextMatrix(i, 19)
            End If
        Next
        'MsgBox partno & " dengan " & prodplanLTPP & " jalan "
        For i = 1 To UBound(in_PartOL)
            If partNo = in_PartOL(i) Then
                in_PartOLQTY(i) = prodplanLtpp - tempTotal
                in_partLTPP(i) = prodplanLtpp
                in_partFC(i) = FC
            End If
        Next
    End If
End Sub

Private Function inGetValue(partNo As String) As Long
    For i = 1 To UBound(in_PartOL)
        If partNo = in_PartOL(i) Then
            inGetValue = in_PartOLQTY(i)
            Exit For
        End If
    Next
End Function

Private Function getTotalPP_ltpp() As Long
    Dim ttl As Long
    For i = 1 To UBound(in_PartOL)
        ttl = ttl + in_partLTPP(i)
    Next
    getTotalPP_ltpp = ttl
End Function

Private Function getTotalfc() As Long
    Dim ttl As Long
    For i = 1 To UBound(in_PartOL)
        ttl = ttl + in_partFC(i)
    Next
    getTotalfc = ttl
End Function

Private Sub inAddValue(partNo As String, ovr_qty As Long, ltpp_qty As Long, fc_qty As Long)
    If (Not in_PartOL) <> -1 Then
        For r = 1 To UBound(in_PartOL)
            If partNo = in_PartOL(r) Then
            
            Exit Sub
            End If
        Next
        
        ReDim Preserve in_PartOL(1 To UBound(in_PartOL) + 1) As String
        ReDim Preserve in_PartOLQTY(1 To UBound(in_PartOLQTY) + 1) As Long
        ReDim Preserve in_partLTPP(1 To UBound(in_partLTPP) + 1) As Long
        ReDim Preserve in_partFC(1 To UBound(in_partFC) + 1) As Long
        in_PartOL(UBound(in_PartOL)) = partNo
        in_PartOLQTY(UBound(in_PartOLQTY)) = ovr_qty
        in_partLTPP(UBound(in_partLTPP)) = ltpp_qty
        in_partFC(UBound(in_partFC)) = fc_qty
    Else
        ReDim in_PartOL(1 To 1) As String
        ReDim in_PartOLQTY(1 To 1) As Long
        ReDim in_partLTPP(1 To 1) As Long
        ReDim in_partFC(1 To 1) As Long
        in_PartOL(1) = partNo
        in_PartOLQTY(1) = ovr_qty
        in_partLTPP(1) = ltpp_qty
        in_partFC(1) = fc_qty
    End If
End Sub


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

Private Sub loadSuggestion()
    Dim x As Byte
    qry = "select qry1.no_mach,presen from (select no_mach,sum(lcvsmach) presen from mpp_gen_d " _
        & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' " _
        & " group by no_mach Having Sum(lcvsmach) > 100 " _
        & " order by 1 asc) qry1 left join " _
        & " ( select no_mach from mpp_setovrtime " _
        & " where EXTRACT(MONTH from wrk_date)=" & Right$(cmbPeriod, 2) & " and " _
        & " EXTRACT(YEAR from wrk_date)=" & Left$(cmbPeriod, 4) _
        & " GROUP BY no_mach " _
        & " order by no_mach asc " _
        & " ) qry2 on qry1.no_mach=qry2.no_mach " _
        & " Where qry2.no_mach Is Null " _
        & " order by 1 asc "
    Set RsBantu = Con.Execute(qry)
    Erase aSuggest
    If RsBantu.RecordCount > 0 Then
        ReDim aSuggest(1 To RsBantu.RecordCount, 1 To 2) As Variant
        x = 1
        While Not RsBantu.EOF
            aSuggest(x, 1) = RsBantu(0)
            aSuggest(x, 2) = RsBantu(1)
            x = x + 1
            RsBantu.MoveNext
        Wend
    End If
End Sub

Private Sub loadBOM()
    qry = "select  bom_par_item,bom_com_item,item_name,pfm_id,bom_qty_perassy,um_name  from mst_item a " _
        & " inner join mst_bom b on a.item_id=b.bom_com_item " _
        & " inner join r_unit_measure c on a.um_id=c.um_id"
    Set rsBOM = Con.Execute(qry)
End Sub

Private Sub loadMstSubcont()
    qry = "select kodesubcont,namasubcont from loadcap_mst_subcont"
    Set rsMstSub = Con.Execute(qry)
    rsMstSub.Fields("kodesubcont").Properties("Optimize") = True
End Sub

Private Function getDescSubcont(pId As String) As String
    rsMstSub.Fields("kodesubcont").Properties("Optimize") = True
    rsMstSub.Filter = adFilterNone
    rsMstSub.Filter = "kodesubcont='" & pId & "'"
    If rsMstSub.RecordCount > 0 Then
        getDescSubcont = rsMstSub!namasubcont
    End If
End Function

Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhLastDayInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 0)
End Function

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

Private Function isi(pMPQ As Double, pCapPDay As Variant, atasBawah As String)
    Private MPQ As Variant
    Private bReach As Boolean
    
    bReach = True
    MPQ = pMPQ
    
    If pMPQ = 0 Then
        isi = 0
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

Private Sub settingFG()
    With fge
        .Cols = 18
        .rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        
        i = 0
        .TextMatrix(0, i) = "MC ":
        .ColWidth(i) = 700
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 1
        .TextMatrix(0, i) = "Customer":
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800
        
        
        i = 2
        .TextMatrix(0, i) = "Part No"::
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2500
        
        i = 3
        .TextMatrix(0, i) = "Part Name"
        .ColWidth(i) = 2700
        
        i = 4
        .TextMatrix(0, i) = "No Mold"
        .ColWidth(i) = 800
        
        i = 5
        .TextMatrix(0, i) = "CAV"
        .ColWidth(i) = 600
        
        i = 6
        .TextMatrix(0, i) = "C/T (sec)"::
         .ColWidth(i) = 700
        
         i = 7
        .TextMatrix(0, i) = "Need MP (persons)"
        .ColWidth(i) = 1000
        
         i = 8
        .TextMatrix(0, i) = "Capacity /Hour (pcs)"
        .ColWidth(i) = 1200
        
        i = 9
        .TextMatrix(0, i) = "Capacity /Shift (pcs)":
        
        
        i = 10
        .TextMatrix(0, i) = "Capacity /days (pcs)":
       
        
        i = 11
        .TextMatrix(0, i) = "Hour Req (hours)":
        .ColWidth(i) = 900
        
        i = 12
        .TextMatrix(0, i) = "Day Req (days)":
        .ColWidth(i) = 900
        
        i = 13
        .TextMatrix(0, i) = "O/s PO (pcs)":
        .ColWidth(i) = 850
        
        i = 14
        .TextMatrix(0, i) = "Prod Plan (pcs)":
        
        i = 15
        .TextMatrix(0, i) = "FC":
    
        i = 16
        .TextMatrix(0, i) = "Tgl"
        .ColWidth(i) = 800
        
        i = 17
        .TextMatrix(0, i) = "Qty"
        .ColWidth(i) = 800
        
        
    End With
    With fgmpp
        .Cols = 6
        .FixedCols = 1
        .TextMatrix(0, 0) = "No"
        .ColWidth(0) = 500
        .TextMatrix(0, 1) = "MPS Doc No"
        .ColWidth(1) = 3000
        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Rev"
        .ColWidth(2) = 500
        .TextMatrix(0, 3) = "Period"
'        .ColWidth(3) = 0
        .TextMatrix(0, 4) = "Revisi LTPP"
        .ColWidth(4) = 0
        .TextMatrix(0, 5) = "LTPP Doc No"
        .ColWidth(5) = 3000
        .ColAlignment(5) = flexAlignLeftCenter
    End With
    With anaGrid
        .Cols = 22 '19 21
        .rows = 5
        .FixedRows = 2
        .FixedCols = 3
        .WordWrap = False
        .ColAlignment(2) = flexAlignLeftCenter

        .MergeCells = flexMergeRestrictRows

        i = 0
        .TextMatrix(0, i) = "MC ":    .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True: .ColWidth(i) = 700
        .ColAlignment(i) = flexAlignLeftCenter

        i = 1
        .TextMatrix(0, i) = "Customer":     .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2800

        i = 2
        .TextMatrix(0, i) = ".":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 500

        i = 3
        .TextMatrix(0, i) = "Part No":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2500

        i = 4
        .TextMatrix(0, i) = "Part Name":     .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 2700

        i = 5
        .TextMatrix(0, i) = "No Mold":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True:    .ColWidth(i) = 800


        i = 6
        .TextMatrix(0, i) = "Color":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True:    .ColWidth(i) = 800
        

        i = 7
        .TextMatrix(0, i) = "CAV":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 600


        i = 8
        .TextMatrix(0, i) = "C/T (sec)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True: .ColWidth(i) = 700

         i = 9
        .TextMatrix(0, i) = "Need MP (persons)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True:   .ColWidth(i) = 1000

         i = 10
        .TextMatrix(0, i) = "Capacity /Hour (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True: .ColWidth(i) = 1200

        i = 11
        .TextMatrix(0, i) = "Capacity /Shift (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 12
        .TextMatrix(0, i) = "Capacity /days (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 13
        .TextMatrix(0, i) = "Hour Req (hours)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 900

        i = 14
        .TextMatrix(0, i) = "Day Req (days)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 900

        i = 15
        .TextMatrix(0, i) = "O/s PO (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 850

         i = 16
        .TextMatrix(0, i) = "Prod Plan LTPP (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 17
        .TextMatrix(0, i) = "Prod Plan (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 18
        .TextMatrix(0, i) = "Overload (pcs)":        .TextMatrix(1, i) = "Overload (pcs)"
        .MergeCol(i) = True

        i = 19
        .TextMatrix(0, i) = "FC":     .TextMatrix(1, i) = "FC"
        .MergeCol(i) = True


        i = 20
        .TextMatrix(0, i) = "Total MPP":        .TextMatrix(1, i) = "Total MPP"
        .MergeCol(i) = True
        .ColWidth(i) = 800

        i = 21
        .TextMatrix(0, i) = "Load Vs Cap":          .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750

        .MergeRow(0) = True
        .MergeRow(1) = False

        .WordWrap = True
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellAlignment = flexAlignCenterCenter
            .Row = 1
            .CellAlignment = flexAlignCenterCenter
        Next
    End With
    
    
    With anaSubcont
        .Cols = 20: .rows = 5
        .FixedRows = 2: .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        
        i = 0
        .TextMatrix(0, i) = "MC ":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 700
        
        i = 1
        .TextMatrix(0, i) = "Customer":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 3000
        
        i = 2
        .TextMatrix(0, i) = "Part No":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        .ColWidth(i) = 2500
        
        i = 3
        .TextMatrix(0, i) = "Part Name":         .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 3000
        
        i = 4
        .Col = i
        .Row = 0
        .Text = "Min Ton":         .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 550
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCells = flexMergeRestrictAll
        
        i = 5
        .TextMatrix(0, i) = "Max Ton":       .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 550
        
        
        i = 6
        .TextMatrix(0, i) = "No Mold":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 800
        
        i = 7
        .TextMatrix(0, i) = "CAV":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 550
        
        i = 8
        .TextMatrix(0, i) = "C/T (sec)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 700
        
         i = 9
        .TextMatrix(0, i) = "Need MP (persons)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 950
        
         i = 10
        .TextMatrix(0, i) = "Capacity /Hour (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 1050
        
        i = 11
        .TextMatrix(0, i) = "Capacity /Shift (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 12
        .TextMatrix(0, i) = "Capacity /days (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        
        i = 13
        .TextMatrix(0, i) = "Hour Req (hours)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 14
        .TextMatrix(0, i) = "Day Req (days)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 15
        .TextMatrix(0, i) = "O/s PO (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 850
        
        i = 16
        .TextMatrix(0, i) = "Prod Plan (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 17
        .TextMatrix(0, i) = "FC":        .TextMatrix(1, i) = "FC"
        .MergeCol(i) = True
       
        
        i = 18
        .TextMatrix(0, i) = "Total":        .TextMatrix(1, i) = "MPP"
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 19
        .TextMatrix(0, i) = "Load Vs Cap":          .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
              
        .MergeRow(0) = True
        .MergeRow(1) = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellAlignment = flexAlignCenterCenter
            .Row = 1
            .CellAlignment = flexAlignCenterCenter
        Next
    End With
    With anaUnproc
        .Cols = 18: .ColWidth(0) = 700: .ColWidth(1) = 3000: .ColWidth(2) = 2500: .ColWidth(3) = 3000
        .ColWidth(5) = 800: .ColWidth(6) = 700: .ColWidth(4) = 2000:
        .ColWidth(6) = 1000
        .ColWidth(7) = 1200: .ColWidth(8) = 1200: .ColWidth(9) = 1200
        .rows = 5
        .FixedRows = 2
        .FixedCols = 4
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        
        i = 0
        .TextMatrix(0, i) = "MC ":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 1
        .TextMatrix(0, i) = "Customer":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 2
        .TextMatrix(0, i) = "Part No":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 3
        .TextMatrix(0, i) = "Part Name":         .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = "No Mold":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = "CAV":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 700
        
        i = 6
        .TextMatrix(0, i) = "C/T (sec)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
         i = 7
        .TextMatrix(0, i) = "Need MP (persons)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
         i = 8
        .TextMatrix(0, i) = "Capacity /Hour (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 9
        .TextMatrix(0, i) = "Capacity /Shift (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 10
        .TextMatrix(0, i) = "Capacity /days (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        
        i = 11
        .TextMatrix(0, i) = "Hour Req (hours)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 12
        .TextMatrix(0, i) = "Day Req (days)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 13
        .TextMatrix(0, i) = "O/s PO (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 850
        
        i = 14
        .TextMatrix(0, i) = "Prod Plan (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 15
        .TextMatrix(0, i) = "FC":        .TextMatrix(1, i) = "FC"
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 16
        .TextMatrix(0, i) = "Total":        .TextMatrix(1, i) = "MPP"
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 17
        .TextMatrix(0, i) = "Load Vs Cap":          .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
              
        .MergeRow(0) = True
        .MergeRow(1) = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellAlignment = flexAlignCenterCenter
            .Row = 1
            .CellAlignment = flexAlignCenterCenter
        Next
    End With
    With anaAssy
        .Cols = 18: .ColWidth(0) = 700: .ColWidth(1) = 3000: .ColWidth(2) = 2500: .ColWidth(3) = 3000
        .ColWidth(5) = 800: .ColWidth(6) = 700: .ColWidth(4) = 2000:
        .ColWidth(6) = 1000
        .ColWidth(7) = 1200: .ColWidth(8) = 1200: .ColWidth(9) = 1200
        .rows = 5
        .FixedRows = 2
        .FixedCols = 4
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows

        
        i = 0
        .TextMatrix(0, i) = "MC ":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 1
        .TextMatrix(0, i) = "Customer":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 2
        .TextMatrix(0, i) = "Part No":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColAlignment(i) = flexAlignLeftCenter
        
        i = 3
        .TextMatrix(0, i) = "Part Name":         .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = "No Mold":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = "CAV":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        .ColWidth(i) = 700
        
        i = 6
        .TextMatrix(0, i) = "C/T (sec)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
         i = 7
        .TextMatrix(0, i) = "Need MP (persons)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
         i = 8
        .TextMatrix(0, i) = "Capacity /Hour (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i):
        .MergeCol(i) = True
        
        i = 9
        .TextMatrix(0, i) = "Capacity /Shift (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 10
        .TextMatrix(0, i) = "Capacity /days (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        
        i = 11
        .TextMatrix(0, i) = "Hour Req (hours)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 12
        .TextMatrix(0, i) = "Day Req (days)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 1000
        
        i = 13
        .TextMatrix(0, i) = "O/s PO (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 850
        
        i = 14
        .TextMatrix(0, i) = "Prod Plan (pcs)":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 15
        .TextMatrix(0, i) = "FC":        .TextMatrix(1, i) = ""
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 16
        .TextMatrix(0, i) = "Total":        .TextMatrix(1, i) = "MPP"
        .MergeCol(i) = True
        .ColWidth(i) = 750
        
        i = 17
        .TextMatrix(0, i) = "Load Vs Cap":          .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        .ColWidth(i) = 750
              
        .MergeRow(0) = True
        .MergeRow(1) = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 0
            .CellAlignment = flexAlignCenterCenter
            .Row = 1
            .CellAlignment = flexAlignCenterCenter
        Next
    End With
End Sub

Public Function DaysInMonth(ByVal dDate As Date) As Integer
    DaysInMonth = Day(DateAdd("m", 1, dDate - Day(dDate) + 1) - 1)
End Function

Private Sub anaGrid_Click()
    Dim Ret As Long
    Dim smach As String
    Dim topM As Long
    Dim samePlan As Boolean
    samePlan = False
    Ret = GetCursorPos(pa)
    With anaGrid
        posisisFind = .RowSel
    If getEditModeStatus = False Then
        If .Col > (COLSDATE - 1) And LenB(.TextMatrix(.Row, 0)) <> 0 Then
            smach = .TextMatrix(.Row, 0)
            For topM = .Row - 1 To 0 Step -1 ' find to the first row
                If .TextMatrix(topM, 0) = smach Then
                    If IsNumeric(.TextMatrix(topM, .Col)) Then
                        If .TextMatrix(topM, .Col) * 1 > 0 And .TextMatrix(topM, .Col) * 1 = .TextMatrix(topM, 12) * 1 Then '11
                            MsgBox "please, change planned qty  [" & .TextMatrix(topM, .Col) & "]", vbExclamation
                            If MsgBox("Do you want to continue ?", vbQuestion + vbYesNo) = vbNo Then
                                
                                anaGrid.SetFocus
                                .Row = topM
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
            For topM = .Row + 1 To .rows - 1 'find to the last row
                If .TextMatrix(topM, 0) = smach Then
                    If IsNumeric(.TextMatrix(topM, .Col)) Then
                        If .TextMatrix(topM, .Col) * 1 > 0 Then
                            MsgBox "please, change planned qty  +[" & .TextMatrix(topM, .Col) & "]", vbExclamation
                            If MsgBox("Do you want to continue ?", vbQuestion + vbYesNo) = vbNo Then
                                
                                anaGrid.SetFocus
                                .Row = topM
                                
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
'            smach = .TextMatrix(.Row, 0)
            If .CellBackColor <> RGB(WO_R, WO_G, WO_B) Then 'Jika wo belum turun
                If IsNumeric(.TextMatrix(.Row, .Col)) Then
                    txtEdit.Text = .Text * 1
                    em_qtyBuf = .Text * 1
                Else
                    txtEdit.Text = 0
                    em_qtyBuf = 0
                End If
                txtEdit.Visible = True
                txtEdit.SelStart = 0
                txtEdit.SelLength = Len(txtEdit.Text)
'                PicEm.Visible = True 1
'                PicEm.Left = mosX 1 'pa.x - (pa.x * 25 / 100) '* Screen.TwipsPerPixelX / 2
'                PicEm.Top = mosY - PicEm.Height 1' pa.y - (pa.y * 45 / 100) '* Screen.TwipsPerPixelY / 2

                txtEdit.Height = .cellHeight / 15
                txtEdit.Top = ScaleY(.cellTop + .Top, vbTwips, vbPixels) + 126 '(cltop * (195 / 100))
                txtEdit.Left = ScaleX(.CellLeft + .Left, vbTwips, vbPixels) + 10
                txtEdit.Width = (.cellWidth / 15)
                txtEdit.Text = .Text
                txtEdit.SelStart = 0
                txtEdit.SelLength = Len(txtEdit.Text)

                txtEdit.SetFocus
                'MsgBox txtEdit.Left & " dan y=" & txtEdit.Top
                em_x = .Col
                em_y = .Row
            End If
        Else
            txtEdit.Visible = False
        End If
    Else
        If .CellBackColor = RGB(214, 255, 3) Then
            PicTrial.Left = pa.x * Screen.TwipsPerPixelX / 2 '- (PicTrial.ScaleWidth / 2)
            PicTrial.Top = pa.Y * Screen.TwipsPerPixelY / 2 '* Screen.TwipsPerPixelY '+ PicTrial.ScaleHeight
            PicTrial.Visible = True
            lblTrial_partno = getTrialPart(.TextMatrix(.Row, 0), Left$(.TextMatrix(1, .Col), 2))
            lblStart = getTrialTime(.TextMatrix(.Row, 0), Left$(.TextMatrix(1, .Col), 2))
        Else
            PicTrial.Visible = False
        End If
        pic_pp_or_p.Visible = False
    End If
    End With
End Sub

Private Sub anaGrid_DblClick()
    Dim kolTerDbClick As Byte
    Dim rowTerDbClick As Long

    With anaGrid
        kolTerDbClick = .Col
        rowTerDbClick = .Row
        If belumSimpan And .Col > (COLSDATE - 1) Then MsgBox "Please save the data first": cmdSave.SetFocus: Exit Sub
        If IsNumeric(.Text) And .Col > (COLSDATE - 1) Then
            If .Text * 1 > 0 And .TextMatrix(.Row, 0) <> "" Then
                If .CellBackColor = RGB(WO_R, WO_G, WO_B) Then
                    MsgBox "WO tersebut telah diturunkan !", vbInformation
                    Exit Sub
                End If

                If getEditModeStatus Then
                    For i = .Col - 1 To COLSDATE Step -1
                        .Col = i
                        If .CellBackColor <> RGB(WO_R, WO_G, WO_B) And IsNumeric(.Text) Then
                            If .Text * 1 > 0 Then
                                If .CellBackColor <> RGB(172.38, 233.51, 235.62) Then
                                MsgBox "WO sebelumnya belum dicetak, cetak dulu WO sebelumnya"
                                .SetFocus
                                Exit Sub
                                End If
                            End If
                        End If
                    Next
                    'MsgBox kolTerDbClick
                    If checklotPerItemid(.TextMatrix(.Row, 3), CByte(Left$(.TextMatrix(1, kolTerDbClick), 2))) Then Exit Sub
                    .Row = rowTerDbClick
                    .Col = kolTerDbClick
                    If .Text * 1 > 0 Then
                        Dim li As ListItem
                        If LVCheckPK(.Col, .Row) = False Then
                            Set li = lvprintp.ListItems.Add(, , .TextMatrix(.Row, 3))
                            li.SubItems(4) = .Col
                            li.SubItems(5) = .Row
                        End If
                        .CellBackColor = RGB(172.38, 233.51, 235.62) ' Biru
                        If pic_pp_or_p.Visible Then
                            pic_pp_or_p.Visible = False
                        Else
                            pic_pp_or_p.Visible = True
                            cuemd_print.Enabled = True
                            lblPlease.Visible = False
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub anaGrid_GotFocus()
    MDI_Parent.mnuFreezColumn.Visible = True
    MDI_Parent.mnuFontSize.Visible = True
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

Private Sub anaGrid_KeyPress(KeyAscii As Integer)
    If getEditModeStatus = False Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            anaGrid_Click

        End If
    End If
End Sub

Private Sub anaGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim cont As Integer
    Dim found As Boolean
    Dim ttlKolom As Byte
    Dim ttlBaris As Long
    If (Button = vbLeftButton) And (Shift = 0) Then
        With anaGrid
            'Do not proceed if the row clicked is
            'not the header row.
            If (.RowHeight(0) + .RowHeight(1) < Y) Then
                Exit Sub
            End If
            ttlKolom = .Cols
            ttlBaris = .rows
            'Initialize variables
            cont = 0
            found = False
          
        End With
    ElseIf (Button = vbLeftButton) And (Shift = 2) Then
        Dim li As ListItem
        With anaGrid
            If .Col > (COLSDATE - 1) Then
                If IsNumeric(.TextMatrix(.Row, .Col)) Then
                    If .Text * 1 > 0 And .TextMatrix(.Row, 0) <> "" Then
                        If .CellBackColor = RGB(WO_R, WO_G, WO_B) Then
                            bDrag = False ' kemarin 23
                        Else
                            If .CellBackColor <> RGB(172.38, 233.51, 235.62) Then
                                .CellBackColor = RGB(172.38, 233.51, 235.62) 'biru asin
                                bT1 = True
                                If LVCheckPK(.Col, .Row) = False Then
                                    Set li = lvprintp.ListItems.Add(, , .TextMatrix(.Row, 3))
                                    li.SubItems(4) = .Col
                                    li.SubItems(5) = .Row
                                End If
                            Else
                                .CellBackColor = RGB(255, 255, 255)
                                bT1 = False
                                For i = lvprintp.ListItems.Count To 1 Step -1
                                    If lvprintp.ListItems(i).SubItems(4) = .Col And lvprintp.ListItems(i).SubItems(5) = .Row Then
                                        lvprintp.ListItems.Remove i
                                    End If
                                Next
                            End If
                            bDrag = True
                        End If

                    End If
                Else
                    bDrag = False
                    bT1 = False
                End If
            End If
        End With
    End If
End Sub

Private Function LVCheckPK(PX As Byte, PY As Long) As Boolean
    Dim B1 As Byte
    With lvprintp
        For B1 = 1 To .ListItems.Count
            If .ListItems(B1).SubItems(4) = PX And .ListItems(B1).SubItems(5) = PY Then
                LVCheckPK = True
                Exit Function
            End If
        Next
    End With
    LVCheckPK = False
End Function



Private Sub anaGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If bDrag Then
        Dim li As ListItem
        With anaGrid
            .Row = .MouseRow
            .Col = .MouseCol
            If IsNumeric(.Text) Then
                If .Text * 1 > 0 And .TextMatrix(.Row, 0) <> "" Then
                    If bT1 Then
                        If .CellBackColor <> RGB(WO_R, WO_G, WO_B) Then 'warna wo running
                            .CellBackColor = RGB(172.38, 233.51, 235.62) ' biru langit
                            If LVCheckPK(.Col, .Row) = False Then
                                    Set li = lvprintp.ListItems.Add(, , .TextMatrix(.Row, 3))
                                    li.SubItems(4) = .Col
                                    li.SubItems(5) = .Row
                            End If
                        End If
                    Else
                        If .CellBackColor <> RGB(WO_R, WO_G, WO_B) Then
                            .CellBackColor = RGB(255, 255, 255)
                            For i = lvprintp.ListItems.Count To 1 Step -1
                                If lvprintp.ListItems(i).SubItems(4) = .Col And lvprintp.ListItems(i).SubItems(5) = .Row Then
                                    lvprintp.ListItems.Remove i
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub anaGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    bDrag = False
    bT1 = False
End Sub

Private Sub anaGrid_Scroll()
    txtEdit.Visible = False
End Sub

Private Sub anaSubcont_Click()
    Dim Ret As Long
    Ret = GetCursorPos(pa)
    With anaSubcont
    If getEditModeStatus = False Then
        If .Col > 19 And LenB(.TextMatrix(.Row, 0)) <> 0 Then
            If IsNumeric(.TextMatrix(.Row, .Col)) Then
                txtEdit.Text = .Text * 1
                em_qtyBuf = .Text * 1
            Else
                txtEdit.Text = 0
                em_qtyBuf = 0
            End If
            txtEdit.SelStart = 0
            txtEdit.SelLength = Len(txtEdit.Text)
            txtEdit.Visible = True
'            PicEm.Visible = True 1
'            PicEm.Left = pa.X * Screen.TwipsPerPixelX / 2
'            PicEm.Top = pa.Y * Screen.TwipsPerPixelY / 2
            txtEdit.SetFocus
            em_x = .Col
            em_y = .Row
        Else
'            PicEm.Visible = False 1
            txtEdit.Visible = False
        End If
    End If
    End With
End Sub

Private Sub anaSubcont_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText anaSubcont.Clip
    End If
End Sub


Private Sub DAYtoGrid(FG As MSFlexGrid, pperiod As String)
    Dim indexTgl As Integer
    FG.rows = 2
    If FG.Name = "anaGrid" Then
        FG.Cols = COLSDATE
        indexTgl = COLSDATE - 1 '21
    ElseIf FG.Name = "anaSubcont" Then
        FG.Cols = 20
        indexTgl = 19
    ElseIf FG.Name = "anaUnproc" Then
        FG.Cols = 18
        indexTgl = 17
    ElseIf FG.Name = "anaAssy" Then
        FG.Cols = 18
        indexTgl = 17
    End If

    FG.Cols = indexTgl + 1 + totalHari
    With FG
        For i = 1 To Int(Format(dhLastDayInMonth(bhulan), "dd"))
            .TextMatrix(0, indexTgl + i) = Format(DateSerial(Left$(pperiod, 4), Right$(pperiod, 2), i), "ddd")
            .TextMatrix(1, indexTgl + i) = Format(DateSerial(Left$(pperiod, 4), Right$(pperiod, 2), i), "dd-mmm")
        Next
    End With
End Sub


Private Sub mc_dlm_Tgl(iTgl As Integer, dBaris As Long)
    Dim qi_mch As Byte 'integer
    With anaGrid
        For qi_mch = 1 To UBound(aMesinInj)
            If .TextMatrix(dBaris, 0) = aMesinInj(qi_mch, 1) Then
                aMesinInj(qi_mch, 3 + iTgl) = 1
                Exit For
            End If
        Next
    End With
End Sub

Private Sub mc_dlm_Tgl_byname(iTgl As Integer, mesin As String)
    Dim qi_mch As Byte 'integer
    For qi_mch = 1 To UBound(aMesinInj)
        If mesin = aMesinInj(qi_mch, 1) Then
            aMesinInj(qi_mch, 3 + iTgl) = 1
            Exit For
        End If
    Next
End Sub

Private Sub mc_dlm_Tgl_subc(iTgl As Integer, dBaris As Long)
    Dim qi_mch As Byte 'integer
    With anaSubcont
        For qi_mch = 1 To UBound(aMesinSubc)
'            List3.AddItem "* If " & .TextMatrix(dBaris, 0) & "=" & aMesinInj(qi_mch, 1) & " Then"
            If .TextMatrix(dBaris, 0) = aMesinSubc(qi_mch, 1) Then
                aMesinSubc(qi_mch, 3 + iTgl) = 1
                Exit For
            End If
        Next
    End With
End Sub

Private Sub part_dlm_tgl(iTgl As Integer, dBaris As Long)
    Dim qi_prt As Long
    With anaGrid
        For qi_prt = 1 To UBound(aPart)
            If .TextMatrix(dBaris, 3) = aPart(qi_prt, 1) And .TextMatrix(dBaris, 0) = aPart(qi_prt, 10) And .TextMatrix(dBaris, 5) = aPart(qi_prt, 7) Then

                aPart(qi_prt, 11 + iTgl) = 1
                Exit For
            End If
        Next
    End With
End Sub

Private Sub part_dlm_tgl_subc(iTgl As Integer, dBaris As Long)
    Dim qi_prt As Long
    With anaSubcont
        For qi_prt = 1 To UBound(aPart_sub)
            If .TextMatrix(dBaris, 2) = aPart_sub(qi_prt, 1) And _
            .TextMatrix(dBaris, 0) = aPart_sub(qi_prt, 10) Then
                aPart_sub(qi_prt, 10 + iTgl) = 1
                Exit For
            End If
        Next
    End With
End Sub

Private Function mc_blm_jln(iTgl As Integer) As String
    Dim qi As Byte 'Integer
    For qi = 1 To UBound(aMesinInj)
        If aMesinInj(qi, iTgl + 3) = 0 Then
            mc_blm_jln = aMesinInj(qi, 1)
            Exit For
        Else
            mc_blm_jln = ""
        End If
    Next
End Function

Private Function mcsub_blm_jln(iTgl As Integer) As String
    Dim qi As Byte 'Integer
    For qi = 1 To UBound(aMesinSubc)
        If aMesinSubc(qi, iTgl + 3) = 0 Then
            mcsub_blm_jln = aMesinSubc(qi, 1)
            Exit For
        Else
            mcsub_blm_jln = ""
        End If
    Next
End Function

Private Function mold_blm_jln(ps_tgl As Integer) As String
    Dim sl1 As Integer
    For sl1 = 1 To UBound(aPart_sub)
        If aPart_sub(sl1, 10 + ps_tgl) = 0 Then
            mold_blm_jln = aPart_sub(sl1, 7)
            Exit For
        Else
            mold_blm_jln = ""
        End If
    Next
End Function


Private Function is_mc_run(iTgl As Integer, p_mc As String) As Variant
    Dim qi2 As Byte 'Integer
    For qi2 = 1 To UBound(aMesinInj)
        If aMesinInj(qi2, 1) = p_mc Then
            is_mc_run = aMesinInj(qi2, iTgl + 3)
            Exit For
        End If
    Next
End Function

Private Function is_mc_run_subc(iTgl As Integer, p_mc As String) As Variant
    Dim qi2 As Byte 'Integer
    For qi2 = 1 To UBound(aMesinSubc)
        If aMesinSubc(qi2, 1) = p_mc Then
            is_mc_run_subc = aMesinSubc(qi2, iTgl + 3)
            Exit For
        End If
    Next
End Function

Private Function getMaterialMCH(pMCH As String) As String
    Dim t As Byte
    t = 1
    Do
        If t > UBound(aMesinInj) Then Exit Do
        If aMesinInj(t, 1) = pMCH Then
            getMaterialMCH = aMesinInj(t, 2)
            Exit Do
        End If
        t = t + 1
    Loop
End Function

Private Function getTonage(pmesin As String) As Integer
    Dim aa As Byte
    For aa = 1 To UBound(aMesinInj)
        If aMesinInj(aa, 1) = pmesin Then
            getTonage = CInt(aMesinInj(aa, 3))
            emateri = aMesinInj(aa, 2)
            Exit For
        Else
            getTonage = 0
            emateri = ""
        End If
    Next
End Function

Private Function hitungKeBelakang(pind As Long, pindK As Long) As Double
    Dim q As Integer
    Dim ttlRun As Double
    ttlRun = 0
    With anaGrid
        For q = pindK To COLSDATE Step -1
            If IsNumeric(.TextMatrix(pind, q)) Then
                ttlRun = (.TextMatrix(pind, q) * 1) + ttlRun
            End If
        Next
        hitungKeBelakang = ttlRun
    End With
End Function

Private Function hitungKeBelakang_subc(rowtoCount As Long, fromCol As Integer, ttlNow As Boolean) As Long
    Dim q As Integer
    Dim ttlRun As Long
    ttlRun = 0
    With anaSubcont
        For q = fromCol To 20 Step -1
            If IsNumeric(.TextMatrix(rowtoCount, q)) Then
                ttlRun = (.TextMatrix(rowtoCount, q) * 1) + ttlRun
            End If
        Next
        If ttlNow Then
            hitungKeBelakang_subc = ttlRun
        Else
            hitungKeBelakang_subc = ttlRun + .TextMatrix(rowtoCount, 12) * 1
        End If
    End With
End Function

Private Sub addStokFor2morrow(iTgl As Integer)
    Dim qi3 As Long
    With anaGrid
        For qi3 = 2 To .rows - 1
            If Len(.TextMatrix(qi3, (COLSDATE - 1) + iTgl)) > 0 Then
                For k = 1 To UBound(aPart)
                    If aPart(k, 1) = .TextMatrix(qi3, 3) Then
                        aPart(k, 3) = aPart(k, 3) * 1 + (.TextMatrix(qi3, (COLSDATE - 1) + iTgl) * 1)

                    End If
                Next
            End If
        Next
    End With
End Sub

Private Sub addStokFor2morrow_subc(iTgl As Integer)
    Dim qi3 As Long
    With anaSubcont
        For qi3 = 2 To .rows - 1
            If Len(.TextMatrix(qi3, 19 + iTgl)) > 0 Then
                For k = 1 To UBound(aPart_sub)
                    If aPart_sub(k, 1) = .TextMatrix(qi3, 3) Then
                        aPart_sub(k, 3) = aPart_sub(k, 3) * 1 + (.TextMatrix(qi3, 19 + iTgl) * 1)
                        
                    End If
                Next
            End If
        Next
    End With
End Sub

Private Function libur_gak(iTgl As Integer) As Boolean
    If (Not aDayOFF) <> -1 Then
        For c = 1 To UBound(aDayOFF)
            If aDayOFF(c) = iTgl Then
                libur_gak = True
                Exit For
            Else
                libur_gak = False
            End If
        Next
    Else
        libur_gak = False
    End If
End Function

Private Function ovr_gak(iTgl As Integer, pmesin As String) As Boolean
    If (Not aDayOvr) <> -1 Then
        For c = 1 To UBound(aDayOvr)
            If aDayOvr(c, 2) = iTgl And aDayOvr(c, 1) = pmesin Then
                ovr_gak = True
                Exit For
            Else
                ovr_gak = False
            End If
        Next
    Else
        ovr_gak = False
    End If
End Function

Private Function isTrial(pmesin As String, pTGL As Integer) As Boolean
    rsMCtrial.Fields("mch").Properties("Optimize") = True
    rsMCtrial.Fields("tgl").Properties("Optimize") = True
    rsMCtrial.Filter = adFilterNone
    rsMCtrial.Filter = "mch='" & pmesin & "' and tgl=" & pTGL
'    If ptgl = 23 Then
'        MsgBox rsMCtrial.RecordCount & vbNewLine & "mch='" & pMesin & "' and tgl=" & ptgl, vbInformation, "kau"
'    End If
    If rsMCtrial.RecordCount > 0 Then
        isTrial = True
    Else
        isTrial = False
    End If
End Function

Private Sub plot_subc_v2(pTGL As Integer)
    Dim catchFutDay As Long
    rsAsubct.Fields("reg_mold").Properties("Optimize") = True
    While Len(mold_blm_jln(pTGL)) > 0
        moldb = mold_blm_jln(pTGL)
        rsAsubct.Filter = adFilterNone
        rsAsubct.Filter = "reg_mold='" & moldb & "'"
        'catchFutDay = tigaHariKedepan(rsA("lcd_itemdid"), ptgl, vBalanc)
        
        hLibur_gak = libur_gak(tanggal)

        With anaSubcont
            If hLibur_gak = False Then
                For i = 2 To .rows - 1
                    If .TextMatrix(i, 6) = moldb Then
'                        MsgBox "fffd"
                        If hitungKeBelakang_subc(i, 19 + pTGL, False) <= .TextMatrix(i, 16) * 1 Then
                            .TextMatrix(i, 19 + pTGL) = FormatNumber(.TextMatrix(i, 12), 0)
                            setPartDone_s pTGL, moldb, .TextMatrix(i, 0)
                            
                        Else
                            If hitungKeBelakang_subc(i, 19 + pTGL, True) <= .TextMatrix(i, 16) * 1 Then
                                .TextMatrix(i, 19 + pTGL) = FormatNumber(.TextMatrix(i, 12), 0)
                                setPartDone_s pTGL, moldb, .TextMatrix(i, 0)
                            Else
                                setPartDone_s pTGL, moldb, .TextMatrix(i, 0)
                            End If
                        End If
                    End If
                Next
            Else
                setPartDone_s pTGL, moldb, rsAsubct("mesin")
            End If
        End With
    Wend
End Sub

Private Function getProdplan(pPART As String, pmoold As String, pmsin As String) As Long
    Dim rr As Long, tempPP As Long
    rr = 1
    Do
        If rr > UBound(aPart) Then Exit Do

        If aPart(rr, 1) = pPART And aPart(rr, 7) = pmoold And aPart(rr, 10) = pmsin Then
            tempPP = aPart(rr, 11) * 1
            Exit Do
        Else
            tempPP = 0
        End If
        rr = rr + 1
    Loop
    getProdplan = tempPP
    
    
End Function

Private Sub setMchDone(p_tgl As Integer, p_meusin As String)
    Dim wi As Byte 'integer
    For wi = 1 To UBound(aMesinInj)
        If p_meusin = aMesinInj(wi, 1) Then
            aMesinInj(wi, 3 + tanggal) = 1
        End If
    Next
End Sub

Private Sub setPartDone(p_tgl As Integer, p_meusin As String, p_paeurt As String, p_mould As String)
    Dim wi2 As Long
    For wi2 = 1 To UBound(aPart)
        If p_paeurt = aPart(wi2, 1) And p_meusin = aPart(wi2, 10) And p_mould = aPart(wi2, 7) Then
            aPart(wi2, 11 + p_tgl) = 1
            Exit For
        End If
    Next
End Sub

Private Sub setPartDone_s(p_tgl As Integer, p_mould As String, pmesin As String)
    Dim wi3 As Long
    For wi3 = 1 To UBound(aPart_sub)
        If p_mould = aPart_sub(wi3, 7) And pmesin = aPart_sub(wi3, 10) Then
            aPart_sub(wi3, 10 + p_tgl) = 1
            Exit For
        End If
    Next
End Sub

Private Sub setRod(inx As Long, patgl As Integer, pengurang As Single) 'Rod = rest of day
    Dim uh As Byte
    For uh = 1 To UBound(aMesinInj_r)
        If aMesinInj_r(uh, 1) = anaGrid.TextMatrix(inx, 0) Then
            aMesinInj_r(uh, patgl + 1) = aMesinInj_r(uh, patgl + 1) * 1 - pengurang * 1
        End If
    Next
End Sub

Private Function getRod(inx As Long, patgl As Integer) As Single
    Dim uh As Byte
    For uh = 1 To UBound(aMesinInj_r)
        If aMesinInj_r(uh, 1) = anaGrid.TextMatrix(inx, 0) Then
            getRod = aMesinInj_r(uh, patgl + 1)
            Exit For
        End If
    Next
End Function

Private Sub plot_v2(tanggal_p As Integer)
    Dim catchNeqty As Long
    Dim curMold As String, curMold2 As String
    Dim ttlRunB4 As Long
    Dim mesin_sbl As String
    While Len(mc_blm_jln(tanggal_p)) > 0
        mesinb = mc_blm_jln(tanggal_p)
        RsA.Filter = adFilterNone
        RsA.Filter = "no_mach='" & mesinb & "'"
        vTerkecil = 999999999999999#
        For i = 1 To RsA.RecordCount
            RsA.AbsolutePosition = i
            If RsA("neqty") > 0 Then
                vBalanc = getBalance(RsA("lcd_itemdid"))
                vBalancH = tigaHariKedepan(RsA("lcd_itemdid"), tanggal_p, vBalanc)
                If getPartinDate(RsA("lcd_itemdid"), tanggal_p, mesinb, RsA("reg_mold")) Then 'jika belum terplot
                    If vBalancH <= vTerkecil Then
                        vTerkecil = vBalanc
                        vItemTerkecil = RsA("lcd_itemdid")
                        curMold = RsA("reg_mold")
                        curMold2 = curMold
                        vItemTerkecil2 = vItemTerkecil
                        vBalancHf = vBalancH
                    End If
                End If
            End If
        Next
        catchNeqty = getProdplan(vItemTerkecil, curMold, mesinb)
        If catchNeqty = 0 Then
            If Len(vItemTerkecil) > 1 Then
                If vItemTerkecil <> "N/A" Then
                    setPartDone tanggal_p, mesinb, vItemTerkecil, curMold
                End If
            Else
                setMchDone tanggal_p, mesinb
            End If
        End If
        
        If vItemTerkecil = "N/A" Then
           setMchDone tanggal_p, mesinb
        End If
        If tanggal_p > 1 Then
            Dim wa As Long
            With anaGrid
                For wa = 2 To .rows - 1
                    If .TextMatrix(wa, 0) = mesinb And IsNumeric(.TextMatrix(wa, (COLSDATE - 1) + tanggal_p - 1)) Then ' 18 + tanggal_p -1 apakah tanggal sebelumnya jalan KAMU
                        ttlRunB4 = hitungKeBelakang(wa, (COLSDATE - 1) + totalHari) 'tanggal_p
                        If ttlRunB4 < .TextMatrix(wa, 17) * 1 Then '16
                            If .TextMatrix(wa, 3) <> vItemTerkecil Then
                                If vBalancHf >= 0 Then
                                    vItemTerkecil = .TextMatrix(wa, 3)
                                    curMold = .TextMatrix(wa, 5)
                                    For i = 2 To .rows - 1
                                        If .TextMatrix(i, 0) = mesinb And .TextMatrix(i, 3) = .TextMatrix(wa, 3) And .TextMatrix(i, 5) = .TextMatrix(wa, 5) Then 'vItemTerkecil
                                            If hitungKeBelakang(i, 20 + totalHari) >= .TextMatrix(i, 17) * 1 Then 'tanggal_p 16
                                                vItemTerkecil = vItemTerkecil2
                                                curMold = curMold2
                                            Else
                                                
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                Next
            End With
        End If
        
        hLibur_gak = libur_gak(tanggal_p)
        With anaGrid
            For i = 2 To .rows - 1
                If .TextMatrix(i, 3) = vItemTerkecil And .TextMatrix(i, 0) = mesinb And .TextMatrix(i, 5) = curMold Then  '
                    For r = 1 To UBound(aPart)
                        If vItemTerkecil = aPart(r, 1) And aPart(r, 10) = mesinb And aPart(r, 7) = curMold Then
                            If hLibur_gak Then  ' jika tgl libur
                                If ovr_gak(tanggal_p, .TextMatrix(i, 0)) Then    ' jika overtime
                                    If is_mc_run(tanggal_p, .TextMatrix(i, 0)) = 0 Then   ' jika mesin belum terplot
                                        If aPart(r, 5) = 0 Then ' jika item pbox 0
                                            ttlRunB4 = hitungKeBelakang(i, (COLSDATE - 1) + totalHari) 'tanggal_p
                                            If ttlRunB4 < .TextMatrix(i, 17) * 1 Then '16
                                                If ttlRunB4 + .TextMatrix(i, 12) * 1 > .TextMatrix(i, 17) * 1 Then '16 11
                                                    .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = .TextMatrix(i, 17) * 1 - ttlRunB4 '16
                                                    setRod i, tanggal_p, (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) / .TextMatrix(i, 12)) '11
                                                    
                                                Else
                                                    .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = FormatNumber(isi(aPart(r, 4) * 1, .TextMatrix(i, 12) * getRod(i, tanggal_p), "a"), 0) '11
                                                    mc_dlm_Tgl tanggal_p, i
                                                End If
                                                
                                                part_dlm_tgl tanggal_p, i
                                                If isTrial(.TextMatrix(i, 0), tanggal_p) Then
                                                    .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) * 1) - .TextMatrix(i, 11) * 1 '10
                                                End If
                                                vItemTerkecil = "N/A"
                                                curMold = "N/A"
                                            Else
                                                part_dlm_tgl tanggal_p, i
                                                vItemTerkecil = "N/A"
                                                curMold = "N/A"
                                            End If
                                        Else
                                            ttlRunB4 = hitungKeBelakang(i, (COLSDATE - 1) + totalHari) 'tanggal_p
                                            If ttlRunB4 < .TextMatrix(i, 17) * 1 Then '16
                                                If ttlRunB4 + .TextMatrix(i, 12) * 1 > .TextMatrix(i, 17) * 1 Then '11 16
                                                    .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = .TextMatrix(i, 17) * 1 - ttlRunB4 '16
                                                    setRod i, tanggal_p, (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) / .TextMatrix(i, 12)) '11
                                                    
                                                Else
                                                    .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = FormatNumber(isi(aPart(r, 5) * 1, .TextMatrix(i, 12) * getRod(i, tanggal_p), "a"), 0) ' 11
                                                    mc_dlm_Tgl tanggal_p, i
                                                End If
                                                
                                                part_dlm_tgl tanggal_p, i
                                                If isTrial(.TextMatrix(i, 0), tanggal_p) Then
                                                    .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) * 1) - .TextMatrix(i, 11) * 1 '10
                                                End If
                                                vItemTerkecil = "N/A"
                                                curMold = "N/A"
                                            Else
                                                part_dlm_tgl tanggal_p, i
                                                vItemTerkecil = "N/A"
                                                curMold = "N/A"
                                            End If
                                        End If
                                    Else
                                        part_dlm_tgl tanggal_p, i
                                        vItemTerkecil = "N/A"
                                        curMold = "N/A"
                                    End If
                                Else
                                    mc_dlm_Tgl tanggal_p, i
                                    part_dlm_tgl tanggal_p, i
                                    vItemTerkecil = "N/A"
                                    curMold = "N/A"
                                End If
                            Else
                                If is_mc_run(tanggal_p, .TextMatrix(i, 0)) = 0 Then    ' jika mesin belum terplot
                                    If aPart(r, 5) = 0 Then ' jika item mpqbox 0
                                        ttlRunB4 = hitungKeBelakang(i, (COLSDATE - 1) + totalHari) 'tanggal_p
                                        If ttlRunB4 < .TextMatrix(i, 17) * 1 Then '16
                                            If ttlRunB4 + .TextMatrix(i, 12) * 1 > .TextMatrix(i, 17) * 1 Then '11 16
                                                .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = .TextMatrix(i, 17) * 1 - ttlRunB4 '16
                                                setRod i, tanggal_p, (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) / .TextMatrix(i, 12)) '11
                                            Else
                                                .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = FormatNumber(isi(aPart(r, 4) * 1, .TextMatrix(i, 12) * getRod(i, tanggal_p), "a"), 0) '11
                                                mc_dlm_Tgl tanggal_p, i
                                            End If
                                            
                                            part_dlm_tgl tanggal_p, i
                                            If isTrial(.TextMatrix(i, 0), tanggal_p) Then
                                                .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) * 1) - .TextMatrix(i, 11) * 1 '10
                                            End If
                                            vItemTerkecil = "N/A"
                                            curMold = "N/A"
                                        Else
                                            part_dlm_tgl tanggal_p, i
                                            vItemTerkecil = "N/A"
                                            curMold = "N/A"
                                        End If
                                    Else
                                        ttlRunB4 = hitungKeBelakang(i, (COLSDATE - 1) + totalHari) 'tanggal_p
                                        If ttlRunB4 < .TextMatrix(i, 17) * 1 Then '16
                                            If ttlRunB4 + .TextMatrix(i, 12) * 1 > .TextMatrix(i, 17) * 1 Then '11 17
                                                .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = .TextMatrix(i, 17) * 1 - ttlRunB4 '16
                                                setRod i, tanggal_p, (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) / .TextMatrix(i, 12)) '11
                                                'MsgBox getRod(i, tanggal_p), vbInformation, "mpq" & .TextMatrix(i, 0)
                                            Else
                                                .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = FormatNumber(isi(aPart(r, 5) * 1, .TextMatrix(i, 12) * getRod(i, tanggal_p), "a"), 0) '11
                                                mc_dlm_Tgl tanggal_p, i
                                            End If
                                            
                                            part_dlm_tgl tanggal_p, i
                                            If isTrial(.TextMatrix(i, 0), tanggal_p) Then
                                                .TextMatrix(i, (COLSDATE - 1) + tanggal_p) = (.TextMatrix(i, (COLSDATE - 1) + tanggal_p) * 1) - .TextMatrix(i, 11) * 1 '10
                                            End If
                                            vItemTerkecil = "N/A"
                                            curMold = "N/A"
                                        Else
                                            part_dlm_tgl tanggal_p, i
                                            vItemTerkecil = "N/A"
                                            curMold = "N/A"
                                        End If
                                    End If
                                Else
                                    part_dlm_tgl tanggal_p, i
                                    vItemTerkecil = "N/A"
                                    curMold = "N/A"
                                End If
                            End If
                        End If
                    Next
                Else
    
                End If
            Next
        
        End With
    Wend
    
End Sub

Private Sub plotMchTrial(pgrid As MSFlexGrid) '
    qry = "select distinct on(mch,date_trial) mch, date_trial::date,extract(day from date_trial) tgl,date_trial dari,date_trialf sampai,part_no from " _
            & " (select part_no,mch,date_trial,date_trialf,extract(epoch  from (date_trialf-date_trial))/60/60/24 hari from mpp_ste " _
            & " where extract(MONTH from date_trial)=" & Right$(cmbPeriod, 2) & " and extract(YEAR from date_trial)=" & Left$(cmbPeriod, 4) & ")  dd " _
            & " order by mch asc"
    Set rsMCtrial = Con.Execute(qry)
    rsMCtrial.Fields("mch").Properties("Optimize") = True
    rsMCtrial.Fields("tgl").Properties("Optimize") = True
    Dim ttlKol As Byte
    Dim ttlBar As Long
    ttlKol = pgrid.Cols - 1
    ttlBar = pgrid.rows - 1
    For k = 1 To rsMCtrial.RecordCount
        rsMCtrial.AbsolutePosition = k
        For i = 2 To ttlBar
            For r = COLSDATE To ttlKol '18
                If pgrid.TextMatrix(i, 0) = rsMCtrial(0) And Left$(pgrid.TextMatrix(1, r), 2) = Format(rsMCtrial(1), "dd") Then
                    pgrid.Col = r
                    pgrid.Row = i
                    pgrid.CellBackColor = RGB(214, 255, 3)
                End If
            Next
        Next
    Next
End Sub

Private Function tigaHariKedepan_sub(pPART As Variant, ptgal As Integer, pbal As Long) As Long
    Dim cii As Integer, stokhari As Integer
    Dim c_bal1 As Long
    Dim Hubond As Long
    Dim cii2 As Long
    cii = ptgal
    c_bal1 = pbal
    stokhari = 1
    cii2 = 1
    Hubond = UBound(aJadwal)

    Do
        If cii > totalHari Then Exit Do
        hLibur_gak = libur_gak(cii)
        If hLibur_gak = False Then 'jika g libur
            If stokhari > 3 Then 'jika pengecekan sudah 3 hari
                Exit Do
            End If
            cii2 = 1
            Do
                If cii2 >= Hubond Then Exit Do
                If aJadwal(cii2, 1) = cii And aJadwal(cii2, 2) = pPART Then 'jika ada jadwal 'jika  ' tgl, part, qty
                    c_bal1 = c_bal1 - (aJadwal(cii2, 3))
                End If
                cii2 = cii2 + 1
            Loop

            stokhari = stokhari + 1
        End If
        cii = cii + 1
    Loop

    tigaHariKedepan_sub = c_bal1

End Function

Private Function getMinStock(parPart As String) As Byte
    Dim ii As Long
    For ii = 1 To UBound(aPart)
        If aPart(ii, 1) = parPart Then
            getMinStock = aPart(ii, 8)
            Exit For
        End If
    Next
End Function

Private Function tigaHariKedepan(pPART As Variant, ptgal As Integer, pbal As Long) As Long
    Dim cii As Integer, stokhari As Integer
    Dim c_bal1 As Long
    Dim Hubond As Long
    Dim cii2 As Long
    cii = ptgal
    c_bal1 = pbal
    stokhari = 1
    cii2 = 1
    Hubond = UBound(aJadwal)
    Do
        If cii > totalHari Then Exit Do
        hLibur_gak = libur_gak(cii)
        If hLibur_gak = False Then 'jika g libur
            If stokhari > getMinStock(CStr(pPART)) Then  'jika pengecekan sudah n hari
                Exit Do
            End If
            cii2 = 1
            Do
                If cii2 >= Hubond Then Exit Do
                If aJadwal(cii2, 1) = cii And aJadwal(cii2, 2) = pPART Then 'jika ada jadwal 'jika  ' tgl, part, qty
                    c_bal1 = c_bal1 - (aJadwal(cii2, 3))
                End If
                cii2 = cii2 + 1
            Loop
            stokhari = stokhari + 1
        End If
        cii = cii + 1
    Loop
    tigaHariKedepan = c_bal1

End Function

Private Function getPartinDate(papart As Variant, tgl As Integer, pmesin As String, pmld As String) As Boolean
    Dim cui As Long
    For cui = 1 To UBound(aPart)
        If aPart(cui, 11 + tgl) = 0 And aPart(cui, 1) = papart And aPart(cui, 10) = pmesin And aPart(cui, 7) = pmld Then

            getPartinDate = True
            Exit For
        Else
            getPartinDate = False
        End If
    Next
End Function

Private Function getPartinDate_subc(papart As Variant, tgl As Integer, pmesin As String) As Boolean
    Dim cui As Long
    For cui = 1 To UBound(aPart_sub)
        If aPart_sub(cui, 10 + tgl) = 0 And aPart_sub(cui, 1) = papart And aPart_sub(cui, 10) = pmesin Then
'            MsgBox aPart(cui, 9 + tgl), vbInformation, "Yihaa"
            getPartinDate_subc = True
            Exit For
        Else
            getPartinDate_subc = False
        End If
    Next
End Function

Private Function getBalance(part As Variant) As Long
    Dim rr As Long
    rr = 1
    Do
        If rr > UBound(aPart) Then Exit Do
        If aPart(rr, 1) = part Then
            getBalance = aPart(rr, 3) * 1
            Exit Do
        End If
        rr = rr + 1
    Loop
End Function

Private Function getBalance_sub(part As Variant) As Long
    Dim rr As Long
    rr = 1
    Do
        If rr > UBound(aPart_sub) Then Exit Do
        If aPart_sub(rr, 1) = part Then
            getBalance_sub = aPart_sub(rr, 3) * 1
            Exit Do
        End If
        rr = rr + 1
    Loop
End Function


Private Function getStsDay(part As Variant, pTGL As Integer) As Variant
    Dim rr As Long
    rr = 1
    Do
        If rr > UBound(aPart) Then Exit Do
        If aPart(rr, 1) = part Then
            getStsDay = aPart(rr, 10 + pTGL)
            Exit Do
        End If
        rr = rr + 1
    Loop
End Function

Private Sub sinkronGridcols()
        flxsh.rows = 3
        flxsh.RowHeight(1) = 0
        flxsh.Cols = anaGrid.Cols
        flxsh.FixedCols = 0
        With anaGrid
            For k = 0 To .Cols - 1
                flxsh.TextMatrix(0, k) = .TextMatrix(0, k)
                flxsh.TextMatrix(1, k) = .ColWidth(k)
                flxsh.Row = 2
                flxsh.Col = k
                flxsh.CellFontName = "Wingdings"
                flxsh.Text = "þ"
            Next
        End With
End Sub

Private Sub Check1_Click()
    If Check1.Value Then
        flxsh.Visible = True
        flxsh.SetFocus
        
    Else
        flxsh.Visible = False
    End If
End Sub




Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        lvprintp.Visible = True
        MsgBox txtRevision
    Else
        lvprintp.Visible = False
    End If
End Sub

Private Sub cmbPeriod_DropDown()
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    qry = "select distinct on (fltpp_ym) fltpp_ym from mpp_gen_d where fltpp_doc='" & CmbDocument & "'" _
    & " and fltpp_rev='" & txtRevision & "'"
    Set RsA = Con.Execute(qry)
    cmbPeriod.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            cmbPeriod.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub

Private Sub formatWarnaBG()
    Dim j As Integer
    For j = 3 To anaGrid.rows - 1
        anaGrid.Col = COLSDATE
        anaGrid.Row = j
        anaGrid.CellBackColor = vbWhite
    Next
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

Private Sub cmbType_Click()
    If cmbType.ListIndex = 0 Then
        anaGrid.Visible = True
        anaSubcont.Visible = False
        anaUnproc.Visible = False
        anaAssy.Visible = False
    ElseIf cmbType.ListIndex = 1 Then
        anaGrid.Visible = False
        anaSubcont.Visible = True
        anaUnproc.Visible = False
        anaAssy.Visible = False
    ElseIf cmbType.ListIndex = 2 Then
        anaGrid.Visible = False
        anaSubcont.Visible = False
        anaUnproc.Visible = True
        anaAssy.Visible = False
    Else
        anaGrid.Visible = False
        anaSubcont.Visible = False
        anaUnproc.Visible = False
        anaAssy.Visible = True
    End If
End Sub

Private Function getEditModeStatus() As Boolean
    If Option1.Value Then  'cmdEditMode.Caption = "Edit Mode"
        getEditModeStatus = False
    Else
        getEditModeStatus = True
        '-----here the query
        
    End If
End Function

Private Sub cmdCommitEdit_Click()
    If MsgBox("Are you sure", vbQuestion + vbYesNo, "Decide") = vbNo Then Exit Sub
    PicEditedList.Visible = False
    Me.Refresh
    Dim rk As Long
    Dim ttlRow As Long
    saveMPS
    
    With fge
        ttlRow = .rows - 1
        For rk = 1 To ttlRow
            If LenB(.TextMatrix(rk, 0)) <> 0 Then
                If checkKeyReve(.TextMatrix(rk, 0), .TextMatrix(rk, 2), .TextMatrix(rk, 4), .TextMatrix(rk, 16), .TextMatrix(rk, 17)) = False Then
                    qry = "insert into mpp_gen_rev (mps_doc,mps_rev,periode,hkw,machine,itemid,mold,plandate,planqty,timeediting) values " _
                        & " ('" & NoDocMPS & "','" & rev_MPS & "','" & cmbPeriod & "'," & HKWs & ",'" & .TextMatrix(rk, 0) & "','" & .TextMatrix(rk, 2) & "','" & .TextMatrix(rk, 4) & "','" & .TextMatrix(rk, 16) & "','" _
                        & .TextMatrix(rk, 17) & "',DEFAULT)"
                    Con.Execute qry
                End If
            End If
        Next
        .rows = 1
    End With
    MsgBox "Updated"
End Sub

Private Function checkKeyReve(p_mch As String, p_item As String, p_mold As String, p_tgl As String, p_qty As Long) As Boolean
    qry = "select count(*) from mpp_gen_rev where mps_doc='" & NoDocMPS & "' " _
    & " and mps_rev='" & rev_MPS & "' and machine='" & p_mch & "' and itemid='" & p_item & "' " _
    & " and mold='" & p_mold & "' and plandate='" & p_tgl & "' and planqty=" & p_qty
    Set RsGet = Con.Execute(qry)
    If RsGet(0) > 1 Then
        checkKeyReve = True
    Else
        checkKeyReve = False
    End If
End Function

Private Sub cmdDelete_Click()
    Dim pesan As String
    Dim deleted_rows As Long
    pesan = "Are you sure want to delete " & vbNewLine _
    & NoDocMPS & " revisi " & rev_MPS
    If MsgBox(pesan, vbQuestion + vbYesNo) = vbYes Then
        qry = "delete from mpp_gen where mpp_doc_no='" & NoDocMPS & "' and mpp_revisi='" & rev_MPS & "'"
        Con.Execute qry, deleted_rows
        MsgBox deleted_rows & " row(s) were deleted", vbInformation
        qry = "delete from mpp_gen_rev where mps_doc='" & NoDocMPS & "' and mps_rev='" & rev_MPS & "'"
        Con.Execute qry, deleted_rows
        MsgBox deleted_rows & " row(s) were deleted", vbInformation, "MPS Log"
        txtfind_KeyPress 13
    End If
    fgmpp.SetFocus
End Sub

Private Sub cmdExport_Click()
On Error GoTo Nah
    If belumSimpan Then MsgBox "Please save the data first", vbExclamation: cmdSave.SetFocus: Exit Sub
    Dim clLeft As Double
    Dim cltop As Double
    Dim clWidth As Double
    Dim clHeight As Double
    Dim brs1 As Double
    Dim totalBaris As Double
    Dim s As Object
    Dim c_material As String
    
    Dim ir As Double
    Dim nourut As Byte
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If LenB(CommonDialog1.FileName) <> 0 Then
        If cmdFileType.ListIndex = 0 Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        Set oExcel = CreateObject(spreasheet) 'New Excel.Application
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)

        Dim cl As Range
        Dim shpOval As Object
        Dim reng As Range
        Dim ttlKolm As Byte
        
        ttlKolm = anaGrid.Cols
        oExcel.DisplayAlerts = False
        Set fSO = New FileSystemObject
        If fSO.FileExists(App.Path & "\logoBPI.png") Then
            'With oSheet.Shapes.AddPicture(FileName:=App.Path & "\logoBPI.png", LINKTOFILE:=msoFalse, savewithdocument:=msoCTrue, Left:=0, Top:=0, Width:=45, Height:=45)
            With oSheet.Shapes.AddPicture(FileName:=App.Path & "\logoBPI.png", LINKTOFILE:=0, savewithdocument:=1, Left:=0, Top:=0, Width:=45, Height:=45)
                .Left = oSheet.Cells(2, 3).Left
                .Top = oSheet.Cells(2, 3).Top
                .Placement = 1
            End With
        Else
            MsgBox "Tidak bisa menemukan file logoBPI.png di " & App.Path
        End If
        PG1.Visible = True
        PG1.Value = 0
        
        oSheet.Columns(1).ColumnWidth = 3.14
        oSheet.Cells(3, 4) = "PT. Banshu Plastic Indonesia"
        oSheet.Cells(2, 10) = "Doc. No."
        oSheet.Cells(3, 10) = "Month"
        oSheet.Cells(4, 10) = "Date"
        oSheet.Cells(5, 10) = "Revise"
        
        oSheet.Cells(2, 11) = NoDocMPS
        oSheet.Cells(3, 11) = ": " & Format(DateSerial(Left$(cmbPeriod, 4), Right$(cmbPeriod, 2), 1), "mmmm yyyy")
        oSheet.Cells(4, 11) = ": " & Format(Now, "dd MMMM yyyy")
        oSheet.Cells(5, 11) = ": Rev-" & rev_MPS
                        
        oSheet.Cells(1, anaGrid.Cols - 3) = "FM-PPC-002-Rev-02"
        oSheet.Cells(2, ttlKolm - 9) = "Disetujui"
        oSheet.Cells(2, ttlKolm - 6) = "Diperiksa"
        oSheet.Cells(2, ttlKolm - 3) = "Dibuat"
               
        oSheet.Cells(2 + 4, ttlKolm - 9) = GetINI("LTPP", "diketahui", vbNullString)
        oSheet.Cells(2 + 4, ttlKolm - 6) = GetINI("LTPP", "diperiksa", vbNullString)
        oSheet.Cells(2 + 4, ttlKolm - 3) = GetINI("LTPP", "dibuat", vbNullString)
        
        ' 8 , 9
        oSheet.Cells(8, 20) = "Total MP"
        
        oSheet.Cells(9, 1) = "Machine " & anaGrid.TextMatrix(2, 0) & " " & getTonage(anaGrid.TextMatrix(2, 0)) & " T (" & emateri & ")"
        oSheet.Cells(10, 1) = "No"
        
        With oSheet
            .Range(.Cells(1, ttlKolm - 3), .Cells(1, ttlKolm - 1)).Merge
            .Cells(1, ttlKolm - 3).HorizontalAlignment = xlRight 'FM-PPC
            
            .Range(.Cells(2, ttlKolm - 9), .Cells(2, ttlKolm - 7)).Merge 'MERGE DISETUJUI

            .Range(.Cells(2, ttlKolm - 6), .Cells(2, ttlKolm - 4)).Merge 'MERGE DIPERIKSA

            .Range(.Cells(2, ttlKolm - 3), .Cells(2, ttlKolm - 1)).Merge
            
            .Range(.Cells(2, ttlKolm - 9), .Cells(2, ttlKolm - 3)).HorizontalAlignment = xlCenter
            
            .Range(.Cells(2 + 1, ttlKolm - 9), .Cells(2 + 3, ttlKolm - 7)).Merge 'MERGE DISETUJUI
            .Range(.Cells(2 + 1, ttlKolm - 6), .Cells(2 + 3, ttlKolm - 4)).Merge 'MERGE DIPERIKSA
            .Range(.Cells(2 + 1, ttlKolm - 3), .Cells(2 + 3, ttlKolm - 1)).Merge 'MERGE DIBUAT
            
            .Range(.Cells(2 + 4, ttlKolm - 9), .Cells(2 + 4, ttlKolm - 7)).Merge
            .Cells(2 + 4, ttlKolm - 9).HorizontalAlignment = xlCenter
            .Range(.Cells(2 + 4, ttlKolm - 6), .Cells(2 + 4, ttlKolm - 4)).Merge
            .Cells(2 + 4, ttlKolm - 6).HorizontalAlignment = xlCenter
            .Range(.Cells(2 + 4, ttlKolm - 3), .Cells(2 + 4, ttlKolm - 1)).Merge
            .Cells(2 + 4, ttlKolm - 3).HorizontalAlignment = xlCenter
            
            .Range(.Cells(2, ttlKolm - 9), .Cells(2 + 4, ttlKolm - 1)).Borders.LineStyle = xlContinuous
            
            'judul
            .Range(.Cells(6, 1), .Cells(6, ttlKolm - 10)).Merge
            .Cells(6, 1) = "Monthly Production Schedule"
            .rows(6).RowHeight = 27
            .Cells(6, 1).Cells.Font.Size = 20
            .Cells(6, 1).Cells.Font.Bold = True
            .Range("A6").HorizontalAlignment = xlCenter
            .Range("A6").VerticalAlignment = xlCenter
            .Range("A6").WrapText = True
        End With
        
        With oSheet
            .PageSetup.Orientation = xlLandscape
            .Range("A10:A" & 10 + 1).Merge
            .Range("A10:A" & 10 + 1).Borders.LineStyle = xlContinuous
            .Range("B10:B" & 10 + 1).Merge
            .Range("C10:C" & 10 + 1).Merge
            .Range("D10:D" & 10 + 1).Merge
            .Range("E10:E" & 10 + 1).Merge
            .Range("F10:F" & 10 + 1).Merge
            .Range("G10:G" & 10 + 1).Merge
            .Range("H10:H" & 10 + 1).Merge
            .Range("I10:I" & 10 + 1).Merge
            .Range("J10:J" & 10 + 1).Merge
            .Range("K10:K" & 10 + 1).Merge
            .Range("L10:L" & 10 + 1).Merge
            .Range("M10:M" & 10 + 1).Merge
            .Range("N10:N" & 10 + 1).Merge
            .Range("P10:P" & 10 + 1).Merge
            .Range("Q10:Q" & 10 + 1).Merge
            .Range("O10:O" & 10 + 1).Merge
            .Range("R10:R" & 10 + 1).Merge
            .Range("S10:S" & 10 + 1).Merge
            .Range("T10:T" & 10 + 1).Merge
            .Range("A10:T" & 10 + 1).HorizontalAlignment = xlCenter
            .Range("A10:T" & 10 + 1).VerticalAlignment = xlCenter
            .Range("A10:T" & 10 + 1).WrapText = True
        End With
        
'        Set cl = Range("D7")  '<-- Range("C2")
'
'        clLeft = cl.Left
'        clTop = cl.Top
'        clHeight = cl.Height
'        clWidth = cl.Width
'
'
'        Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, clTop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
'        shpOval.TextFrame.Characters.Text = "10"
'        shpOval.TextFrame.HorizontalAlignment = xlCenter
'        shpOval.TextFrame.VerticalAlignment = xlCenter
        
        st_ReqHour = 0
        st_ReqDay = 0
        st_OSpo = 0
'        st_PP = 0
        'st_FC = 0
        st_ttlmpp = 0
        st_lc = 0
        
        
        
        With anaGrid
            totalBaris = .rows - 1
            ir = 10
            For r = 0 To totalBaris
                DoEvents
                If LenB(.TextMatrix(r, (COLSDATE - 1))) = 0 And r <= totalBaris - 1 Then '11
                    nourut = 1
                    If LenB(.TextMatrix(r + 1, 0)) <> 0 Then
                        ir = ir + 1
                        oSheet.Cells(ir, 1) = "MACHINE " & .TextMatrix(r + 1, 0) & " " & getTonage(.TextMatrix(r + 1, 0)) & " T  (" & emateri & ")"
                        ir = ir + 1
                        oSheet.Range(oSheet.Cells(10, 1), oSheet.Cells(11, ttlKolm - 1)).Copy oSheet.Range("A" & ir)
                        ir = ir + 2
                    End If
                End If
                .Row = r
                For k = 2 To ttlKolm - 1
                    oSheet.Cells(ir, k) = Trim(.TextMatrix(r, k))
                    oSheet.Cells(ir, k).Borders.LineStyle = xlContinuous
                    .Col = k

                    If .CellBackColor = RGB(255, 155, 155) Or .CellBackColor = RGB(255, 255, 0) Or _
                    .CellBackColor = RGB(111, 255, 0) Or .CellBackColor = RGB(WO_R, WO_G, WO_B) Then
                        oSheet.Cells(ir, k).Interior.Color = .CellBackColor
                    End If
                    If .CellBackColor = RGB(214, 255, 3) And IsNumeric(.Text) Then
                        Set reng = oSheet.Cells(ir, k)
                        If reng.Comment Is Nothing Then reng.AddComment
                        reng.Comment.Text "ada trial " & getDescSch(.TextMatrix(r, 0), Left$(.TextMatrix(1, k), 2))
                    End If
                    If k = 5 Then
                        If Right$(.TextMatrix(r, k), 2) = vbCrLf Or Right$(.TextMatrix(r, k), 2) = vbNewLine Then
                            oSheet.Cells(ir, k) = RTrim(Left$(.TextMatrix(r, k), Len(.TextMatrix(r, k)) - 2))
                        Else
                            oSheet.Cells(ir, k) = RTrim(.TextMatrix(r, k))
                        End If
                    End If
                    If k = 2 Then
                        oSheet.Cells(ir, k) = RTrim(.TextMatrix(r, 1))
                    End If
                    'warna Revisi

                    If .CellBackColor = RGB(255, 85, 127) Then
                        Set cl = Range(oSheet.Cells(ir, k), oSheet.Cells(ir, k))
                        clLeft = cl.Left
                        cltop = cl.Top
                        clHeight = cl.Height
                        clWidth = cl.Width / 2.6
                        'Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)    ', clLeft, clTop, 4, 10
                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)    ', clLeft, clTop, 4, 10
                        shpOval.TextFrame.Characters.Text = "1"
                        shpOval.TextFrame.Characters.Font.Color = 1
                        shpOval.TextFrame.Characters.Font.Size = 9
                        shpOval.TextFrame.HorizontalAlignment = xlCenter
                        shpOval.TextFrame.VerticalAlignment = xlCenter
                        shpOval.Fill.ForeColor.RGB = RGB(255, 85, 127)
                    ElseIf .CellBackColor = RGB(255, 255, 1) Then ' kuning
                        Set cl = Range(oSheet.Cells(ir, k), oSheet.Cells(ir, k))
                        clLeft = cl.Left
                        cltop = cl.Top
                        clHeight = cl.Height
                        clWidth = cl.Width / 2.6
                        'Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
                        shpOval.TextFrame.Characters.Text = "2"
                        shpOval.TextFrame.Characters.Font.Color = 1
                        shpOval.TextFrame.Characters.Font.Size = 9
                        shpOval.TextFrame.HorizontalAlignment = xlCenter
                        shpOval.TextFrame.VerticalAlignment = xlCenter
                        shpOval.Fill.ForeColor.RGB = RGB(255, 255, 1)
                    ElseIf .CellBackColor = RGB(127, 212, 255) Then
                        Set cl = Range(oSheet.Cells(ir, k), oSheet.Cells(ir, k))
                        clLeft = cl.Left
                        cltop = cl.Top
                        clHeight = cl.Height
                        clWidth = cl.Width / 2.6
                        'Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
                        shpOval.TextFrame.Characters.Text = "3"
                        shpOval.TextFrame.Characters.Font.Color = 1
                        shpOval.TextFrame.Characters.Font.Size = 9
                        shpOval.TextFrame.HorizontalAlignment = xlCenter
                        shpOval.TextFrame.VerticalAlignment = xlCenter
                        shpOval.Fill.ForeColor.RGB = RGB(127, 212, 255)
                    ElseIf .CellBackColor = RGB(255, 127, 212) Then
                        Set cl = Range(oSheet.Cells(ir, k), oSheet.Cells(ir, k))
                        clLeft = cl.Left
                        cltop = cl.Top
                        clHeight = cl.Height
                        clWidth = cl.Width / 2.6
                        'Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
                        shpOval.TextFrame.Characters.Text = "4"
                        shpOval.TextFrame.Characters.Font.Color = 1
                        shpOval.TextFrame.Characters.Font.Size = 9
                        shpOval.TextFrame.HorizontalAlignment = xlCenter
                        shpOval.TextFrame.VerticalAlignment = xlCenter
                        shpOval.Fill.ForeColor.RGB = RGB(255, 127, 212) ' pink
                    ElseIf .CellBackColor = RGB(0, 255, 0) Then
                        Set cl = Range(oSheet.Cells(ir, k), oSheet.Cells(ir, k))
                        clLeft = cl.Left
                        cltop = cl.Top
                        clHeight = cl.Height
                        clWidth = cl.Width / 2.6
                        'Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
                        shpOval.TextFrame.Characters.Text = "5"
                        shpOval.TextFrame.Characters.Font.Color = 1
                        shpOval.TextFrame.Characters.Font.Size = 9
                        shpOval.TextFrame.HorizontalAlignment = xlCenter
                        shpOval.TextFrame.VerticalAlignment = xlCenter
                        shpOval.Fill.ForeColor.RGB = RGB(0, 255, 0) ' green
                    ElseIf .CellBackColor = RGB(212, 170, 255) Then
                        Set cl = Range(oSheet.Cells(ir, k), oSheet.Cells(ir, k))
                        clLeft = cl.Left
                        cltop = cl.Top
                        clHeight = cl.Height
                        clWidth = cl.Width / 2.6
                        'Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
                        shpOval.TextFrame.Characters.Text = "6"
                        shpOval.TextFrame.Characters.Font.Color = 1
                        shpOval.TextFrame.Characters.Font.Size = 9
                        shpOval.TextFrame.HorizontalAlignment = xlCenter
                        shpOval.TextFrame.VerticalAlignment = xlCenter
                        shpOval.Fill.ForeColor.RGB = RGB(212, 170, 255) ' ungu
                    End If
                Next
                If r > 2 And r <= totalBaris Then
                    oSheet.Cells(ir, 1).Borders.LineStyle = xlContinuous
                End If
                If r > 1 Then
                    oSheet.Cells(ir, 1) = .TextMatrix(r, 2)
                End If
                ir = ir + 1
                PG1.Value = FormatNumber((r * 100) / totalBaris, 0)
                PG1.ToolTipText = PG1.Value & "%"

                If .TextMatrix(r, 1) = "Total" Then
                    st_ReqHour = .TextMatrix(r, 13) * 1 + st_ReqHour '12
                    st_ReqDay = .TextMatrix(r, 14) * 1 + st_ReqDay '13
                    If IsNumeric(.TextMatrix(r, 15)) Then '14
                        st_OSpo = .TextMatrix(r, 15) * 1 + st_OSpo '14
                    Else
                        st_OSpo = 0 + st_OSpo
                    End If

                    st_ttlmpp = st_ttlmpp + .TextMatrix(r, 20)
                    If LenB(.TextMatrix(r, (COLSDATE - 1))) > 0 Then
                        st_lc = st_lc + Left$(.TextMatrix(r, (COLSDATE - 1)), Len(.TextMatrix(r, 20)) - 1) * 1
                    End If
                End If
            Next
        End With

        '===========START========Total Man Power=============
        With oSheet
            .Cells(8, 21) = "=sumif($T$11:$T$" & ir - 1 & ",""MP"",U11:U" & ir - 1 & ")"
            Set cl = .Range(.Cells(8, 21), .Cells(8, 21))
            cl.Copy .Range(.Cells(8, 22), .Cells(8, ttlKolm - 1))
            .Range(.Cells(8, 20), .Cells(8, ttlKolm - 1)).Borders.LineStyle = xlContinuous
        End With
        '============END=========Total Man Power=============

        'total inj
        ir = ir + 1
        oSheet.Cells(ir, 12) = "Total" '11
        oSheet.Cells(ir, 13) = st_ReqHour '12
        oSheet.Cells(ir, 14) = st_ReqDay '13
        oSheet.Cells(ir, 15) = st_OSpo '14
        oSheet.Cells(ir, 16) = getTotalPP_ltpp ' 15
        oSheet.Cells(ir, 19) = getTotalfc ' st_FC '18
        oSheet.Cells(ir, 20) = st_ttlmpp '19
        oSheet.Cells(ir, 21) = st_lc '20
        oSheet.Range("L" & ir & ":T" & ir).Interior.Color = RGB(255, 248, 11) ' kuning

        ir = ir + 9

'        st_ReqHour = 0
'        st_ReqDay = 0
'        st_OSpo = 0
'        st_PP = 0
'        st_FC = 0
'        st_ttlmpp = 0
'        st_lc = 0
        
'        oSheet.Cells(ir, 1) = "MACHINE " & anaSubcont.TextMatrix(2, 0) & " " & getDescSubcont(anaSubcont.TextMatrix(2, 0))
'        ir = ir + 1
'        oSheet.Range(oSheet.Cells(10, 1), oSheet.Cells(11, ttlKolm - 1)).Copy oSheet.Range("A" & ir)
'        ir = ir + 3
'        With anaSubcont
'            totalBaris = .rows - 1
'            For r = 2 To .rows - 1
'                If LenB(.TextMatrix(r, 12)) = 0 And r <= totalBaris - 1 Then
'                    nourut = 1
'                    If LenB(.TextMatrix(r + 1, 0)) <> 0 Then
'                        ir = ir + 1
'                        oSheet.Cells(ir, 1) = "MACHINE " & .TextMatrix(r + 1, 0) & " " & getDescSubcont(.TextMatrix(r + 1, 0))
'                        ir = ir + 1
'                        oSheet.Range(oSheet.Cells(10, 1), oSheet.Cells(11, ttlKolm - 1)).Copy oSheet.Range("A" & ir)
'                        ir = ir + 2
'                    End If
'                End If
'                .Row = r
'                For k = 2 To .Cols - 1
'                    .Col = k
'                    If k < .Cols - 1 Then
'                        oSheet.Cells(ir, k).Borders.LineStyle = xlContinuous
'                    End If
'                    If k > 19 Then
'                        oSheet.Cells(ir, k - 1) = .TextMatrix(r, k)
'                    End If
'                    If .CellBackColor = RGB(255, 155, 155) Or .CellBackColor = RGB(255, 255, 0) Or _
'                    .CellBackColor = RGB(111, 255, 0) Or .CellBackColor = RGB(WO_R, WO_G, WO_B) Then
'                        oSheet.Cells(ir, k - 1).Interior.Color = .CellBackColor
'                    End If
'                    If .CellBackColor = RGB(255, 85, 127) Then
'                        Set cl = Range(oSheet.Cells(ir, k - 1), oSheet.Cells(ir, k - 1))
'                        clLeft = cl.Left
'                        cltop = cl.Top
'                        clHeight = cl.Height
'                        clWidth = cl.Width / 2.6
'                        Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)    ', clLeft, clTop, 4, 10
'                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
'                        shpOval.TextFrame.Characters.Text = "1"
'                        shpOval.TextFrame.Characters.Font.Color = 1
'                        shpOval.TextFrame.Characters.Font.Size = 9
'                        shpOval.TextFrame.HorizontalAlignment = xlCenter
'                        shpOval.TextFrame.VerticalAlignment = xlCenter
'                        shpOval.Fill.ForeColor.RGB = RGB(255, 85, 127)
'                    ElseIf .CellBackColor = RGB(255, 255, 1) Then ' kuning
'                        Set cl = Range(oSheet.Cells(ir, k - 1), oSheet.Cells(ir, k - 1))
'                        clLeft = cl.Left
'                        cltop = cl.Top
'                        clHeight = cl.Height
'                        clWidth = cl.Width / 2.6
'                        Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
'                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
'                        shpOval.TextFrame.Characters.Text = "2"
'                        shpOval.TextFrame.Characters.Font.Color = 1
'                        shpOval.TextFrame.Characters.Font.Size = 9
'                        shpOval.TextFrame.HorizontalAlignment = xlCenter
'                        shpOval.TextFrame.VerticalAlignment = xlCenter
'                        shpOval.Fill.ForeColor.RGB = RGB(255, 255, 1)
'                    ElseIf .CellBackColor = RGB(127, 212, 255) Then
'                        Set cl = Range(oSheet.Cells(ir, k - 1), oSheet.Cells(ir, k - 1))
'                        clLeft = cl.Left
'                        cltop = cl.Top
'                        clHeight = cl.Height
'                        clWidth = cl.Width / 2.6
'                        Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
'                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
'                        shpOval.TextFrame.Characters.Text = "3"
'                        shpOval.TextFrame.Characters.Font.Color = 1
'                        shpOval.TextFrame.Characters.Font.Size = 9
'                        shpOval.TextFrame.HorizontalAlignment = xlCenter
'                        shpOval.TextFrame.VerticalAlignment = xlCenter
'                        shpOval.Fill.ForeColor.RGB = RGB(127, 212, 255)
'                    ElseIf .CellBackColor = RGB(255, 127, 212) Then
'                        Set cl = Range(oSheet.Cells(ir, k - 1), oSheet.Cells(ir, k - 1))
'                        clLeft = cl.Left
'                        cltop = cl.Top
'                        clHeight = cl.Height
'                        clWidth = cl.Width / 2.6
'                        Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
'                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
'                        shpOval.TextFrame.Characters.Text = "4"
'                        shpOval.TextFrame.Characters.Font.Color = 1
'                        shpOval.TextFrame.Characters.Font.Size = 9
'                        shpOval.TextFrame.HorizontalAlignment = xlCenter
'                        shpOval.TextFrame.VerticalAlignment = xlCenter
'                        shpOval.Fill.ForeColor.RGB = RGB(255, 127, 212) ' pink
'                    ElseIf .CellBackColor = RGB(0, 255, 0) Then
'                        Set cl = Range(oSheet.Cells(ir, k - 1), oSheet.Cells(ir, k - 1))
'                        clLeft = cl.Left
'                        cltop = cl.Top
'                        clHeight = cl.Height
'                        clWidth = cl.Width / 2.6
'                        Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
'                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
'                        shpOval.TextFrame.Characters.Text = "5"
'                        shpOval.TextFrame.Characters.Font.Color = 1
'                        shpOval.TextFrame.Characters.Font.Size = 9
'                        shpOval.TextFrame.HorizontalAlignment = xlCenter
'                        shpOval.TextFrame.VerticalAlignment = xlCenter
'                        shpOval.Fill.ForeColor.RGB = RGB(0, 255, 0) ' green
'                    ElseIf .CellBackColor = RGB(212, 170, 255) Then
'                        Set cl = Range(oSheet.Cells(ir, k - 1), oSheet.Cells(ir, k - 1))
'                        clLeft = cl.Left
'                        cltop = cl.Top
'                        clHeight = cl.Height
'                        clWidth = cl.Width / 2.6
'                        Set shpOval = oSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, clLeft, cltop, clWidth, clHeight)   ', clLeft, clTop, 4, 10
'                        'Set shpOval = oSheet.Shapes.AddShape(7, clLeft, clTop, clWidth, clHeight)
'                        shpOval.TextFrame.Characters.Text = "6"
'                        shpOval.TextFrame.Characters.Font.Color = 1
'                        shpOval.TextFrame.Characters.Font.Size = 9
'                        shpOval.TextFrame.HorizontalAlignment = xlCenter
'                        shpOval.TextFrame.VerticalAlignment = xlCenter
'                        shpOval.Fill.ForeColor.RGB = RGB(212, 170, 255) ' ungu
'                    End If
'                Next
''                .Col = 19
''                If .CellBackColor = RGB(255, 155, 155) Or .CellBackColor = RGB(255, 255, 0) Or .CellBackColor = RGB(111, 255, 0) Then
''                    oSheet.Cells(ir, 18).Interior.Color = .CellBackColor
''                End If
'                If LenB(.TextMatrix(r, 0)) <> 0 Then
'                    oSheet.Cells(ir, 1) = nourut
'                End If
'                oSheet.Cells(ir, 2) = .TextMatrix(r, 1)
'                oSheet.Cells(ir, 3) = .TextMatrix(r, 2)
'                oSheet.Cells(ir, 4) = .TextMatrix(r, 3)
'                If Right$(.TextMatrix(r, 6), 2) = vbCrLf Or Right$(.TextMatrix(r, 6), 2) = vbNewLine Then
'                    oSheet.Cells(ir, 5) = RTrim$(Left$(.TextMatrix(r, 6), Len(.TextMatrix(r, 6)) - 2))
'                Else
'                    oSheet.Cells(ir, 5) = RTrim$(.TextMatrix(r, 6))
'                End If
'                oSheet.Cells(ir, 6) = .TextMatrix(r, 7)
'                oSheet.Cells(ir, 7) = .TextMatrix(r, 8)
'                oSheet.Cells(ir, 8) = .TextMatrix(r, 9)
'                oSheet.Cells(ir, 9) = .TextMatrix(r, 10)
'                oSheet.Cells(ir, 10) = .TextMatrix(r, 11)
'                oSheet.Cells(ir, 11) = .TextMatrix(r, 12)
'                oSheet.Cells(ir, 12) = .TextMatrix(r, 13)
'                oSheet.Cells(ir, 13) = .TextMatrix(r, 14)
'                oSheet.Cells(ir, 14) = .TextMatrix(r, 15)
'                oSheet.Cells(ir, 15) = .TextMatrix(r, 16)
'                oSheet.Cells(ir, 16) = .TextMatrix(r, 17)
'                oSheet.Cells(ir, 17) = .TextMatrix(r, 18)
'                oSheet.Cells(ir, 18) = .TextMatrix(r, 19)
'                If r > 2 And r <= totalBaris Then
'                    oSheet.Cells(ir, 1).Borders.LineStyle = xlContinuous
'                End If
'
'                ir = ir + 1
'
'                PG1.Value = FormatNumber((r * 100) / totalBaris, 0)
'                PG1.ToolTipText = PG1.Value & "%"
'                If LenB(.TextMatrix(r, 0)) <> 0 Then
'                    nourut = nourut + 1
'                End If
'                If .TextMatrix(r, 1) = "Total" Then
'                    st_ReqHour = .TextMatrix(r, 13) * 1 + st_ReqHour
'                    st_ReqDay = .TextMatrix(r, 14) * 1 + st_ReqDay
'                    If IsNumeric(.TextMatrix(r, 15)) Then
'                        st_OSpo = .TextMatrix(r, 15) * 1 + st_OSpo
'                    Else
'                        st_OSpo = 0 + st_OSpo
'                    End If
'                    st_PP = .TextMatrix(r, 16) * 1 + st_PP
'                    If IsNumeric(.TextMatrix(r, 17)) Then
'                        st_FC = .TextMatrix(r, 17) * 1 + st_FC
'                    Else
'                        st_FC = 0 + st_FC
'                    End If
'                    st_ttlmpp = st_ttlmpp + .TextMatrix(r, 18)
'                    If Len(.TextMatrix(r, 19)) > 0 Then
'                        st_lc = st_lc + Left$(.TextMatrix(r, 19), Len(.TextMatrix(r, 19)) - 1) * 1
'                    End If
'                End If
'            Next
'            oSheet.Columns("B:G").AutoFit
'        End With
'        ir = ir + 1
'        oSheet.Cells(ir, 11) = "Total"
'        oSheet.Cells(ir, 12) = st_ReqHour
'        oSheet.Cells(ir, 13) = st_ReqDay
'        oSheet.Cells(ir, 14) = st_OSpo
'        oSheet.Cells(ir, 15) = st_PP
'        oSheet.Cells(ir, 16) = st_FC
'        oSheet.Cells(ir, 17) = st_ttlmpp
'        oSheet.Cells(ir, 18) = st_lc
'        oSheet.Range("L" & ir & ":R" & ir).Interior.Color = RGB(255, 248, 11) ' kuning
    
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        If MsgBox("open the file ", vbQuestion + vbYesNo, "Tentukan") = vbYes Then
            oExcel.Visible = True
        Else
            oExcel.Quit
        End If
        
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
        PG1.Visible = False
        
    End If
    Exit Sub
Nah:
    oExcel.Quit
    Set oSheet = Nothing
    Set oBook = Nothing
    Set oExcel = Nothing
    PG1.Visible = False
    MsgBox Err.Description, vbInformation, "Maaf"
End Sub

Private Function getDescSch(pmesin As String, pTGL As String) As String
    rsMCtrial.Fields("mch").Properties("Optimize") = True
    rsMCtrial.Fields("tgl").Properties("Optimize") = True
    rsMCtrial.Filter = adFilterNone
    rsMCtrial.Filter = "mch='" & pmesin & "' and tgl=" & pTGL
    If rsMCtrial.RecordCount > 0 Then
        getDescSch = "untuk " & rsMCtrial("part_no") & ", dari " & rsMCtrial("dari") & " sampai " & rsMCtrial("sampai")
    End If
End Function

Private Function getTrialPart(pmesin As String, pTGL As String) As String
    rsMCtrial.Filter = adFilterNone
    rsMCtrial.Filter = "mch='" & pmesin & "' and tgl=" & pTGL
    If rsMCtrial.RecordCount > 0 Then
        getTrialPart = rsMCtrial("part_no")
    End If
End Function

Private Function getTrialTime(pmesin As String, pTGL As String) As String
    rsMCtrial.Filter = adFilterNone
    rsMCtrial.Filter = "mch='" & pmesin & "' and tgl=" & pTGL
    If rsMCtrial.RecordCount > 0 Then
        getTrialTime = rsMCtrial("dari") & " to " & rsMCtrial("sampai")
    End If
End Function



Private Function bilbulat(parmch As Variant, partgl As Long) As Boolean
    Dim Tsday As Single
    Tsday = 0
    rsWO.Filter = "mesinno='" & parmch & "' and hari='" & partgl & "'"
    If rsWO.RecordCount > 0 Then
        If rsWO.RecordCount = 1 Then
            If rsWO("dayrun") = 1 Then
                bilbulat = True
            Else
                bilbulat = False
            End If
        Else
            While Not rsWO.EOF
                Tsday = Tsday + rsWO("dayrun")
                rsWO.MoveNext
            Wend
            If Round(Tsday) = 1 Then
                bilbulat = True
            Else
                bilbulat = False
            End If
        End If
    Else
        bilbulat = False
    End If
End Function

Private Function bilbulat2(parmch As Variant, partgl As Long) As Single
    rsWO.Filter = "mesinno='" & parmch & "' and hari='" & partgl & "' and dayrun<>1"
    If rsWO.RecordCount > 0 Then
        If rsWO.RecordCount = 1 Then
            
            With anaGrid
            For r = 1 To UBound(aPart)
                If Trim(rsWO("partno")) = aPart(r, 1) And rsWO("mesinno") = aPart(r, 10) And rsWO("moldno") = aPart(r, 7) Then
                    aPart(r, 11 + rsWO("hari")) = 1
                    Exit For
                End If
            Next
            End With
            bilbulat2 = rsWO("dayrun")
        Else
            bilbulat2 = 0
        End If
    Else
        bilbulat2 = 0
    End If
End Function

Private Sub syncNeQtywithWO()
    While Not rsWO.EOF
        For i = 1 To UBound(aPart)
            If aPart(i, 1) = rsWO("partno") And aPart(i, 7) = rsWO("moldno") And aPart(i, 10) = rsWO("mesinno") Then
                aPart(i, 11) = aPart(i, 11) * 1 - rsWO("qty")
            End If
        Next
        rsWO.MoveNext
    Wend
End Sub

Private Sub lMaterialPurg()
    qry = "select * from v_matpurging"
    Set rsPurg = Con.Execute(qry)
End Sub

Private Sub cmdGenerate_Click()
    If Len(txtRevision) < 1 Then txtRevision.SetFocus: Exit Sub
    cmbPeriod.ListIndex = 0
    '----suggestion overload
    loadSuggestion
        
    If (Not aSuggest) <> -1 Then picSign_ovr.SetFocus: Exit Sub
    
    Dim currentDate As Byte
    If Right(CmbDocument, 7) = Format(Now, "MM\/yyyy") Then
        currentDate = Val(Format(Now, "dd"))
    Else
        currentDate = 1
    End If
    
    
    Screen.MousePointer = 11
    cmbType.ListIndex = 0
    belumSimpan = True
    NoDocMPS = ""
    bhulan = DateSerial(Left$(cmbPeriod, 4), Right$(cmbPeriod, 2) * 1, 1)
    totalHari = Int(Format(dhLastDayInMonth(bhulan), "dd"))
    
    DAYtoGrid anaGrid, cmbPeriod
    DAYtoGrid anaSubcont, cmbPeriod
    DAYtoGrid anaUnproc, cmbPeriod
    DAYtoGrid anaAssy, cmbPeriod
    lblstate.Visible = True
    lblstate.Caption = "Preparing data : 'day off' "
    Me.Refresh
    '^day off
    qry = "select extract(day from work_date) harioff from plansys_setoffday " _
        & " where extract(month from work_date)=" & Right$(cmbPeriod, 2) & " and extract(year from work_date)=" & Left$(cmbPeriod, 4) & " and work_status=false order by 1"
    Set RsA = Con.Execute(qry)
    If RsA.RecordCount > 0 Then
        Erase aDayOFF
        ReDim aDayOFF(1 To RsA.RecordCount) As Variant
        i = 1
        While Not RsA.EOF
            aDayOFF(i) = RsA(0)
            i = 1 + i
            RsA.MoveNext
        Wend
    End If
    loadFWO
    
    lblstate.Caption = "Preparing data : 'overtime'"
    Me.Refresh
    
    '^day ovr
    qry = "select extract(day from wrk_date) hariovr,no_mach from mpp_setovrtime " _
        & " where extract(year from wrk_date)=" & Left$(cmbPeriod, 4) & " and extract(month from wrk_date)=" & Right$(cmbPeriod, 2)
    Set RsA = Con.Execute(qry)
    If RsA.RecordCount > 0 Then
        Erase aDayOvr
        ReDim aDayOvr(1 To RsA.RecordCount, 1 To 2) As Variant
        i = 1
        While Not RsA.EOF
            aDayOvr(i, 1) = RsA(1)
            aDayOvr(i, 2) = RsA(0)
            i = 1 + i
            RsA.MoveNext
        Wend
    End If
    
    lblstate.Caption = "Preparing data : 'delivery schedule'"
    Me.Refresh
    
    qry = "select extract(day from delv_date),part_no,sum(qty) qty from mpp_delv_plan where  " _
        & " extract(month from delv_date)=" & Right$(cmbPeriod, 2) & " and extract(year from delv_date)=" & Left$(cmbPeriod, 4) _
        & " group by part_no , delv_date order by 1 asc"
    Set RsA = Con.Execute(qry)
    
    If RsA.RecordCount > 0 Then
        Erase aJadwal
        ReDim aJadwal(1 To RsA.RecordCount, 1 To 3)
        i = 1
        While Not RsA.EOF
            aJadwal(i, 1) = RsA(0)
            aJadwal(i, 2) = RsA(1)
            aJadwal(i, 3) = RsA(2)
            i = i + 1
            RsA.MoveNext
        Wend
    Else
        Erase aJadwal
        ReDim aJadwal(1 To 1, 1 To 3) As Variant
    End If
    
    lblstate.Caption = "Preparing data : 'master product'"
    Me.Refresh
    '# INJ
    qry = "select  a.no_mach, ton_mach ,lc_customer,  lcd_itemdid, lc_itemname, cav, ct, cap_p_day,fltpp_hkw,  neday, reg_mold, lc_pp, lc_subcont, lcneed_mp need_mp, hourpshift,neqty  " _
        & ",item_muloq,item_perbox,coalesce(ith_qty,0) stok,coalesce(minstock,3) minstock,maxstock,cav_std,ct_scnd  " _
        & ",lcneed_mp,isno,typelabel,lc_fc,shiftusg,mpower,lc_sisa_pp,prod_plan_1 " _
        & ",coalesce(colordesc,'-') colordesc from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
        & " left join mst_item b on a.lcd_itemdid=b.item_id " _
        & " left join loadcap_mst_product_r c on a.lcd_itemdid=c.partno " _
        & " left join ltpp_generate d on a.lcd_itemdid=d.assy_no and a.fltpp_doc=d.ltpp_doc" _
        & " where stscode_id='01' AND fltpp_doc='" & CmbDocument & "' and lc_subcont='no' " _
        & " and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' " _
        & " order by a.no_mach asc, lc_customer asc, lcd_itemdid asc"

    Set RsA = Con.Execute(qry)
    Set rsA_aks = Con.Execute(qry)
    RsA.Fields("no_mach").Properties("Optimize") = True
    anaGrid.rows = 2
    HKWs = RsA("fltpp_hkw")
    If RsA.RecordCount > 0 Then
        Erase aPart
        ReDim aPart(1 To RsA.RecordCount, 1 To 11 + totalHari + 1 + 9) As Variant '9 kolom baru
        Erase in_PartOL
        Erase in_PartOLQTY
        Erase in_partLTPP
        Erase in_partFC
        i = 1
        While Not RsA.EOF
            aPart(i, 1) = RsA("lcd_itemdid").Value: aPart(i, 2) = RsA("stok").Value: aPart(i, 3) = RsA("stok").Value
            aPart(i, 4) = RsA("item_muloq").Value: aPart(i, 5) = RsA("item_perbox").Value: aPart(i, 6) = RsA("lc_pp").Value
            aPart(i, 7) = RsA("reg_mold"): aPart(i, 8) = RsA("minstock"): aPart(i, 9) = RsA("maxstock")
            aPart(i, 10) = RsA("no_mach"): aPart(i, 11) = RsA("neqty")
            
            For k = 12 To totalHari + 1
                aPart(i, k) = "0"
            Next
            inAddValue RsA("lcd_itemdid"), RsA("lc_sisa_pp"), RsA("prod_plan_1"), RsA("lc_fc")
            i = 1 + i
            RsA.MoveNext
        Wend
        syncNeQtywithWO
        ' / dapat mesin distinct
        qry = "select wr.no_mach,material,ton_mach from " _
         & "(select distinct on (no_mach)  no_mach,ton_mach from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
         & " left join mst_item b on a.lcd_itemdid=b.item_id " _
         & " where stscode_id='01' AND fltpp_doc='" & CmbDocument & "' and lc_subcont='no' " _
         & " and fltpp_rev=" & txtRevision & "  and fltpp_ym='" & cmbPeriod & "' " _
         & " order by no_mach asc, lc_customer asc, lcd_itemdid asc) wr left join " _
         & " v_mesin_mater vmm on wr.no_mach=vmm.no_mach"
         'and neday>0
        Set rsB = Con.Execute(qry)

        ReDim aMesinInj(1 To rsB.RecordCount, 1 To totalHari + 3) As Variant
        ReDim aMesinInj_r(1 To rsB.RecordCount, 1 To totalHari + 1) As Variant
        i = 1
        While Not rsB.EOF
            aMesinInj(i, 1) = rsB(0)
            aMesinInj(i, 2) = IIf(IsNull(rsB(1)), " ", rsB(1))
            aMesinInj(i, 3) = IIf(IsNull(rsB(2)), " ", rsB(2))
            aMesinInj_r(i, 1) = rsB(0)
            
            For k = 4 To totalHari + 3
                If bilbulat(aMesinInj(i, 1), k - 3) Then
                    aMesinInj(i, k) = "1"
                Else
                    aMesinInj(i, k) = "0"
                End If
            Next
            
            For k = 2 To totalHari + 1
                aMesinInj_r(i, k) = 1 - bilbulat2(aMesinInj_r(i, 1), k - 1)
            Next
            
            i = 1 + i
            rsB.MoveNext
        Wend
        anaGrid.rows = anaGrid.rows + 1
        anaGrid.FixedRows = 2
        SkinLabel4.Caption = "HKW : " & HKWs

        i = 2
        RsA.MoveFirst
        With anaGrid
            .Col = 5
            While Not RsA.EOF
                If temp_mch = RsA("no_mach") Then
                    noItemPerMesin = noItemPerMesin + 1
                Else
                    noItemPerMesin = 1
                End If
                 .TextMatrix(i, 0) = RsA("no_mach")
                 .TextMatrix(i, 1) = RsA("lc_customer")
                 .TextMatrix(i, 2) = noItemPerMesin
                 .TextMatrix(i, 3) = RsA("lcd_itemdid")
                 .TextMatrix(i, 4) = RsA("lc_itemname")
                     .Row = i
                 .CellAlignment = flexAlignLeftCenter
                 .TextMatrix(i, 5) = RsA("reg_mold")
                 .TextMatrix(i, 6) = RsA("colordesc")
                 .TextMatrix(i, 7) = RsA("cav")
                 .TextMatrix(i, 8) = RsA("ct")
                 .TextMatrix(i, 9) = RsA("need_mp")
                If RsA("ct") > 0 Then
                    .TextMatrix(i, 10) = FormatNumber((3600 / RsA("ct")) * RsA("cav"), 0) ' capacity per hour
                    
                    .TextMatrix(i, 11) = (.TextMatrix(i, 10) * 1) * RsA("hourpshift") ' capacity per shift
                    .TextMatrix(i, 11) = FormatNumber(.TextMatrix(i, 11), 0)
                Else
                    .TextMatrix(i, 10) = 0
                    .TextMatrix(i, 11) = 0
                End If
                .TextMatrix(i, 12) = FormatNumber(RsA("cap_p_day"), 0)
                .TextMatrix(i, 16) = FormatNumber(RsA("prod_plan_1"), 0) 'FormatNumber(rsA("lc_pp"), 0)
                If RsA("neqty") > 0 Then
                    If RsA("item_perbox") = 0 Then
                        .TextMatrix(i, 17) = FormatNumber(isi(RsA("item_muloq"), CLng(RsA("neqty")), "a"), 0)
                    Else
                        .TextMatrix(i, 17) = FormatNumber(isi(RsA("item_perbox"), CLng(RsA("neqty")), "a"), 0)
                    End If
                Else
                    .TextMatrix(i, 17) = 0
                End If
                
                .TextMatrix(i, 18) = 0 'FormatNumber(rsA("lc_fc"), 0)
                .TextMatrix(i, 19) = FormatNumber(RsA("lc_fc"), 0)
                RsA.MoveNext
                If RsA.EOF Then

                Else
                    temp_mch = RsA("no_mach")
                End If
                RsA.MovePrevious
                If temp_mch = RsA("no_mach") Then
                    .rows = .rows + 1
                    i = i + 1
                Else
                    .rows = .rows + 4
                    i = i + 4
                End If
                
                temp_mch = RsA("no_mach")
                RsA.MoveNext
            Wend
            .rows = .rows + 1
            For r = 2 To .rows - 1
                .TextMatrix(r, 18) = FormatNumber(inGetValue(.TextMatrix(r, 3)), 0)
            Next
        End With
        
        RsA.AbsolutePosition = 1
        rsWO.Filter = adFilterNone
        If rsWO.RecordCount > 0 Then
            
            With anaGrid
                For i = 2 To .rows - 1
                    .Row = i
                    If LenB(.TextMatrix(i, 0)) <> 0 Then
                        qry = "mesinno='" & .TextMatrix(i, 0) & "' and " _
                        & " partno='" & .TextMatrix(i, 3) & "' and moldno='" & .TextMatrix(i, 5) & "'"
                        rsWO.Filter = qry
                        If rsWO.RecordCount > 0 Then
                            For k = COLSDATE To .Cols - 1
                                qry = "mesinno='" & .TextMatrix(i, 0) & "' and " _
                                & " partno='" & .TextMatrix(i, 3) & "' and moldno='" & .TextMatrix(i, 5) & "' and hari='" & Left$(.TextMatrix(1, k), 2) * 1 & "'"
                                rsWO.Filter = qry
                                If rsWO.RecordCount > 0 Then
    
                                    .TextMatrix(i, k) = rsWO("qty")
                                    
                                    .Col = k
                                    .CellBackColor = RGB(WO_R, WO_G, WO_B)
                                    mc_dlm_Tgl_byname rsWO("hari"), rsWO("mesinno")
                                End If
                            Next
                        End If
                    End If
                Next
            End With
        End If
        'tandai mesin trial

        'Call plotMchTrial(anagrid)
        qry = "select distinct on(mch,date_trial) mch, date_trial::date,extract(day from date_trial) tgl,date_trial dari,date_trialf sampai,part_no from " _
            & " (select part_no,mch,date_trial,date_trialf,extract(epoch  from (date_trialf-date_trial))/60/60/24 hari from mpp_ste " _
            & " where extract(MONTH from date_trial)=" & Right$(cmbPeriod, 2) & " and extract(YEAR from date_trial)=" & Left$(cmbPeriod, 4) & ")  dd " _
            & " order by mch asc"
        Set rsMCtrial = Con.Execute(qry)
        rsMCtrial.Fields("mch").Properties("Optimize") = True
        rsMCtrial.Fields("tgl").Properties("Optimize") = True
        Dim ttlKol As Byte
        Dim ttlBar As Long
        ttlKol = anaGrid.Cols - 1
        ttlBar = anaGrid.rows - 1
        For k = 1 To rsMCtrial.RecordCount
            rsMCtrial.AbsolutePosition = k
            For i = 2 To ttlBar
                For r = COLSDATE To ttlKol '18
                    If anaGrid.TextMatrix(i, 0) = rsMCtrial(0) And Left$(anaGrid.TextMatrix(1, r), 2) = Format(rsMCtrial(1), "dd") Then
                        anaGrid.Col = r
                        anaGrid.Row = i
                        anaGrid.CellBackColor = RGB(214, 255, 3)
                    End If
                Next
            Next
        Next
        '--------,,,,,-------
        
        For tanggal = currentDate To totalHari
            lblstate.Caption = "Plotting dates : " & tanggal
            Me.Refresh
            qry = "select part_no,sum(qty) qty from mpp_delv_plan where delv_date='" & Left$(cmbPeriod, 4) & "-" & Right$(cmbPeriod, 2) & "-" & tanggal & "' " _
                & " group by part_no "
            Set rsB = Con.Execute(qry)

            If rsB.RecordCount > 0 Then
                Erase aDelv
                i = 1
                ReDim aDelv(1 To rsB.RecordCount, 1 To 2) As Variant
                While Not rsB.EOF
                    aDelv(i, 1) = rsB(0): aDelv(i, 2) = rsB(1)
                    i = i + 1
                    rsB.MoveNext
                Wend
                '~ pengambilan stok - delivery
                For r = 1 To UBound(aPart)
                    For i = 1 To UBound(aDelv)
                        If aPart(r, 1) = aDelv(i, 1) Then
                            aPart(r, 3) = (aPart(r, 3) * 1) - (aDelv(i, 2) * 1)
                        End If
                    Next
                Next

                '~ saatnya mulai loop
                plot_v2 tanggal
                addStokFor2morrow tanggal
            Else
                plot_v2 tanggal
                addStokFor2morrow tanggal
            End If
        Next

        '** hitung total
        Call hitungRekapINJ
       
    End If
    
    cmdSave.Visible = True
    
    
    '# subcont
    qry = "select max(no_mach) mesin, min(lc_customer) lc_customer, a.lcd_itemdid, min(lc_itemname) lc_itemname, max(lc_pp) lc_pp, reg_mold ,min(minton_mch) minton,max(maxton_mch) maxton," _
    & " max(lcneed_mp) need_mp,min(hourpshift) hourpshift,max(cav) cav, min(ct) ct, max(cap_p_day) capday,min(item_perbox) item_perbox,min(item_muloq) item_muloq" _
    & " ,coalesce(ith_qty,0) stok,max(minstock) minstock,max(maxstock) maxstock," _
    & " max(neqty) neqty,max(lc_fc) lc_fc from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id  left join mst_item b on a.lcd_itemdid=b.item_id " _
    & " left join loadcap_mst_product_r c on a.lcd_itemdid=c.partno " _
    & " left join ( " _
    & " select lcd_itemdid,min(ton_mach) minton_mch,max(ton_mach) maxton_mch from mpp_gen_d " _
    & " where fltpp_doc='" & CmbDocument & "' " _
    & " and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' and lc_pp>0 " _
    & " group by lcd_itemdid " _
    & " ) vsub on a.lcd_itemdid=vsub.lcd_itemdid " _
    & " where fltpp_doc='" & CmbDocument & "' and lc_subcont='yes' and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' and lc_pp>0 and neday>0 " _
    & " and substring(no_mach from 1 for 1)='S'" _
    & " GROUP by reg_mold, a.lcd_itemdid,ith_qty order by 1, 2 asc, 3 asc"

    Set rsAsubct = Con.Execute(qry)

    anaSubcont.rows = 2
    If rsAsubct.RecordCount > 0 Then
        Erase aPart_sub
        ReDim aPart_sub(1 To rsAsubct.RecordCount, 1 To 10 + totalHari + 1) As Variant
        i = 1
        While Not rsAsubct.EOF
            aPart_sub(i, 1) = rsAsubct("lcd_itemdid")
            aPart_sub(i, 2) = rsAsubct("stok"): aPart_sub(i, 3) = rsAsubct("stok")
            aPart_sub(i, 4) = rsAsubct("item_muloq")
            aPart_sub(i, 5) = rsAsubct("item_perbox")
            aPart_sub(i, 6) = rsAsubct("lc_pp"): aPart_sub(i, 7) = rsAsubct("reg_mold")
            aPart_sub(i, 8) = rsAsubct("minstock")
            aPart_sub(i, 9) = rsAsubct("maxstock")
            aPart_sub(i, 10) = rsAsubct("mesin")
            For k = 11 To totalHari + 1
                aPart_sub(i, k) = "0"
            Next
            
            i = 1 + i
            rsAsubct.MoveNext
        Wend
        
        anaSubcont.rows = anaSubcont.rows + 1
        i = 2
        r = 2
        Dim tempProcedqty As Variant
        rsAsubct.MoveFirst
        While Not rsAsubct.EOF
            With anaSubcont
                If temp_mch = rsAsubct("mesin") Then
                    noItemPerMesin = noItemPerMesin + 1
                Else
                    noItemPerMesin = 1
                End If
                    tempProcedqty = rsAsubct("neqty") '0
                    .TextMatrix(i, 16) = FormatNumber(tempProcedqty, 0) 'FormatNumber(rsAsubct("lc_pp") - tempProcedqty, 0)
                    .TextMatrix(i, 0) = rsAsubct("mesin")
                    .TextMatrix(i, 1) = rsAsubct("lc_customer")
                    .TextMatrix(i, 2) = rsAsubct("lcd_itemdid")
                    .TextMatrix(i, 3) = rsAsubct("lc_itemname")
                    .TextMatrix(i, 4) = rsAsubct("minton")
                    .TextMatrix(i, 5) = rsAsubct("maxton")
                    .Col = 6: .Row = i
                    .CellAlignment = flexAlignLeftCenter
                    .TextMatrix(i, 6) = rsAsubct("reg_mold")
                    .TextMatrix(i, 7) = rsAsubct("cav")
                    .TextMatrix(i, 8) = rsAsubct("ct")
                    .TextMatrix(i, 9) = FormatNumber(rsAsubct("need_mp"), 1)
                    If rsAsubct("ct") > 0 Then
                        .TextMatrix(i, 10) = FormatNumber((3600 / rsAsubct("ct")) * rsAsubct("cav"), 0)
                        .TextMatrix(i, 10) = FormatNumber(.TextMatrix(i, 10), 0)
                        .TextMatrix(i, 11) = (.TextMatrix(i, 10) * 1) * rsAsubct("hourpshift")
                    Else
                        .TextMatrix(i, 10) = 0
                    End If
                    .TextMatrix(i, 12) = FormatNumber(rsAsubct("capday"), 0)
                    .TextMatrix(i, 17) = FormatNumber(rsAsubct("lc_fc"), 0)
                    
                    rsAsubct.MoveNext
                    If rsAsubct.EOF Then
                        
                    Else
                        temp_mch = rsAsubct("mesin")
                    End If
                    rsAsubct.MovePrevious
                    If temp_mch = rsAsubct("mesin") Then
                        .rows = .rows + 1
                        i = i + 1
                    Else
                        .rows = .rows + 3
                        i = i + 3
                    End If
                    
            End With
            temp_mch = rsAsubct("mesin")
            rsAsubct.MoveNext
        Wend

        ' / dapat mesin sub distinct
        qry = " select distinct on (mesin) mesin from " _
        & "(select max(no_mach) mesin" _
        & " from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id  left join mst_item b on a.lcd_itemdid=b.item_id " _
        & " left join ( " _
        & " select lcd_itemdid,min(ton_mach) minton_mch,max(ton_mach) maxton_mch from mpp_gen_d " _
        & " where fltpp_doc='" & CmbDocument & "' " _
        & " and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' and lc_pp>0 " _
        & " group by lcd_itemdid " _
        & " ) vsub on a.lcd_itemdid=vsub.lcd_itemdid " _
        & " where fltpp_doc='" & CmbDocument & "' and lc_subcont='yes' and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' and lc_pp>0 " _
        & " and substring(no_mach from 1 for 1)='S'" _
        & " GROUP by reg_mold, a.lcd_itemdid order by 1) dist_tab1"
        Set rsB = Con.Execute(qry)

        ReDim aMesinSubc(1 To rsB.RecordCount, 1 To totalHari + 3) As Variant
        i = 1
        While Not rsB.EOF
            aMesinSubc(i, 1) = rsB(0)
            For k = 4 To totalHari + 3
                aMesinSubc(i, k) = "0"
            Next
            
            i = 1 + i
            rsB.MoveNext
        Wend

'        For tanggal = 1 To totalHari
'            '~ saatnya mulai loop
'            plot_subc_v2 tanggal
'        Next
        hitungRekapSUB
    End If
    
    '# unprocessed
    qry = "select  no_mach, ton_mach ,lc_customer,  lcd_itemdid, lc_itemname, cav, ct, cap_p_day,fltpp_hkw,  neday, reg_mold, lc_pp, lc_subcont, lcneed_mp need_mp, hourpshift,neqty,lc_sisa_pp  " _
        & ",item_muloq,item_perbox from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
        & " left join mst_item b on a.lcd_itemdid=b.item_id " _
        & " where fltpp_doc='" & CmbDocument & "' " _
        & " and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' and lc_sisa_pp>0 and lc_subcont='no' and neday=0 " _
        & " order by no_mach asc, lc_customer asc, lcd_itemdid asc"
    Set rsAunpro = Con.Execute(qry)
    anaUnproc.rows = 2
    If rsAunpro.RecordCount > 0 Then
        qry = "select distinct on (lcd_itemdid) lcd_itemdid, stok,item_muloq  from " _
            & " (select lcd_itemdid,  coalesce(ith_qty,0) stok,item_muloq  " _
            & " from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
            & " left join mst_item b on a.lcd_itemdid=b.item_id " _
            & " where fltpp_doc='" & CmbDocument & "' and fltpp_rev=" & txtRevision & " and fltpp_ym='" & cmbPeriod & "' and lc_sisa_pp>0 and lc_subcont='no' and neday=0 " _
            & " order by no_mach asc, lc_customer asc, lcd_itemdid asc) tabela"
        Set rsB = Con.Execute(qry)
        
        Erase aPart_unpc
        ReDim aPart_unpc(1 To rsB.RecordCount, 1 To 4) As Variant
        i = 1
        While Not rsB.EOF
            aPart_unpc(i, 1) = rsB(0).Value: aPart_unpc(i, 2) = rsB(1).Value: aPart_unpc(i, 3) = rsB(1).Value: aPart_unpc(i, 4) = rsB(2).Value
            i = 1 + i
            rsB.MoveNext
        Wend
        anaUnproc.rows = anaUnproc.rows + 1
        i = 2
        r = 2
        While Not rsAunpro.EOF
            With anaUnproc
                 .TextMatrix(i, 0) = rsAunpro("no_mach")
                 .TextMatrix(i, 1) = rsAunpro("lc_customer")
                 .TextMatrix(i, 2) = rsAunpro("lcd_itemdid")
                 .TextMatrix(i, 3) = rsAunpro("lc_itemname")
                 .Col = 4
                 .Row = i
                 .CellAlignment = flexAlignLeftCenter
                 .TextMatrix(i, 4) = rsAunpro("reg_mold")
                 .TextMatrix(i, 5) = rsAunpro("cav")
                 .TextMatrix(i, 6) = rsAunpro("ct")
                 .TextMatrix(i, 7) = FormatNumber(rsAunpro("need_mp"), 1)
                If rsAunpro("ct") > 0 Then
                    .TextMatrix(i, 8) = FormatNumber((3600 / rsAunpro("ct")) * rsAunpro("cav"), 0)
                    .TextMatrix(i, 8) = FormatNumber(.TextMatrix(i, 8), 0)
                    .TextMatrix(i, 9) = (.TextMatrix(i, 8) * 1) * rsAunpro("hourpshift")
                Else
                    .TextMatrix(i, 8) = 0
                End If
                .TextMatrix(i, 10) = FormatNumber(rsAunpro("cap_p_day"), 0)
                If rsAunpro("item_perbox") = 0 Then
                    .TextMatrix(i, 14) = FormatNumber(isi(rsAunpro("item_muloq"), rsAunpro("lc_sisa_pp"), "a"), 0)
                Else
                    .TextMatrix(i, 14) = FormatNumber(isi(rsAunpro("item_perbox"), rsAunpro("lc_sisa_pp"), "a"), 0)
                End If
                
                rsAunpro.MoveNext
                If rsAunpro.EOF Then
                    
                Else
                    temp_mch = rsAunpro("no_mach")
                End If
                rsAunpro.MovePrevious
                If temp_mch = rsAunpro("no_mach") Then
                    .rows = .rows + 1
                    i = i + 1
                Else
                    .rows = .rows + 3
                    i = i + 3
                End If
            End With
            temp_mch = rsAunpro("no_mach")
            r = r + 1
            rsAunpro.MoveNext
        Wend
    End If
    
    lblstate.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub cmdLoad_Click()
    If PicListMPP.Visible Then
        PicListMPP.Visible = False
    Else
        qry = "select * from (select distinct on (mpp_doc_no) mpp_doc_no,mpp_revisi,ml_ym,ml_rev,ml_doc  from mpp_gen where extract(year from plandate)=" & Format(Now, "yyyy") & " and planqty>0 ) v1 order by ml_ym desc, mpp_revisi desc limit 1"
        Set RsBantu = Con.Execute(qry)
        fgmpp.rows = 1
        clearIN
        If RsBantu.RecordCount > 0 Then
            fgmpp.rows = 1 + RsBantu.RecordCount
            fgmpp.TextMatrix(1, 0) = 1
            fgmpp.TextMatrix(1, 1) = RsBantu("mpp_doc_no")
            fgmpp.TextMatrix(1, 2) = RsBantu("mpp_revisi")
            fgmpp.TextMatrix(1, 3) = RsBantu("ml_ym")
            fgmpp.TextMatrix(1, 4) = RsBantu("ml_rev")
            fgmpp.TextMatrix(1, 5) = RsBantu("ml_doc")
            CmbDocument = RsBantu("ml_doc") 'CmbDocument.AddItem RsBantu("ml_doc")
            txtRevision.AddItem RsBantu("ml_rev")
            cmbPeriod.AddItem RsBantu("ml_ym")
        End If
        Set RsBantu = Nothing
        PicListMPP.Visible = True
      
    End If
    sinkronGridcols
End Sub



Private Sub saveMPS()
    Dim u As Long, totalBaris As Long
    Dim rsMPP As ADODB.Recordset
    Dim revisi As String
    Dim nourut As String
    Dim forcas As Long
    Dim ostdpo As Double
    Dim qtyTgl As Double
    Dim ttlrows As Long
    Dim ttlCOls As Byte
    
    PG1.Visible = True
    If Len(cmbPeriod) < 5 Then Exit Sub
    qry = "select mpp_revisi,mpp_doc_no from mpp_gen where ml_ym='" & cmbPeriod & "' order by mpp_doc_no desc,mpp_revisi desc limit 1"
    
    Set rsMPP = Con.Execute(qry)
    If rsMPP.RecordCount > 0 Then
        revisi = rsMPP(0)
        revisi = revisi * 1 + 1
        If revisi * 1 < 10 Then
            revisi = "0" & revisi
        End If
        nourut = Left$(rsMPP(1), 2)
        nourut = nourut * 1 + 1
        If Len(nourut) = 1 Then nourut = "0" & nourut
    Else
        revisi = "00"
        nourut = "01"
    End If
    NoDocMPS = nourut & "/PPC/MPS/" & Right$(cmbPeriod, 2) & "/" & Mid$(cmbPeriod, 3, 2)
    rev_MPS = revisi
'    MsgBox rev_MPS
    Set rsMPP = Nothing
    Set rsMPP = New ADODB.Recordset
    rsMPP.Open "select * from mpp_gen where mpp_doc_no='" & NoDocMPS & "'", Con, adOpenDynamic, adLockOptimistic

    
    With anaGrid
        totalBaris = .rows - 2
        ttlrows = .rows - 1
        ttlCOls = .Cols - 1
        For u = 2 To ttlrows
            DoEvents
            If IsNumeric(.TextMatrix(u, 19)) Then '18
                forcas = .TextMatrix(u, 19) * 1 '18
            Else
                forcas = 0
            End If
            ostdpo = IIf(IsNumeric(.TextMatrix(u, 15)), .TextMatrix(u, 15), 0) '14
            If LenB(.TextMatrix(u, 0)) <> 0 Then
                For i = COLSDATE To ttlCOls
                    If IsNumeric(.TextMatrix(u, i)) Then
                        qtyTgl = .TextMatrix(u, i) * 1
                    Else
                        qtyTgl = 0
                    End If
                    If qtyTgl > 0 Then
                        rsMPP.AddNew
                        rsMPP!lcd_itemdid = .TextMatrix(u, 3)
                        rsMPP!partname = .TextMatrix(u, 4)
                        rsMPP!lc_customer = .TextMatrix(u, 1)
                        rsMPP!no_mach = .TextMatrix(u, 0)
                        rsMPP!ton_mach = getTonage(.TextMatrix(u, 0))
                        rsMPP!reg_mold = .TextMatrix(u, 5)
                        rsMPP!cav = .TextMatrix(u, 7) '6
                        rsMPP!ct = .TextMatrix(u, 8) '7
                        rsMPP!cap_p_day = .TextMatrix(u, 12) * 1 '11
                        rsMPP!lcvsmach = 1 * Left$(.TextMatrix(u, (COLSDATE - 1)), Len(.TextMatrix(u, (COLSDATE - 1))) - 1)
                        rsMPP!ml_doc = CmbDocument
                        rsMPP!ml_rev = txtRevision
                        rsMPP!ml_ym = cmbPeriod
                        rsMPP!ml_hkw = HKWs
                        rsMPP!ml_subcont = "no"
                        rsMPP!mpp_pp = .TextMatrix(u, 17) * 1 '16
                        rsMPP!needmp = .TextMatrix(u, 9) '8
                        rsMPP!cap_p_hour = .TextMatrix(u, 10) * 1 '9
                        rsMPP!cap_p_shift = .TextMatrix(u, 11) * 1 '10
                        rsMPP!hour_req = .TextMatrix(u, 13) * 1 '12
                        rsMPP!day_req = .TextMatrix(u, 14) * 1 '13
                        rsMPP!ost_po = ostdpo
                        rsMPP!FC = forcas
                        rsMPP!mpp_doc_no = NoDocMPS
                        rsMPP!plandate = Left$(cmbPeriod, 4) & "-" & Right$(cmbPeriod, 2) & "-" & Left$(.TextMatrix(1, i), 2)
                        rsMPP!planQty = qtyTgl
                        rsMPP!mpp_revisi = revisi
                        rsMPP.Update
                    End If
                Next
            End If
            PG1.Value = FormatNumber(((u - 1) * 100) / totalBaris, 0)
            PG1.ToolTipText = PG1.Value & " %"
        Next
    End With
    Set rsMPP = Nothing
    
    'subcont
'    With anaSubcont
'        totalBaris = .rows - 2
'        ttlrows = .rows - 1
'        ttlCOls = .Cols - 1
'        For u = 2 To ttlrows
'            If IsNumeric(.TextMatrix(u, 17)) Then
'                forcas = .TextMatrix(u, 17) * 1
'            Else
'                forcas = 0
'            End If
'            ostdpo = IIf(IsNumeric(.TextMatrix(u, 15)), .TextMatrix(u, 15), 0)
'            For i = 20 To ttlCOls
'                If LenB(.TextMatrix(u, 0)) <> 0 Then
'                    If IsNumeric(.TextMatrix(u, i)) Then
'                        qtyTgl = .TextMatrix(u, i) * 1
'                    Else
'                        qtyTgl = 0
'                    End If
'
'                    qry = "insert into mpp_gen (idmppgen ," _
'                        & " lcd_itemdid,partname,lc_customer , " _
'                        & " no_mach,ton_mach,reg_mold, cav,ct,cap_p_day,lcvsmach, " _
'                        & " ml_doc,ml_rev , ml_ym , " _
'                        & " ml_hkw , ml_subcont ,  mpp_pp , " _
'                        & " needmp , cap_p_hour, cap_p_shift, " _
'                        & " hour_req , day_req ,ost_po ,fc, mpp_doc_no,plandate,planqty,mpp_revisi) " _
'                        & " values (DEFAULT,'" & .TextMatrix(u, 2) & "', " _
'                        & "'" & .TextMatrix(u, 3) & "','" & .TextMatrix(u, 1) & "'," _
'                        & "'" & .TextMatrix(u, 0) & "'," & getTonage(.TextMatrix(u, 0)) & ",'" & .TextMatrix(u, 6) & "'," _
'                        & .TextMatrix(u, 7) & "," & .TextMatrix(u, 8) & "," _
'                        & .TextMatrix(u, 12) * 1 & "," & 1 * Left$(.TextMatrix(u, 19), Len(.TextMatrix(u, 19)) - 1) & "," _
'                        & "'" & CmbDocument & "'," & txtRevision & ",'" & cmbPeriod & "'," _
'                        & HKWs & ",'yes'," & .TextMatrix(u, 16) * 1 & "," & .TextMatrix(u, 9) & "," _
'                        & .TextMatrix(u, 10) * 1 & "," & .TextMatrix(u, 11) * 1 & "," & .TextMatrix(u, 13) * 1 & "" _
'                        & "," & .TextMatrix(u, 14) * 1 & "," & ostdpo & "" _
'                        & "," & forcas & ",'" & NoDocMPS & "'" _
'                        & ",'" & Left$(cmbPeriod, 4) & "-" & Right$(cmbPeriod, 2) & "-" & Left$(.TextMatrix(1, i), 2) & "'," _
'                        & qtyTgl & ",'" & revisi & "')"
'                    Con.Execute qry
'                End If
'            Next
'            PG1.Value = FormatNumber(((u - 1) * 100) / totalBaris, 0)
'            PG1.ToolTipText = PG1.Value & " %"
'        Next
'    End With
   
    Exit Sub
End Sub

Private Sub cmdlu_findDoc_Click()
    PopUp_MLDOC.Show 1
    CmbDocument.Text = PopUp_MLDOC.lu_nodoc
    txtRevision.Enabled = True
End Sub

Private Sub cmdSave_Click()
On Error GoTo hereEX
    If belumSimpan = False Then
        If fge.rows > 2 Then
            PicEditedList.Visible = True
        Else
            MsgBox "the data is already exist"
        End If
        Exit Sub
    End If
    If MsgBox("Are you sure want to save the data ", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    saveMPS
    
     MsgBox "Saved successfully", vbInformation, "Good"
    Exit Sub
hereEX:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub initWO(pPrint As Boolean)
    With anaGrid
        If .Col > (COLSDATE - 1) Then
            If IsNumeric(.TextMatrix(.Row, .Col)) Then
                If .TextMatrix(.Row, .Col) * 1 > 0 Then
                    Call getWO(pPrint, .Row, .Col)
                End If
            End If
        End If
    End With
End Sub

Private Sub Command1_Click()
    Dim xf As Double, pos As Integer
    Dim ttlrows As Double
    Dim stringCari As String
    With anaGrid
        ttlrows = .rows - 1
        If posisisFind + 1 >= ttlrows Then
            posisisFind = 2
        Else
            posisisFind = 1 + posisisFind
        End If
        For xf = posisisFind To ttlrows
            stringCari = LCase$(.TextMatrix(xf, 3))
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

Private Sub cuemd_print_Click()
    If MsgBox("Print selected schedule ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    cuemd_print.Enabled = False
    Dim ttlBaris As Long
    Dim ay As Long
    lblPlease.Visible = True
    PG1.Visible = True
    Me.Refresh
    pStPrinter = False
    PopUp_PrinterPrint.Show 1
    If pStPrinter = False Then MsgBox "Canceled": Exit Sub
    With anaGrid
        ttlBaris = lvprintp.ListItems.Count
        For ay = 1 To ttlBaris
            .Col = lvprintp.ListItems(ay).SubItems(4)
            .Row = lvprintp.ListItems(ay).SubItems(5)
            
            'Call initWO(True)
            Call getWO(True, .Row, .Col)
            .CellBackColor = RGB(WO_R, WO_G, WO_B)
            
            PG1.Value = FormatNumber(((ay) * 100) / ttlBaris, 0)
            PG1.ToolTipText = PG1.Value & "%"
        Next
        lvprintp.ListItems.Clear
    End With
    PG1.Visible = False
    pic_pp_or_p.Visible = False
End Sub

Private Sub GenerateCode128(Str As String, Optional BarWidth As Integer = 1)
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
    
End Sub

Private Sub cuemd_printprev_Click()
    Call initWO(False)
End Sub

Private Sub fge_Click()
    lblTtl_rev.Caption = fge.rows - 1
End Sub

Private Sub fgmpp_Click()
    With fgmpp
        NoDocMPS = .TextMatrix(.Row, 1)
        rev_MPS = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub fgmpp_DblClick()
    fgmpp_KeyPress 13
    cmdSave.Visible = False
End Sub

Private Sub fgmpp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then
        Clipboard.Clear
        Clipboard.SetText fgmpp.Clip
    ElseIf KeyCode = 46 Then
        cmdDelete_Click
    End If
End Sub

Private Sub fgmpp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        loadMstSubcont
        lMaterialPurg
        lvprintp.ListItems.Clear
        With fgmpp
            PicListMPP.Visible = False
            Screen.MousePointer = 11
            NoDocMPS = .TextMatrix(.Row, 1)
            rev_MPS = .TextMatrix(.Row, 2)
            LoadMPPINJ NoDocMPS, rev_MPS, .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), .TextMatrix(.Row, 5)
            
            CmbDocument.Text = .TextMatrix(.Row, 5)
            txtRevision.Text = .TextMatrix(.Row, 4)
            cmbPeriod.Text = .TextMatrix(.Row, 3)
            txtRevision.Enabled = False
            'tandai mesin trial
            plotMchTrial anaGrid
            loadFWO
            loadRevisiLog
            Screen.MousePointer = 0
            loadMPPINJ_d NoDocMPS, rev_MPS, .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), .TextMatrix(.Row, 5)
            hitungRekapINJ
            reinitOverloadQty
            loadMPPSUBC NoDocMPS, rev_MPS, .TextMatrix(.Row, 5), .TextMatrix(.Row, 4), .TextMatrix(.Row, 3)
            loadBOM
            belumSimpan = False
            '----suggestion overload
            loadSuggestion
        End With
        fge.rows = 1
        fge.rows = 2
        sinkronGridcols
    End If
    
End Sub

Private Sub loadRevisiLog()
    Dim temprevv As String
    qry = "select * from mpp_gen_rev where mps_doc='" & NoDocMPS & "' and mps_rev<='" & rev_MPS & "'"
    Set rsa_rev = Con.Execute(qry)
    Erase sd_mpsrev
    rsa_rev.Sort = "mps_rev asc"
    For i = 1 To rsa_rev.RecordCount
        rsa_rev.AbsolutePosition = i
        If temprevv <> rsa_rev("mps_rev") Then
            If (Not sd_mpsrev) <> -1 Then
                ReDim Preserve sd_mpsrev(1 To UBound(sd_mpsrev) + 1) As String
                sd_mpsrev(UBound(sd_mpsrev)) = rsa_rev("mps_rev")
            Else
                ReDim sd_mpsrev(1 To 1) As String
                sd_mpsrev(1) = rsa_rev("mps_rev")
            End If
        End If
        temprevv = rsa_rev("mps_rev")
    Next
End Sub

Private Sub loadMPPSUBC(pDocMps As String, pDocMpsRev As String, pdocltpp As String, pdocltpprev As String, pperiod As String)
    Dim tempdate As String
    qry = "select * from (select distinct on (reg_mold) * from (select a.lcd_itemdid,partname,a.lc_customer,a.no_mach,a.ton_mach,a.reg_mold ," _
    & " a.cav ,a.ct , a.cap_p_day ,a.lcvsmach , ml_doc ,ml_rev ,ml_ym ,ml_hkw ,ml_subcont , " _
    & " mpp_pp,needmp ,cap_p_hour,cap_p_shift ,hour_req ,day_req ,ost_po ,fc,mpp_doc_no , " _
    & " plandate ,planqty , mpp_revisi,hourpshift from mpp_gen a " _
    & " inner join mpp_gen_d b on a.ml_ym=b.fltpp_ym and a.ml_subcont=lc_subcont and ml_doc=b.fltpp_doc and a.partname=lc_itemname " _
    & " and a.lcd_itemdid=b.lcd_itemdid and a.no_mach=b.no_mach and a.reg_mold=b.reg_mold and a.lc_customer=b.lc_customer and a.ml_rev=b.fltpp_rev " _
    & " where mpp_doc_no='" & pDocMps & "' and mpp_revisi='" & pDocMpsRev & "' and ml_ym='" & pperiod & "' " _
    & " and ml_doc='" & pdocltpp & "' and ml_subcont='yes' and ml_rev=" & pdocltpprev & ") as wrapmpp) as wrapmpp2 order by no_mach asc,lc_customer asc,lcd_itemdid asc"
    Set rsAsubct = Con.Execute(qry)
    anaSubcont.rows = 2
    If rsAsubct.RecordCount > 0 Then
        bhulan = DateSerial(Left$(rsAsubct("ml_ym"), 4), Right$(rsAsubct("ml_ym"), 2) * 1, 1)
        totalHari = Int(Format(dhLastDayInMonth(bhulan), "dd"))
        DAYtoGrid anaSubcont, rsAsubct("ml_ym")
        Dim tempProcedqty As Variant
        rsAsubct.MoveFirst
        i = 2
        With anaSubcont
            .rows = 3
            .Col = 6
            While Not rsAsubct.EOF
                If temp_mch = rsAsubct("no_mach") Then
                    noItemPerMesin = noItemPerMesin + 1
                Else
                    noItemPerMesin = 1
                End If
                tempProcedqty = rsAsubct("mpp_pp")
                .TextMatrix(i, 16) = FormatNumber(tempProcedqty, 0) 'FormatNumber(rsAsubct("lc_pp") - tempProcedqty, 0)
                .TextMatrix(i, 0) = rsAsubct("no_mach")
                .TextMatrix(i, 1) = rsAsubct("lc_customer")
                .TextMatrix(i, 2) = rsAsubct("lcd_itemdid")
                .TextMatrix(i, 3) = rsAsubct("partname")
                .TextMatrix(i, 4) = 0
                .TextMatrix(i, 5) = rsAsubct("ton_mach")
                .Row = i
                .CellAlignment = flexAlignLeftCenter
                .TextMatrix(i, 6) = rsAsubct("reg_mold")
                .TextMatrix(i, 7) = rsAsubct("cav")
                .TextMatrix(i, 8) = rsAsubct("ct")
                .TextMatrix(i, 9) = FormatNumber(rsAsubct("needmp"), 1)
                If rsAsubct("ct") > 0 Then
                    .TextMatrix(i, 10) = FormatNumber((3600 / rsAsubct("ct")) * rsAsubct("cav"), 0)
                    .TextMatrix(i, 10) = FormatNumber(.TextMatrix(i, 10), 0)
                    .TextMatrix(i, 11) = (.TextMatrix(i, 10) * 1) * rsAsubct("hourpshift")
                Else
                    .TextMatrix(i, 10) = 0
                    .TextMatrix(i, 11) = 0
                End If
                .TextMatrix(i, 12) = FormatNumber(rsAsubct("cap_p_day"), 0)
                .TextMatrix(i, 17) = FormatNumber(rsAsubct("fc"), 0)
                
                rsAsubct.MoveNext
                If rsAsubct.EOF Then
                    
                Else
                    temp_mch = rsAsubct("no_mach")
                End If
                rsAsubct.MovePrevious
                If temp_mch = rsAsubct("no_mach") Then
                    .rows = .rows + 1
                    i = i + 1
                Else
                    .rows = .rows + 3
                    i = i + 3
                End If
                temp_mch = rsAsubct("no_mach")
                rsAsubct.MoveNext
            Wend
        End With
    End If
    
    qry = "select a.lcd_itemdid,partname,a.lc_customer,a.no_mach,a.ton_mach,a.reg_mold ," _
    & " a.cav ,a.ct , a.cap_p_day ,a.lcvsmach , ml_doc ,ml_rev ,ml_ym ,ml_hkw ,ml_subcont , " _
    & " mpp_pp,needmp ,cap_p_hour,cap_p_shift ,hour_req ,day_req ,ost_po ,fc,mpp_doc_no , " _
    & " plandate ,planqty , mpp_revisi,hourpshift from mpp_gen a " _
    & " inner join mpp_gen_d b on a.ml_ym=b.fltpp_ym and a.ml_subcont=lc_subcont and ml_doc=b.fltpp_doc and a.partname=lc_itemname " _
    & " and a.lcd_itemdid=b.lcd_itemdid and a.no_mach=b.no_mach and a.reg_mold=b.reg_mold and a.lc_customer=b.lc_customer and a.ml_rev=b.fltpp_rev " _
    & " where mpp_doc_no='" & pDocMps & "' and mpp_revisi='" & pDocMpsRev & "' and ml_ym='" & pperiod & "' " _
    & " and ml_doc='" & pdocltpp & "' and ml_subcont='yes' and ml_rev=" & pdocltpprev & ""
    Set rsAsubct = Con.Execute(qry)
    If rsAsubct.RecordCount > 0 Then
        Dim totalBaris As Long
        rsAsubct.Fields("no_mach").Properties("Optimize") = True
        rsAsubct.Fields("lcd_itemdid").Properties("Optimize") = True
        rsAsubct.Fields("reg_mold").Properties("Optimize") = True
        rsAsubct.Fields("plandate").Properties("Optimize") = True
        With anaSubcont
            PG1.Visible = True
            PG1.Value = 0
            totalBaris = .rows - 1
            For i = 2 To .rows - 1
                DoEvents
                .Row = i
                For k = 20 To .Cols - 1
                    qry = "no_mach='" & .TextMatrix(i, 0) & "' and " _
                        & "lcd_itemdid='" & .TextMatrix(i, 2) & "' and " _
                        & "reg_mold='" & .TextMatrix(i, 6) & "' " _
                        & "and plandate='" & Left$(pperiod, 4) & "-" & Right$(pperiod, 2) & "-" & Left$(.TextMatrix(1, k), 2) & "'"
                    rsAsubct.Filter = adFilterNone
                    rsAsubct.Filter = qry
                    If rsAsubct.RecordCount > 0 Then
                        .TextMatrix(i, k) = FormatNumber(rsAsubct("planqty"), 0)
                    End If
                Next
                PG1.Value = FormatNumber(((i) * 100) / totalBaris, 0)
                PG1.ToolTipText = PG1.Value & "%"
                
            Next
            PG1.Visible = False
            
            If (Not sd_mpsrev) <> -1 Then
                For r = 1 To UBound(sd_mpsrev)
                    For i = 2 To .rows - 1
                        DoEvents
                        If LenB(.TextMatrix(i, 0)) <> 0 Then
                            For c = 20 To .Cols - 1
                                tempdate = Left$(pperiod, 4) & "-" & Right$(pperiod, 2) & "-" & Left$(.TextMatrix(1, c), 2)
                                qry = "mps_rev='" & sd_mpsrev(r) & "' and machine='" & .TextMatrix(i, 0) & "' and itemid='" & .TextMatrix(i, 2) & "' " _
                                        & " and mold='" & .TextMatrix(i, 6) & "' and plandate='" & tempdate & "' and planqty=" & .TextMatrix(i, c) * 1
                                rsa_rev.Filter = adFilterNone
                                rsa_rev.Filter = qry
                                If rsa_rev.RecordCount > 0 Then
                                    .Row = i
                                    .Col = c
                                    If sd_mpsrev(r) = "01" Then
                                        .CellBackColor = RGB(255, 85, 127) ' merah muda
                                    ElseIf sd_mpsrev(r) = "02" Then
                                        .CellBackColor = RGB(255, 255, 1) ' kuning
                                    ElseIf sd_mpsrev(r) = "03" Then
                                        .CellBackColor = RGB(127, 212, 255) ' biru muda
                                    ElseIf sd_mpsrev(r) = "04" Then
                                        .CellBackColor = RGB(255, 127, 212) ' pink
                                    ElseIf sd_mpsrev(r) = "05" Then
                                        .CellBackColor = RGB(0, 255, 0) ' green
                                    ElseIf sd_mpsrev(r) = "06" Then
                                        .CellBackColor = RGB(212, 170, 255) ' ungu
                                    End If
                                End If
                            Next
                        End If
                        PG1.Value = FormatNumber(((i) * 100) / totalBaris, 0)
                        PG1.ToolTipText = PG1.Value & "% [" & r & "]"
                    Next
                Next
            End If
        End With
        hitungRekapSUB
    End If
End Sub

Private Sub loadFWO()
    qry = "SELECT mesinno,partno,moldno,qty,issudate,qty/cappday dayrun,extract(day from issudate) hari FROM worko a where " _
    & " substring(wo_no from 9 for 2)='" & Right$(cmbPeriod, 2) & "' and substring(wo_no from 12 for 2)='" & Mid$(cmbPeriod, 3, 2) & "' "
    Set rsWO = Con.Execute(qry)
    If rsWO.RecordCount > 0 Then
        rsWO.Fields("mesinno").Properties("Optimize") = True
        rsWO.Fields("partno").Properties("Optimize") = True
        rsWO.Fields("moldno").Properties("Optimize") = True
        rsWO.Fields("issudate").Properties("Optimize") = True
        rsWO.Fields("hari").Properties("Optimize") = True
    End If
End Sub

Private Sub loadMPPINJ_d(pMpp_doc As String, pMpp_rev As String, pperiod As String, prev As String, pNoLTPP As String)
    Dim totalBaris As Long
    Dim tempdate As String
    Dim totalRDB As Long
    Dim totalKolom As Byte
'    Dim dimana As String
   
    If ckprodplan.Value Then
        'dimana = " and lc_pp>0 or planqty >0 "
        qry = "SELECT rnomachine no_mach,ritemid lcd_itemdid,rreg_mold reg_mold," _
        & " rplandate plandate,rplanqty planqty, rcav_std cav_std, rtypelabel typelabel, rct_scnd ct_scnd, " _
        & " rlcneed_mp lcneed_mp, risno isno,rshiftusg shiftusg, rmpower mpower, " _
        & " rhourpshift hourpshift, rtimeupdate timeupdate, rlc_pp lc_pp FROM " _
        & " f_loadcapdetail('" & pNoLTPP & "'," & prev & ",'" & pperiod & "','" & pMpp_doc & "','" & pMpp_rev & "','f')"
        Set RsA = Con.Execute(qry)
    Else
        qry = "SELECT rnomachine no_mach,ritemid lcd_itemdid,rreg_mold reg_mold," _
        & " rplandate plandate,rplanqty planqty, rcav_std cav_std, rtypelabel typelabel, rct_scnd ct_scnd, " _
        & " rlcneed_mp lcneed_mp, risno isno,rshiftusg shiftusg, rmpower mpower, " _
        & " rhourpshift hourpshift, rtimeupdate timeupdate, rlc_pp lc_pp FROM " _
        & " f_loadcapdetail('" & pNoLTPP & "'," & prev & ",'" & pperiod & "','" & pMpp_doc & "','" & pMpp_rev & "','')"
        Set RsA = Con.Execute(qry)
    End If

'    qry = "select a.no_mach,a.lcd_itemdid,a.reg_mold,plandate,planqty,cav_std " _
        & ",typelabel,ct_scnd,lcneed_mp,isno,shiftusg,mpower,hourpshift,timeupdate,lc_pp" _
        & " from mpp_gen a left join " _
        & " ( select  a.no_mach,   lcd_itemdid,  reg_mold, lc_pp, lcneed_mp need_mp, hourpshift " _
        & " ,coalesce(ith_qty,0) stok,cav_std,ct_scnd " _
        & " ,lcneed_mp,mpower,shiftusg,timeupdate from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
        & " where fltpp_doc='" & pNoLTPP & "' and lc_subcont='no' " _
        & " and fltpp_rev=" & prev & "  and fltpp_ym='" & pperiod & "' " _
        & " order by a.no_mach asc, lc_customer asc, lcd_itemdid asc " _
        & " ) vb1 on a.lcd_itemdid=vb1.lcd_itemdid and a.no_mach=vb1.no_mach and a.reg_mold=vb1.reg_mold" _
        & " left join loadcap_mst_product_r c on a.lcd_itemdid=c.partno " _
        & " where mpp_doc_no='" & pMpp_doc & "' and mpp_revisi='" & pMpp_rev & "' " & dimana
    
  
     'qry = "select a.no_mach,a.lcd_itemdid,a.reg_mold,plandate,planqty,cav_std " _
        & ",typelabel,ct_scnd,lcneed_mp,isno,shiftusg,mpower,hourpshift,timeupdate,lc_pp" _
        & " from mpp_gen a left join " _
        & " ( select  a.no_mach,   lcd_itemdid,  reg_mold, lc_pp, lcneed_mp need_mp, hourpshift " _
        & " ,coalesce(ith_qty,0) stok,cav_std,ct_scnd " _
        & " ,lcneed_mp,mpower,shiftusg,timeupdate from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
        & " where fltpp_doc='" & pNoLTPP & "' and lc_subcont='no' " _
        & " and fltpp_rev=" & prev & "  and fltpp_ym='" & pperiod & "' " _
        & " order by a.no_mach asc, lc_customer asc, lcd_itemdid asc " _
        & " ) vb1 on a.lcd_itemdid=vb1.lcd_itemdid and a.no_mach=vb1.no_mach and a.reg_mold=vb1.reg_mold" _
        & " left join loadcap_mst_product_r c on a.lcd_itemdid=c.partno " _
        & " where mpp_doc_no='" & pMpp_doc & "' and mpp_revisi='" & pMpp_rev & "' "
    qry = "SELECT rnomachine no_mach,ritemid lcd_itemdid,rreg_mold reg_mold," _
    & " rplandate plandate,rplanqty planqty, rcav_std cav_std, rtypelabel typelabel, rct_scnd ct_scnd, " _
    & " rlcneed_mp lcneed_mp, risno isno,rshiftusg shiftusg, rmpower mpower, " _
    & " rhourpshift hourpshift, rtimeupdate timeupdate, rlc_pp lc_pp,rtypelabelbox typelabelbox FROM " _
    & " f_loadcapdetail('" & pNoLTPP & "'," & prev & ",'" & pperiod & "','" & pMpp_doc & "','" & pMpp_rev & "','')"
    
    
    Set rsA_aks = Con.Execute(qry)
        
    If RsA.RecordCount > 0 Then
           
        RsA.Fields("no_mach").Properties("Optimize") = True
        RsA.Fields("lcd_itemdid").Properties("Optimize") = True
        RsA.Fields("reg_mold").Properties("Optimize") = True
             
        
        rsa_rev.Fields("machine").Properties("Optimize") = True
        rsa_rev.Fields("itemid").Properties("Optimize") = True
        rsa_rev.Fields("mold").Properties("Optimize") = True
        rsa_rev.Fields("plandate").Properties("Optimize") = True
        rsa_rev.Fields("planqty").Properties("Optimize") = True
        rsa_rev.Fields("mps_rev").Properties("Optimize") = True
        RsA.Sort = "plandate ASC"
        With anaGrid
            PG1.Visible = True
            PG1.Value = 0
            totalBaris = .rows - 1
            totalKolom = .Cols - 1
            
            For i = 2 To totalBaris
                DoEvents
                .Row = i
                If LenB(.TextMatrix(i, 0)) <> 0 Then
                    qry = "no_mach='" & .TextMatrix(i, 0) & "' and " _
                            & "lcd_itemdid='" & .TextMatrix(i, 3) & "' and " _
                            & "reg_mold='" & .TextMatrix(i, 5) & "' "
                    RsA.Filter = adFilterNone
                    RsA.Filter = qry

                    RsA.Sort = "plandate ASC"
                    totalRDB = RsA.RecordCount
                    If totalRDB > 0 Then
                        
                        For r = 1 To totalRDB
                            RsA.AbsolutePosition = r
                            If RsA("planqty") > 0 Then
                            For k = COLSDATE To totalKolom
                                If Format(RsA!plandate, "dd") = Left$(.TextMatrix(1, k), 2) Then
                                    .TextMatrix(i, k) = FormatNumber(RsA("planqty"), 0)
                                End If
                            Next
                            End If
                        Next
                        
                    End If
                End If

                PG1.Value = FormatNumber(((i) * 100) / totalBaris, 0)
                PG1.ToolTipText = PG1.Value & "%"
            Next
            
            If (Not sd_mpsrev) <> -1 Then
                For r = 1 To UBound(sd_mpsrev)
                    For i = 2 To totalBaris
                        DoEvents
                        If LenB(.TextMatrix(i, 0)) <> 0 Then
                            For c = COLSDATE To totalKolom
                                tempdate = Left$(pperiod, 4) & "-" & Right$(pperiod, 2) & "-" & Left$(.TextMatrix(1, c), 2)
                                If IsNumeric(.TextMatrix(i, c)) Then
                                    qry = "mps_rev='" & sd_mpsrev(r) & "' and machine='" & .TextMatrix(i, 0) & "' and itemid='" & .TextMatrix(i, 3) & "' " _
                                        & " and mold='" & .TextMatrix(i, 5) & "' and plandate='" & tempdate & "' and planqty=" & .TextMatrix(i, c) * 1
                                Else
                                    qry = "mps_rev='" & sd_mpsrev(r) & "' and machine='" & .TextMatrix(i, 0) & "' and itemid='" & .TextMatrix(i, 3) & "' " _
                                        & " and mold='" & .TextMatrix(i, 5) & "' and plandate='" & tempdate & "' and planqty=0"
                                End If
                                
                                rsa_rev.Filter = adFilterNone
                                rsa_rev.Filter = qry
                                If rsa_rev.RecordCount > 0 Then
                                    .Row = i
                                    .Col = c
                                    If sd_mpsrev(r) = "01" Then
                                        .CellBackColor = RGB(255, 85, 127) ' merah muda
                                    ElseIf sd_mpsrev(r) = "02" Then
                                        .CellBackColor = RGB(255, 255, 1) ' kuning
                                    ElseIf sd_mpsrev(r) = "03" Then
                                        .CellBackColor = RGB(127, 212, 255) ' biru muda
                                    ElseIf sd_mpsrev(r) = "04" Then
                                        .CellBackColor = RGB(255, 127, 212) ' pink
                                    ElseIf sd_mpsrev(r) = "05" Then
                                        .CellBackColor = RGB(0, 255, 0) ' green
                                    ElseIf sd_mpsrev(r) = "06" Then
                                        .CellBackColor = RGB(212, 170, 255) ' ungu
                                    End If
                                End If
                            Next
                        End If
                        PG1.Value = FormatNumber(((i) * 100) / totalBaris, 0)
                        PG1.ToolTipText = PG1.Value & "% [" & r & "]"
                    Next
                Next
            End If
            '#LOADING WO
            For i = 2 To totalBaris
                DoEvents
                .Row = i
                If LenB(.TextMatrix(i, 0)) <> 0 Then
                    qry = "mesinno='" & .TextMatrix(i, 0) & "' and " _
                            & "partno='" & .TextMatrix(i, 3) & "' and " _
                            & "moldno='" & .TextMatrix(i, 5) & "' "
                    rsWO.Filter = adFilterNone
                    rsWO.Filter = qry
                    totalRDB = rsWO.RecordCount
                    If totalRDB > 0 Then
                        For r = 1 To totalRDB
                            rsWO.AbsolutePosition = r
                            For k = COLSDATE To totalKolom
                                If Format(rsWO!issudate, "dd") = Left$(.TextMatrix(1, k), 2) Then
                                    .Col = k
                                    .CellBackColor = RGB(WO_R, WO_G, WO_B)
                                End If
                            Next
                        Next
                    End If
                End If
    
                PG1.Value = FormatNumber(((i) * 100) / totalBaris, 0)
                PG1.ToolTipText = PG1.Value & "% [Loading WO Data]"
            Next

        End With
        PG1.Visible = False
    End If
End Sub

Private Sub LoadMPPINJ(PmppDoc As String, PmppRev As String, pperiod As String, prev As String, pNoLTPP As String)
    Dim dimana As String
    Dim groupBy As String
    If ckprodplan.Value Then
        dimana = " and lc_pp>0 or planqty>0"
    End If
    'groupBy = "lcd_itemdid,partname,lc_customer,cav,no_mach,lcvsmach,reg_mold,ml_ym,ton_mach,ct," _
    & " cap_p_day,ml_doc,ml_rev,ml_hkw,ml_subcont,mpp_pp,needmp,cap_p_hour,cap_p_shift,hour_req,ost_po,fc"

    'qry = "select ml_ym,vb1.lcd_itemdid,v1.partname,v1.lc_customer,v1.no_mach,v1.ton_mach,isno,typelabel," _
    & " v1.reg_mold,v1.cav,v1.ct,v1.cap_p_day,v1.lcvsmach,ml_doc,ml_rev,ml_hkw,ml_subcont,mpp_pp,needmp,cap_p_hour,planqty," _
    & " cap_p_shift,hour_req,ost_po,fc,coalesce(ith_qty,0) stok, item_muloq,item_perbox,cav_std,hourpshift from " _
    & " (select lcd_itemdid,partname,lc_customer,cav,no_mach,lcvsmach,reg_mold,ml_ym,ton_mach,ct,cap_p_day,ml_doc,ml_rev,ml_hkw,ml_subcont,mpp_pp,needmp,cap_p_hour,cap_p_shift,hour_req,ost_po,fc,sum(planqty) planqty from mpp_gen " _
    & " where mpp_doc_no ='" & PmppDoc & "' and mpp_revisi='" & PmppRev & "' and ml_rev=" & prev & " and ml_doc='" & pNoLTPP & "' " _
    & " and ml_subcont='no' group by " & groupBy & " ) v1 left join v_stockith vd on v1.lcd_itemdid=vd.ith_item_id " _
    & "left join mst_item b on v1.lcd_itemdid=b.item_id " _
    & " left join loadcap_mst_product_r c on v1.lcd_itemdid=c.partno " _
    & " right join " _
    & " (select  a.no_mach,   lcd_itemdid,  reg_mold, lc_pp, lcneed_mp need_mp, hourpshift " _
        & " ,coalesce(ith_qty,0) stok,cav_std,ct_scnd " _
        & " ,lcneed_mp,mpower,shiftusg from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
         & " where fltpp_doc='" & pNoLTPP & "' and lc_subcont='no' " _
         & " and fltpp_rev=" & prev & "  and fltpp_ym='" & pperiod & "' " _
         & " order by a.no_mach asc, lc_customer asc, lcd_itemdid asc " _
    & " ) vb1 on v1.lcd_itemdid=vb1.lcd_itemdid and v1.no_mach=vb1.no_mach and v1.reg_mold=vb1.reg_mold" _
    & dimana & " order by no_mach asc,lc_customer asc,lcd_itemdid asc"
    
    qry = "select fltpp_ym ml_ym,v2.lcd_itemdid,item_name partname,v2.lc_customer,v2.no_mach,v2.ton_mach,isno,typelabel,hourpshift," _
     & " v2.reg_mold,lc_pp,v2.cav,v2.ct,v2.cap_p_day,v2.lcvsmach,fltpp_doc ml_doc,fltpp_rev ml_rev,fltpp_hkw ml_hkw,coalesce(mpp_pp,0) mpp_pp, " _
      & " need_mp needmp,cap_p_hour,coalesce(planqty,0) planqty,cap_p_shift,lc_fc fc,coalesce(ith_qty,0) stok, item_muloq,item_perbox,coalesce(colordesc,'-') colordesc from " _
    & " (select lcd_itemdid,partname,no_mach,reg_mold,mpp_pp,cap_p_hour,cap_p_shift,ost_po, sum(planqty) planqty from mpp_gen " _
    & " where mpp_doc_no ='" & PmppDoc & "' and mpp_revisi='" & PmppRev & "' and ml_rev=" & prev & " and ml_doc='" & pNoLTPP & "' and ml_ym='" & pperiod & "' " _
     & " and ml_subcont='no' group by lcd_itemdid,partname,no_mach,reg_mold, mpp_pp,cap_p_hour,cap_p_shift,ost_po " _
   & " ) v1 right join (select  a.no_mach,   lcd_itemdid,  reg_mold, lc_pp, lcneed_mp need_mp, hourpshift " _
         & " ,coalesce(ith_qty,0) stok,ct_scnd,ton_mach,cav,ct,cap_p_day,lcvsmach,lc_fc " _
         & " ,lcneed_mp,mpower,shiftusg,fltpp_ym,lc_customer,fltpp_doc,fltpp_rev,fltpp_hkw from mpp_gen_d a left join v_stockith vd on a.lcd_itemdid=vd.ith_item_id " _
         & " where fltpp_doc='" & pNoLTPP & "' and lc_subcont='no' and fltpp_rev=" & prev & "  and fltpp_ym='" & pperiod & "' " _
   & " ) v2 on v1.lcd_itemdid=v2.lcd_itemdid and v1.no_mach=v2.no_mach and v1.reg_mold=v2.reg_mold " _
   & " left join loadcap_mst_product_r c on v2.lcd_itemdid=c.partno left join v_stockith vd on v2.lcd_itemdid=vd.ith_item_id " _
   & " left join mst_item b on v2.lcd_itemdid=b.item_id " _
   & " where stscode_id='01' " _
   & dimana _
   & " order by v2.no_mach ASC, v2.lc_customer asc,v2.lcd_itemdid asc"
    Set RsA = Con.Execute(qry)
    
    anaGrid.rows = 2
    If RsA.RecordCount > 0 Then
        bhulan = DateSerial(Left(RsA("ml_ym"), 4), Right$(RsA("ml_ym"), 2) * 1, 1)
        totalHari = Int(Format(dhLastDayInMonth(bhulan), "dd"))
        
        ''---khusus
        Dim indexTgl As Integer
        
        anaGrid.Cols = COLSDATE
        indexTgl = COLSDATE - 1
        anaGrid.Cols = indexTgl + 1 + totalHari
        With anaGrid
            For i = 1 To Int(Format(dhLastDayInMonth(bhulan), "dd"))
                .TextMatrix(0, indexTgl + i) = Format(DateSerial(Left$(RsA("ml_ym"), 4), Right$(RsA("ml_ym"), 2), i), "ddd")
                .TextMatrix(1, indexTgl + i) = Format(DateSerial(Left$(RsA("ml_ym"), 4), Right$(RsA("ml_ym"), 2), i), "dd-mmm")
            Next
        End With
        
        ''end khusus
        
        Erase aPart
        ReDim aPart(1 To RsA.RecordCount, 1 To 10 + totalHari + 1) As Variant
        '^day off
        qry = "select extract(day from work_date) harioff from plansys_setoffday " _
            & " where extract(month from work_date)=" & Right$(RsA("ml_ym"), 2) & " and extract(year from work_date)=" & Left$(RsA("ml_ym"), 4) & " and work_status=false order by 1"
        Set RsBantu = Con.Execute(qry)
        If RsBantu.RecordCount > 0 Then
            Erase aDayOFF
            ReDim aDayOFF(1 To RsBantu.RecordCount) As Variant
            i = 1
            While Not RsBantu.EOF
                aDayOFF(i) = RsBantu(0)
                i = 1 + i
                RsBantu.MoveNext
            Wend
        End If
        
        '^day ovr
        qry = "select extract(day from wrk_date) hariovr,no_mach from mpp_setovrtime " _
            & " where extract(year from wrk_date)=" & Left$(RsA("ml_ym"), 4) & " and extract(month from wrk_date)=" & Right$(RsA("ml_ym"), 2)
        Set RsBantu = Con.Execute(qry)
        If RsBantu.RecordCount > 0 Then
            Erase aDayOvr
            ReDim aDayOvr(1 To RsBantu.RecordCount, 1 To 2) As Variant
            i = 1
            While Not RsBantu.EOF
                aDayOvr(i, 1) = RsBantu(1)
                aDayOvr(i, 2) = RsBantu(0)
                i = 1 + i
                RsBantu.MoveNext
            Wend
        End If
        
        'jadwal 1 bulan " & right$(cmbPeriod, 2) & " | " & Left(cmbPeriod, 4) & "
        qry = "select extract(day from delv_date),part_no,sum(qty) qty from mpp_delv_plan where " _
            & " extract(month from delv_date)=" & Right$(RsA("ml_ym"), 2) & " and extract(year from delv_date)=" & Left$(RsA("ml_ym"), 4) _
            & " group by part_no , delv_date order by 1 asc"
        Set RsBantu = Con.Execute(qry)
        If RsBantu.RecordCount > 0 Then
            Erase aJadwal
            ReDim aJadwal(1 To RsBantu.RecordCount, 1 To 3)
            i = 1
            While Not RsBantu.EOF
                aJadwal(i, 1) = RsBantu(0)
                aJadwal(i, 2) = RsBantu(1)
                aJadwal(i, 3) = RsBantu(2)
                i = i + 1
                RsBantu.MoveNext
            Wend
        Else
            Erase aJadwal
            ReDim aJadwal(1 To 1, 1 To 3) As Variant
        End If
        Set RsBantu = Nothing
        
        
        qry = "select v2.no_mach,material,ton_mach from " _
        & " (select v1.no_mach,v1.ton_mach from " _
        & " (select distinct on (no_mach) * from mpp_gen " _
        & " where mpp_doc_no ='" & PmppDoc & "' and mpp_revisi='" & PmppRev & "' " _
        & " ) v1 order by no_mach asc) v2 inner join v_mesin_mater vmm on v2.no_mach=vmm.no_mach"

        Set rsB = Con.Execute(qry)

        ReDim aMesinInj(1 To rsB.RecordCount, 1 To totalHari + 3) As Variant
        i = 1
        While Not rsB.EOF
            aMesinInj(i, 1) = rsB(0)
            aMesinInj(i, 2) = IIf(IsNull(rsB(1)), " ", rsB(1))
            aMesinInj(i, 3) = IIf(IsNull(rsB(2)), " ", rsB(2))

            For k = 4 To totalHari + 3
                aMesinInj(i, k) = "0"
            Next
            i = 1 + i
            rsB.MoveNext
        Wend

        HKWs = RsA("ml_hkw")
        SkinLabel4.Caption = "HKW : " & HKWs
        Erase in_PartOL
        Erase in_PartOLQTY
        Erase in_partLTPP
        Erase in_partFC
        i = 1
        While Not RsA.EOF
            inAddValue RsA("lcd_itemdid"), 0, 0, 0
            i = 1 + i
            RsA.MoveNext
        Wend
        
        i = 2
        anaGrid.rows = 3
        anaGrid.FixedRows = 2
        RsA.MoveFirst
        
        Set faMachine = RsA.Fields.Item("no_mach")
        Set faCustomer = RsA.Fields.Item("lc_customer")
        Set faItemId = RsA.Fields.Item("lcd_itemdid")
        Set faItemName = RsA.Fields.Item("partname")
        Set faMold = RsA.Fields.Item("reg_mold")
        Set faColor = RsA.Fields.Item("colordesc")
        Set faCavity = RsA.Fields.Item("cav")
        Set faCT = RsA.Fields.Item("ct")
        Set faNeedMP = RsA.Fields.Item("needmp")
        Set faCapPHour = RsA.Fields.Item("cap_p_hour")
        Set faCapPShift = RsA.Fields.Item("cap_p_shift")
        Set faCapPDay = RsA.Fields.Item("cap_p_day")
        Set faProdPlan = RsA.Fields.Item("lc_pp")
        Set faMPSProdPlan = RsA.Fields.Item("mpp_pp")
        Set faFC = RsA.Fields.Item("fc")
        
        With anaGrid
            .rows = 2
            .rows = 2 + (ttlActMCH * 3) + RsA.RecordCount
            While Not RsA.EOF
                If temp_mch = faMachine.Value Then
                    noItemPerMesin = noItemPerMesin + 1
                    i = i + 1
                Else
                    noItemPerMesin = 1
                    If i > 2 Then
                        i = i + 4
                    Else
                        i = i + 0
                    End If
                End If
               
                .TextMatrix(i, 0) = faMachine.Value
                .TextMatrix(i, 1) = faCustomer.Value
                .TextMatrix(i, 2) = noItemPerMesin
                .TextMatrix(i, 3) = faItemId.Value
                .TextMatrix(i, 4) = faItemName.Value
                .Col = 5:    .Row = i
                .CellAlignment = flexAlignLeftCenter
                .TextMatrix(i, 5) = faMold.Value
                .TextMatrix(i, 6) = faColor.Value
                .TextMatrix(i, 7) = faCavity.Value
                .TextMatrix(i, 8) = faCT.Value
                .TextMatrix(i, 9) = FormatNumber(faNeedMP.Value, 1)
                If RsA("ct") > 0 Then
                    .TextMatrix(i, 10) = FormatNumber((3600 / faCT.Value) * faCavity.Value, 0) 'IIf(IsNull(faCapPHour.Value), 0, FormatNumber(faCapPHour.Value, 0))
                    .TextMatrix(i, 11) = (.TextMatrix(i, 10) * 1) * RsA("hourpshift")  'IIf(IsNull(faCapPShift.Value), 0, FormatNumber(faCapPShift.Value, 0))
                Else
                    .TextMatrix(i, 10) = 0
                    .TextMatrix(i, 11) = 0
                End If
                .TextMatrix(i, 12) = FormatNumber(faCapPDay.Value, 0)
                .TextMatrix(i, 16) = FormatNumber(faProdPlan.Value, 0)
                .TextMatrix(i, 17) = FormatNumber(faMPSProdPlan.Value, 0)
                .TextMatrix(i, 19) = FormatNumber(faFC.Value, 0)
                temp_mch = faMachine.Value
                RsA.MoveNext
            Wend
            'MsgBox "mulai"
'            While Not RsA.EOF
'                If temp_mch = faMachine.Value Then  ' rsA("no_mach")
'                    noItemPerMesin = noItemPerMesin + 1
'                Else
'                    noItemPerMesin = 1
'                End If
'                 .TextMatrix(i, 0) = faMachine.Value ' rsA("no_mach")
'                 .TextMatrix(i, 1) = faCustomer.Value ' rsA("lc_customer")
'                 .TextMatrix(i, 2) = noItemPerMesin
'                 .TextMatrix(i, 3) = faItemId.Value ' rsA("lcd_itemdid")
'                 .TextMatrix(i, 4) = faItemName.Value ' rsA("partname")
'                 .Col = 5:    .Row = i
'                 .CellAlignment = flexAlignLeftCenter
'                 .TextMatrix(i, 5) = faMold.Value 'rsA("reg_mold")
'                 .TextMatrix(i, 6) = faCavity.Value ' rsA("cav")
'                 .TextMatrix(i, 7) = faCT.Value ' rsA("ct")
'                 .TextMatrix(i, 8) = FormatNumber(faNeedMP.Value, 1) 'FormatNumber(rsA("needmp"), 1)
'                If RsA("ct") > 0 Then
'                    .TextMatrix(i, 9) = IIf(IsNull(faCapPHour.Value), 0, faCapPHour.Value) ' IIf(IsNull(rsA("cap_p_hour")), 0, rsA("cap_p_hour"))
'                    .TextMatrix(i, 9) = FormatNumber(.TextMatrix(i, 9), 0)
'                    .TextMatrix(i, 10) = IIf(IsNull(faCapPShift.Value), 0, faCapPShift.Value)  'IIf(IsNull(rsA("cap_p_shift")), 0, rsA("cap_p_shift"))
'                    .TextMatrix(i, 10) = FormatNumber(.TextMatrix(i, 10), 0)
'                Else
'                    .TextMatrix(i, 9) = 0
'                    .TextMatrix(i, 10) = 0
'                End If
'                .TextMatrix(i, 11) = FormatNumber(faCapPDay.Value, 0) 'FormatNumber(rsA("cap_p_day"), 0)
'                .TextMatrix(i, 15) = FormatNumber(faProdPlan.Value, 0) 'FormatNumber(rsA("lc_pp"), 0)
'                .TextMatrix(i, 16) = FormatNumber(faMPSProdPlan.Value, 0) 'FormatNumber(rsA("mpp_pp"), 0)
'                .TextMatrix(i, 18) = FormatNumber(faFC.Value, 0) 'FormatNumber(rsA("fc"), 0)
'                RsA.MoveNext
'                If RsA.EOF Then
'
'                Else
'                    temp_mch = faMachine.Value 'rsA("no_mach")
'                End If
'                RsA.MovePrevious
'                If temp_mch = faMachine.Value Then 'rsA("no_mach")
'                    .rows = .rows + 1
'                    i = i + 1
'                Else
'                    .rows = .rows + 4
'                    i = i + 4
'                End If
'
'                temp_mch = faMachine.Value 'rsA("no_mach")
'
'                RsA.MoveNext
'            Wend
'            .rows = .rows + 1
        End With
        RsA.AbsolutePosition = 1
    End If
    anaGrid.Refresh
    Set RsA = Nothing
End Sub

Private Sub fgmpp_RowColChange()
    With fgmpp
        NoDocMPS = .TextMatrix(.Row, 1)
        rev_MPS = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub flxsh_Click()
    With flxsh
        If .RowSel > 1 Then
            If .TextMatrix(.RowSel, .ColSel) = "q" Then
                .TextMatrix(.RowSel, .ColSel) = "þ"
                anaGrid.ColWidth(.ColSel) = .TextMatrix(1, .ColSel)
            Else
                .TextMatrix(.RowSel, .ColSel) = "q"
                anaGrid.ColWidth(.ColSel) = 0
            End If
        End If
    End With
End Sub

Private Sub flxsh_KeyPress(KeyAscii As Integer)
'    With flxsh
        If KeyAscii = 32 Then
            flxsh_Click
'            If .Text = "q" Then
'                .Text = "þ"
'                .CellForeColor = RGB(200, 0, 0)
'            Else
'                .Text = "R"
'                .CellForeColor = RGB(42, 127, 255)
'            End If
        End If
'    End With
End Sub

Private Sub Form_Deactivate()
    MDI_Parent.mnuFreezColumn.Visible = False
    MDI_Parent.mnuFontSize.Visible = False
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
    MDI_Parent.mnuFreezColumn.Visible = True
    MDI_Parent.mnuFontSize.Visible = True
End Sub

Private Sub StopTabs()
    On Error Resume Next
    Dim Ctrlku As Control
    For Each Ctrlku In Me.Controls
        Ctrlku.TabStop = False
    Next
    
End Sub

Private Sub Form_Load()
'    On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingFG
    Me.Height = 7755
    Me.Width = 14640
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
    belumSimpan = True
    cmbType.ListIndex = 0
    Call WheelHook(Me.hwnd)
    cmdFileType.ListIndex = 0
    StopTabs
    loadmachine
End Sub

Private Sub loadmachine()
    qry = "SELECT COUNT(*) FROM loadcap_mst_mach "
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        ttlActMCH = RsBantu(0)
    End If
    Set RsBantu = Nothing
End Sub

Private Sub Form_LostFocus()
    MDI_Parent.mnuFreezColumn.Visible = False
    MDI_Parent.mnuFontSize.Visible = False
End Sub

Private Sub Form_Resize()
    ResizeControls
    CmbDocument.Left = SkinLabel5.Left
    CmbDocument.Top = SkinLabel3.Top
    txtRevision.Top = SkinLabel2.Top
    txtRevision.Left = SkinLabel5.Left
    cmbPeriod.Left = SkinLabel5.Left

    
    cmbType.Left = SkinLabel5.Left
    cmbType.Top = SkinLabel9.Top
    cmdFileType.Top = CmdExport.Top
    cmdFileType.Width = Label16.Width
    cmdFileType.Left = Label16.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Call WheelUnHook(Me.hwnd)
        DelTab Me
        Erase aSuggest
        MDI_Parent.mnuFreezColumn.Visible = False
        MDI_Parent.mnuFontSize.Visible = False
    End If
End Sub

Private Sub Label11_Click()
    PicListMPP.Visible = False
End Sub

Private Sub Label13_Click()
    If cuemd_print.Enabled Then
        pic_pp_or_p.Visible = False
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



Private Sub Label4_Click()
    PicTrial.Visible = False
End Sub

Private Sub clearIN()
    CmbDocument = ""
    txtRevision.Clear
    cmbPeriod.Clear
End Sub

Private Sub Label9_Click()
    PicEditedList.Visible = False
End Sub

Private Sub lblCLOSE_Click()
    picSuggest_detail.Visible = False
End Sub


Private Sub lvprintp_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox KeyCode
End Sub

Private Sub Option1_Click()
    StopTabs
End Sub

Private Sub Option1_LostFocus()
    StopTabs
End Sub

Private Sub Option2_Click()
'    PicEm.Visible = False 1
    txtEdit.Visible = False
    StopTabs
End Sub

Private Sub Option2_LostFocus()
    StopTabs
End Sub

Private Sub picSign_ovr_Click()
    If (Not aSuggest) <> -1 Then
        Dim sug As String
        Dim isug As Byte
        
        sug = "Mesin di bawah ini butuh waktu lebih : " & vbNewLine
        For isug = 1 To UBound(aSuggest)
            sug = sug & aSuggest(isug, 1) & "=" & aSuggest(isug, 2) & "%" & vbNewLine
        Next
        sug = sug & vbNewLine
        sug = sug & "Hal tersebut dikarenakan presentasi menu loading > 100%"
        txtSugDetail = sug
        picSuggest_detail.Visible = True
        picSuggest_detail.Top = Me.Height / 2 - picSuggest_detail.Height / 2
        picSuggest_detail.Left = Me.Width / 2 - picSuggest_detail.Width / 2
        Form_Resize
    End If
End Sub

Private Sub Timer1_Timer()
    If (Not aSuggest) <> -1 Then
        If picSign_ovr.BackColor <> RGB(255, 179, 28) Then
            picSign_ovr.BackColor = RGB(255, 179, 28)
        Else
            picSign_ovr.BackColor = RGB(255, 255, 255)
        End If
        picSign_ovr.ToolTipText = "we have some suggestions for you, " & vbNewLine & "click me"
    Else
        If picSign_ovr.BackColor <> RGB(20, 247, 153) Then
            picSign_ovr.BackColor = RGB(20, 247, 153) 'hijau
            picSign_ovr.ToolTipText = ":)"
        End If
    End If
End Sub

Private Sub reinitOverloadQty()
    inSetValueZero anaGrid.TextMatrix(anaGrid.Row, 3)
    With anaGrid
        For r = 2 To .rows - 1
            If IsNumeric(.TextMatrix(r, 16)) Then '15
                If .TextMatrix(r, 16) * 1 > 0 Then '15
                    inSetValue .TextMatrix(r, 3), .TextMatrix(r, 16) * 1, .TextMatrix(r, 19) * 1 '15 18
                End If
            End If
        Next
        For r = 2 To .rows - 1
            If LenB(.TextMatrix(r, 1)) <> 0 Then
                .TextMatrix(r, 18) = FormatNumber(inGetValue(.TextMatrix(r, 3)), 0) '17
            End If
        Next
    End With
End Sub

Private Sub editValue(gridChoosed As MSFlexGrid)  'VBFlexGrid
On Error GoTo eXku
    gridChoosed.Text = FormatNumber(txtEdit.Text, 0)
'    PicEm.Visible = False
    txtEdit.Visible = False
    gridChoosed.SetFocus
    Dim kolit As Byte
    If txtEdit <> em_qtyBuf Then
        If gridChoosed.Name = "anaGrid" Then
            If belumSimpan = False Then
                Dim tgl As String
                Dim rowaf As Byte
                Dim lcvm As String
                Dim ostdpo As Double
                Dim forcas As Double
                Dim planddate As String
                
                tgl = "20" & Right(NoDocMPS, 2) & "-" & Mid(NoDocMPS, 12, 2) & "-" & Left(gridChoosed.TextMatrix(1, em_x), 2)
                qry = "update mpp_gen set planqty=" & txtEdit & "" _
                & " where mpp_doc_no='" & NoDocMPS & "' and no_mach='" & gridChoosed.TextMatrix(em_y, 0) & "' " _
                & " and lcd_itemdid='" & gridChoosed.TextMatrix(em_y, 3) & "' and reg_mold='" & gridChoosed.TextMatrix(em_y, 5) & "'" _
                & " and plandate='" & tgl & "'"
                Con.Execute qry, rowaf
                With gridChoosed
                    If rowaf = 0 Then
                        lcvm = 1 * Left$(.TextMatrix(em_y, (COLSDATE - 1)), Len(.TextMatrix(em_y, (COLSDATE - 1))) - 1)
                        ostdpo = IIf(IsNumeric(.TextMatrix(em_y, 15)), .TextMatrix(em_y, 15), 0) '14
                        planddate = Left$(cmbPeriod, 4) & "-" & Right$(cmbPeriod, 2) & "-" & Left$(.TextMatrix(1, em_x), 2)
                        If IsNumeric(.TextMatrix(em_y, 19)) Then '18
                            forcas = .TextMatrix(em_y, 19) * 1 '18
                        Else
                            forcas = 0
                        End If
                        qry = "INSERT INTO mpp_gen values (DEFAULT,'" & .TextMatrix(em_y, 3) & "'," _
                        & "'" & .TextMatrix(em_y, 4) & "'," _
                        & "'" & .TextMatrix(em_y, 1) & "'," _
                        & "'" & .TextMatrix(em_y, 0) & "'," _
                        & getTonage(.TextMatrix(em_y, 0)) & "," _
                        & "'" & .TextMatrix(em_y, 5) & "'," _
                        & .TextMatrix(em_y, 7) & "," & .TextMatrix(em_y, 8) & "," _
                        & .TextMatrix(em_y, 12) * 1 & "," _
                        & lcvm & ",'" & CmbDocument & "'," & txtRevision & "," _
                        & "'" & cmbPeriod & "'," & HKWs & ",'no'," _
                        & .TextMatrix(em_y, 17) * 1 & "," & .TextMatrix(em_y, 9) & "," _
                        & .TextMatrix(em_y, 10) * 1 & "," & .TextMatrix(em_y, 11) * 1 & "," _
                        & .TextMatrix(em_y, 13) * 1 & "," & .TextMatrix(em_y, 14) * 1 & "," _
                        & ostdpo & "," & forcas & ",'" & NoDocMPS & "','" & planddate & "'," & txtEdit & ",'" & rev_MPS & "')"
                        Con.Execute qry
                        qry = "SELECT rnomachine no_mach,ritemid lcd_itemdid,rreg_mold reg_mold," _
                        & " rplandate plandate,rplanqty planqty, rcav_std cav_std, rtypelabel typelabel, rct_scnd ct_scnd, " _
                        & " rlcneed_mp lcneed_mp, risno isno,rshiftusg shiftusg, rmpower mpower, " _
                        & " rhourpshift hourpshift, rtimeupdate timeupdate, rlc_pp lc_pp,rtypelabelbox typelabelbox FROM " _
                        & " f_loadcapdetail('" & CmbDocument & "'," & txtRevision & ",'" & cmbPeriod & "','" & NoDocMPS & "','" & rev_MPS & "','')"
                        Set rsA_aks = Con.Execute(qry)
                    End If
                End With
                
            End If
            
            hitungRekapINJ
            reinitOverloadQty
            kolit = 0
        Else
            hitungRekapSUB
            kolit = 1
        End If
        gridChoosed.Col = em_x
        gridChoosed.Row = em_y
        gridChoosed.SetFocus
        If belumSimpan = False Then
            Dim temp_tgl As String
            Dim temp_rindx As Byte
            With gridChoosed
                temp_tgl = Left$(cmbPeriod, 4) & "-" & Right$(cmbPeriod, 2) & "-" & Left$(.TextMatrix(1, em_x), 2)
                temp_rindx = checkPrimaryEL(.TextMatrix(em_y, 0), .TextMatrix(em_y, 3 - kolit), .TextMatrix(em_y, 5 + kolit), temp_tgl)
                If temp_rindx > 0 Then
                    If .Name = "anaGrid" Then
                        fge.TextMatrix(temp_rindx, 0) = .TextMatrix(em_y, 0)
                        fge.TextMatrix(temp_rindx, 1) = .TextMatrix(em_y, 1)
                        fge.TextMatrix(temp_rindx, 2) = .TextMatrix(em_y, 3)
                        fge.TextMatrix(temp_rindx, 3) = .TextMatrix(em_y, 4)
                        fge.TextMatrix(temp_rindx, 4) = .TextMatrix(em_y, 5)
                        fge.TextMatrix(temp_rindx, 5) = .TextMatrix(em_y, 7)
                        fge.TextMatrix(temp_rindx, 6) = .TextMatrix(em_y, 8)
                        fge.TextMatrix(temp_rindx, 7) = .TextMatrix(em_y, 9)
                        fge.TextMatrix(temp_rindx, 8) = .TextMatrix(em_y, 10)
                        fge.TextMatrix(temp_rindx, 9) = .TextMatrix(em_y, 11)
                        fge.TextMatrix(temp_rindx, 10) = .TextMatrix(em_y, 12)
                        fge.TextMatrix(temp_rindx, 11) = .TextMatrix(em_y, 13)
                        fge.TextMatrix(temp_rindx, 12) = .TextMatrix(em_y, 14)
                        fge.TextMatrix(temp_rindx, 13) = .TextMatrix(em_y, 15)
                        fge.TextMatrix(temp_rindx, 14) = .TextMatrix(em_y, 17)
                        fge.TextMatrix(temp_rindx, 15) = .TextMatrix(em_y, 19)
                        fge.TextMatrix(temp_rindx, 16) = temp_tgl
                        fge.TextMatrix(temp_rindx, 17) = .TextMatrix(em_y, em_x) * 1
                    Else
                        fge.TextMatrix(temp_rindx, 0) = .TextMatrix(em_y, 0)
                        fge.TextMatrix(temp_rindx, 1) = .TextMatrix(em_y, 1)
                        fge.TextMatrix(temp_rindx, 2) = .TextMatrix(em_y, 2)
                        fge.TextMatrix(temp_rindx, 3) = .TextMatrix(em_y, 3)
                        fge.TextMatrix(temp_rindx, 4) = .TextMatrix(em_y, 6)
                        fge.TextMatrix(temp_rindx, 5) = .TextMatrix(em_y, 7)
                        fge.TextMatrix(temp_rindx, 6) = .TextMatrix(em_y, 8)
                        fge.TextMatrix(temp_rindx, 7) = .TextMatrix(em_y, 9)
                        fge.TextMatrix(temp_rindx, 8) = .TextMatrix(em_y, 10)
                        fge.TextMatrix(temp_rindx, 9) = .TextMatrix(em_y, 11)
                        fge.TextMatrix(temp_rindx, 10) = .TextMatrix(em_y, 12)
                        fge.TextMatrix(temp_rindx, 11) = .TextMatrix(em_y, 13)
                        fge.TextMatrix(temp_rindx, 12) = .TextMatrix(em_y, 14)
                        fge.TextMatrix(temp_rindx, 13) = .TextMatrix(em_y, 15)
                        fge.TextMatrix(temp_rindx, 14) = .TextMatrix(em_y, 16)
                        fge.TextMatrix(temp_rindx, 15) = .TextMatrix(em_y, 17)
                        fge.TextMatrix(temp_rindx, 16) = temp_tgl
                        fge.TextMatrix(temp_rindx, 17) = .TextMatrix(em_y, em_x) * 1
                    End If
                Else
                    fge.rows = fge.rows + 1
                    If gridChoosed.Name = "anaGrid" Then
                        fge.TextMatrix(fge.rows - 1, 0) = .TextMatrix(em_y, 0)
                        fge.TextMatrix(fge.rows - 1, 1) = .TextMatrix(em_y, 1)
                        fge.TextMatrix(fge.rows - 1, 2) = .TextMatrix(em_y, 3)
                        fge.TextMatrix(fge.rows - 1, 3) = .TextMatrix(em_y, 4)
                        fge.TextMatrix(fge.rows - 1, 4) = .TextMatrix(em_y, 5)
                        fge.TextMatrix(fge.rows - 1, 5) = .TextMatrix(em_y, 7)
                        fge.TextMatrix(fge.rows - 1, 6) = .TextMatrix(em_y, 8)
                        fge.TextMatrix(fge.rows - 1, 7) = .TextMatrix(em_y, 9)
                        fge.TextMatrix(fge.rows - 1, 8) = .TextMatrix(em_y, 10)
                        fge.TextMatrix(fge.rows - 1, 9) = .TextMatrix(em_y, 11)
                        fge.TextMatrix(fge.rows - 1, 10) = .TextMatrix(em_y, 12)
                        fge.TextMatrix(fge.rows - 1, 11) = .TextMatrix(em_y, 13)
                        fge.TextMatrix(fge.rows - 1, 12) = .TextMatrix(em_y, 14)
                        fge.TextMatrix(fge.rows - 1, 13) = .TextMatrix(em_y, 15)
                        fge.TextMatrix(fge.rows - 1, 14) = .TextMatrix(em_y, 17)
                        fge.TextMatrix(fge.rows - 1, 15) = .TextMatrix(em_y, 19)
                        fge.TextMatrix(fge.rows - 1, 16) = temp_tgl
                        fge.TextMatrix(fge.rows - 1, 17) = .TextMatrix(em_y, em_x) * 1
                    Else
                        fge.TextMatrix(fge.rows - 1, 0) = .TextMatrix(em_y, 0)
                        fge.TextMatrix(fge.rows - 1, 1) = .TextMatrix(em_y, 1)
                        fge.TextMatrix(fge.rows - 1, 2) = .TextMatrix(em_y, 2)
                        fge.TextMatrix(fge.rows - 1, 3) = .TextMatrix(em_y, 3)
                        fge.TextMatrix(fge.rows - 1, 4) = .TextMatrix(em_y, 6)
                        fge.TextMatrix(fge.rows - 1, 5) = .TextMatrix(em_y, 7)
                        fge.TextMatrix(fge.rows - 1, 6) = .TextMatrix(em_y, 8)
                        fge.TextMatrix(fge.rows - 1, 7) = .TextMatrix(em_y, 9)
                        fge.TextMatrix(fge.rows - 1, 8) = .TextMatrix(em_y, 10)
                        fge.TextMatrix(fge.rows - 1, 9) = .TextMatrix(em_y, 11)
                        fge.TextMatrix(fge.rows - 1, 10) = .TextMatrix(em_y, 12)
                        fge.TextMatrix(fge.rows - 1, 11) = .TextMatrix(em_y, 13)
                        fge.TextMatrix(fge.rows - 1, 12) = .TextMatrix(em_y, 14)
                        fge.TextMatrix(fge.rows - 1, 13) = .TextMatrix(em_y, 15)
                        fge.TextMatrix(fge.rows - 1, 14) = .TextMatrix(em_y, 16)
                        fge.TextMatrix(fge.rows - 1, 15) = .TextMatrix(em_y, 17)
                        fge.TextMatrix(fge.rows - 1, 16) = temp_tgl
                        fge.TextMatrix(fge.rows - 1, 17) = .TextMatrix(em_y, em_x) * 1
                    End If
                End If
            End With
        End If
    End If
    Exit Sub
eXku:
    Clipboard.Clear
    Clipboard.SetText qry
    MsgBox Err.Description & vbNewLine & qry
End Sub


Private Sub CoordinateMouse()
    Dim Ret As Long
    
'Restituisce la posizione x,y relativamente allo schermo:
    Ret = GetCursorPos(pa)
'Converte la posizione x,y relativamente al form specificato(.hWnd):
    ScreenToClient Me.hwnd, pa
'Le due funzioni, GetCursorPos e ScreenToClient, restituiscono la
'posizione del mouse con valori espressi in pixel. Per convertire i valori in Twip :
    mosX = pa.x '* Screen.TwipsPerPixelX
    mosY = pa.Y '* Screen.TwipsPerPixelY
End Sub

Private Sub Timer2_Timer()
    CoordinateMouse
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(txtEdit) Then
            If cmbType.Text = "Machine Inj" Then
                editValue anaGrid
            Else
                 editValue anaSubcont
            End If
        Else
            MsgBox "the number is not valid", vbExclamation
        End If
    ElseIf KeyAscii = vbKeyEscape Then
'        PicEm.Visible = False '
        txtEdit.Visible = False
        If cmbType.Text = "Machine Inj" Then
            anaGrid.SetFocus
        ElseIf cmbType.Text = "Subcont" Then
            anaSubcont.SetFocus
        End If
    ElseIf KeyAscii = vbKeyTab Then

        If IsNumeric(txtEdit) Then
            If cmbType.Text = "Machine Inj" Then
                editValue anaGrid
            Else
                 editValue anaSubcont
            End If
        Else
            MsgBox "the number is not valid", vbExclamation
        End If
        If anaGrid.Col < anaGrid.Cols - 1 Then
            anaGrid.Col = anaGrid.Col + 1
        End If
        anaGrid.SetFocus
    End If
End Sub

Private Function checkPrimaryEL(p1_mc As String, p2_part As String, p3_mold As String, p4_tgl As String) As Byte
    Dim ui As Byte
    With fge
        For ui = 1 To .rows - 1
            If p1_mc = .TextMatrix(ui, 0) And p2_part = .TextMatrix(ui, 2) And p3_mold = .TextMatrix(ui, 4) And p4_tgl = .TextMatrix(ui, 16) Then
                checkPrimaryEL = ui
                Exit Function
            End If
        Next
    End With
    checkPrimaryEL = 0
End Function

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim ttlData As Long
        txtFind = FilterIn(txtFind)
        qry = "select * from (select distinct on (mpp_doc_no,mpp_revisi,ml_doc,ml_rev) mpp_doc_no,mpp_revisi,ml_ym,ml_rev,ml_doc  from mpp_gen where mpp_doc_no like '%" & txtFind & "%' ) v1 order by ml_ym desc, mpp_revisi desc limit 3"
        Set RsBantu = Con.Execute(qry)
        
        fgmpp.rows = 1
        If RsBantu.RecordCount > 0 Then
            RsBantu.Fields("mpp_doc_no").Properties("Optimize") = True
            RsBantu.Fields("mpp_revisi").Properties("Optimize") = True
            With fgmpp
                If Len(Trim(txtFind)) > 0 Then
                    RsBantu.Filter = adFilterNone
                    RsBantu.Filter = "mpp_doc_no LIKE '*" & txtFind & "*'"
                    If RsBantu.RecordCount > 0 Then
                        clearIN
                        .rows = RsBantu.RecordCount + 1
                        ttlData = RsBantu.RecordCount
                        For i = 1 To ttlData
                            RsBantu.AbsolutePosition = i
                            .TextMatrix(i, 0) = i
                            .TextMatrix(i, 1) = RsBantu("mpp_doc_no")
                            .TextMatrix(i, 2) = RsBantu("mpp_revisi")
                            .TextMatrix(i, 3) = RsBantu("ml_ym")
                            .TextMatrix(i, 4) = RsBantu("ml_rev")
                            .TextMatrix(i, 5) = RsBantu("ml_doc")
                            CmbDocument = RsBantu("ml_doc") 'CmbDocument.AddItem RsBantu("ml_doc")
                            txtRevision.AddItem RsBantu("ml_rev")
                            cmbPeriod.AddItem RsBantu("ml_ym")
                        Next
                    Else
                        RsBantu.Filter = adFilterNone
                        RsBantu.Filter = "mpp_revisi LIKE '*" & txtFind & "*'"
                        If RsBantu.RecordCount > 0 Then
                            clearIN
                            .rows = RsBantu.RecordCount + 1
                            ttlData = RsBantu.RecordCount
                            For i = 1 To ttlData
                                RsBantu.AbsolutePosition = i
                                .TextMatrix(i, 0) = i
                                .TextMatrix(i, 1) = RsBantu("mpp_doc_no")
                                .TextMatrix(i, 2) = RsBantu("mpp_revisi")
                                .TextMatrix(i, 3) = RsBantu("ml_ym")
                                .TextMatrix(i, 4) = RsBantu("ml_rev")
                                .TextMatrix(i, 5) = RsBantu("ml_doc")
                                CmbDocument = RsBantu("ml_doc") 'CmbDocument.AddItem RsBantu("ml_doc")
                                txtRevision.AddItem RsBantu("ml_rev")
                                cmbPeriod.AddItem RsBantu("ml_ym")
                            Next
                        Else
                            .rows = 1
                        End If
                    End If
                Else
                    .rows = RsBantu.RecordCount + 1
                    clearIN
                    ttlData = RsBantu.RecordCount
                    For i = 1 To ttlData
                        RsBantu.AbsolutePosition = i
                        .TextMatrix(i, 0) = i
                        .TextMatrix(i, 1) = RsBantu("mpp_doc_no")
                        .TextMatrix(i, 2) = RsBantu("mpp_revisi")
                        .TextMatrix(i, 3) = RsBantu("ml_ym")
                        .TextMatrix(i, 4) = RsBantu("ml_rev")
                        .TextMatrix(i, 5) = RsBantu("ml_doc")
                        CmbDocument = RsBantu("ml_doc") 'CmbDocument.AddItem RsBantu("ml_doc")
                        txtRevision.AddItem RsBantu("ml_rev")
                        cmbPeriod.AddItem RsBantu("ml_ym")
                    Next
                End If
            End With
        End If
    End If
End Sub

Private Sub txtFindNext_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    ElseIf KeyAscii = vbKeyEscape Then
        PicFIND.Visible = False
    ElseIf KeyAscii = 1 Then
        txtFindNext.SelStart = 0
        txtFindNext.SelLength = Len(txtFindNext.Text)
    End If
End Sub

Private Sub txtRevision_Click()
    cmbPeriod_DropDown
    
End Sub

Private Sub txtRevision_DropDown()
    If Len(CmbDocument) < 2 Then CmbDocument.SetFocus: Exit Sub
    qry = "select distinct on (fltpp_rev) fltpp_rev from mpp_gen_d where fltpp_doc='" & CmbDocument & "'"
    Set RsA = Con.Execute(qry)
    txtRevision.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            txtRevision.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub

Private Function checklotPerItemid(partId As String, tgl As Byte) As Boolean
    Dim xb As Long
    Dim xc As Long
    PGCheckLot.Value = 0
    PGCheckLot.Visible = True
    PGCheckLotLabel.Visible = True
    With anaGrid
        
        For xb = 2 To .rows - 2
            If .TextMatrix(xb, 3) = partId And IsNumeric(.TextMatrix(xb, 20)) Then '19
                If .TextMatrix(xb, 20) * 1 > 0 Then '19
                    For xc = COLSDATE To .Cols - 1
                        .Row = xb
                        .Col = xc
                        If IsNumeric(.Text) Then
                        If .CellBackColor <> RGB(WO_R, WO_G, WO_B) And .Text * 1 > 0 Then
                            If .CellBackColor <> RGB(172.38, 233.51, 235.62) Then
                                If Left$(.TextMatrix(1, xc), 2) * 1 < tgl Then
                                    MsgBox "ada wo yang belum dicetak " & .Text
                                    hideProgresscheckLot
                                    .TopRow = xb
                                    checklotPerItemid = True
                                    Exit Function
                                End If
                            End If
                        End If
                        End If
                    Next
                End If
            End If
            DoEvents
            For xc = COLSDATE To .Cols - 1
                If LenB(.TextMatrix(xb, 3)) <> 0 Then
                    .Row = xb
                    .Col = xc
                    If IsNumeric(.Text) Then
                    If .CellBackColor <> RGB(WO_R, WO_G, WO_B) And .Text * 1 > 0 Then
                        If .CellBackColor = RGB(172.38, 233.51, 235.62) Then
                            For r = xc To COLSDATE Step -1
                                .Col = r
                                If .CellBackColor <> RGB(172.38, 233.51, 235.62) Then
                                    If .CellBackColor <> RGB(WO_R, WO_G, WO_B) And IsNumeric(.Text) Then
                                        If .Text * 1 > 0 Then
                                            MsgBox "ada wo yang belum dicetak dengan Assy " & .TextMatrix(xb, 3) 'KRIUK
                                            .TopRow = xb
                                            checklotPerItemid = True
                                            hideProgresscheckLot
                                        Exit Function
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                    End If
                End If
            Next
            PGCheckLot.Value = FormatNumber(((xb - 1) * 100) / (.rows - 2), 0)
            PGCheckLot.ToolTipText = PGCheckLot.Value & "%"
        Next
    End With
    hideProgresscheckLot
    checklotPerItemid = False
End Function

Private Sub hideProgresscheckLot()
    PGCheckLot.Visible = False
    PGCheckLotLabel.Visible = False
End Sub

Private Function getlot_check(ItemID As String) As String
    Dim tempS As Variant
    Dim FMY As String
    Dim ft_month As String
    Dim ft_year As String
    ft_year = Mid$(cmbPeriod, 3, 2)
    ft_month = Right$(cmbPeriod, 2)
    'Right$(cmbPeriod, 2) & " " & Mid$(cmbPeriod, 3, 2)
    'FMY = " " & Mid$(curLot, 3, 2) & " " & Right$(curLot, 2)
    FMY = " " & ft_month & " " & ft_year
    'qry = "select lotno from worko where partno='" & itemId & "' and " _
     & " substring(lotno from 3 for 2)='" & Mid$(curLot, 3, 2) & "' and substring(lotno from 5 for 2)='" & Right$(curLot, 2) & "' " _
    & " order by substring(lotno from 5 for 2) desc, substring(lotno from 3 for 2) desc ,substring(lotno from 1 for 2) desc limit 1"
    qry = "select lotno from worko where partno='" & ItemID & "' and " _
     & " substring(lotno from 3 for 2)='" & ft_month & "' and substring(lotno from 5 for 2)='" & ft_year & "' " _
    & " order by substring(lotno from 5 for 2) desc, substring(lotno from 3 for 2) desc ,substring(lotno from 1 for 2) desc limit 1"
    
    Set RsBantu = Con.Execute(qry)
    If RsBantu.RecordCount > 0 Then
        tempS = Left(RsBantu(0), 2)
        tempS = tempS * 1 + 1
        If Len(tempS) = 1 Then
            getlot_check = "0" & tempS & FMY
        Else
            getlot_check = tempS & FMY
        End If
    Else
        getlot_check = "01" & FMY
    End If
End Function

Private Sub getWO(pPrintPP As Boolean, ygd As Double, xgd As Double)
On Error GoTo Exc
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim q As Integer
    Dim QRYane As String
    Dim sqlRep As String
    Dim lot As String
    
    '========save
    Dim it_issue As String
    Dim it_partno As String
    Dim it_partnm As String
    Dim it_mold As String
    Dim it_mesin As String
    Dim it_qty As Long
    Dim it_cavstd As String
    Dim it_cav As String
    Dim it_ct1 As Single
    Dim it_ct2 As Single
    Dim it_targetshift As Long
    Dim it_mp As Single
    Dim it_leadtime As Single
    Dim it_datesupply As String
    Dim it_isNO As String
    Dim it_label As String
    Dim it_labelbox As String
    Dim it_WOno As String
    Dim it_qtyreq As Single
    Dim it_ttlqtyreq As Single
    Dim it_reqqtypurg As Single
    Dim it_purg As Single
    Dim qryinside As String
    Dim it_cappday As Long
    Dim it_pp As Long
    '========end save
    
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
        rsBOM.Filter = "bom_par_item='" & anaGrid.TextMatrix(anaGrid.Row, 3) & "'"
        'lot = getLot(anaGrid.Row, anaGrid.Col)
        it_issue = Left$(cmbPeriod, 4) & "-" & Right$(cmbPeriod, 2) & "-" & Left$(anaGrid.TextMatrix(1, xgd), 2)
        it_pp = anaGrid.TextMatrix(anaGrid.Row, 16) * 1 '15
        it_partno = anaGrid.TextMatrix(anaGrid.Row, 3)
        lot = getlot_check(it_partno)
        
        it_partnm = anaGrid.TextMatrix(anaGrid.Row, 4)
        it_mold = anaGrid.TextMatrix(anaGrid.Row, 5)
        it_mesin = anaGrid.TextMatrix(anaGrid.Row, 0)
        
        rsPurg.Filter = "no_mach='" & it_mesin & "'"
        If rsPurg.RecordCount > 0 Then
            it_purg = rsPurg("qtypurg")
        Else
            it_purg = 0
        End If
                
        it_cav = anaGrid.TextMatrix(anaGrid.Row, 7) '6
        it_qty = anaGrid.TextMatrix(anaGrid.Row, anaGrid.Col)
        it_cavstd = getCavStd(anaGrid.TextMatrix(anaGrid.Row, 0), anaGrid.TextMatrix(anaGrid.Row, 5), anaGrid.TextMatrix(anaGrid.Row, 3), it_issue)
        it_ct2 = getCTscnd(anaGrid.TextMatrix(anaGrid.Row, 0), anaGrid.TextMatrix(anaGrid.Row, 5), anaGrid.TextMatrix(anaGrid.Row, 3)) * 1
        it_ct1 = anaGrid.TextMatrix(anaGrid.Row, 8) * 1 '7
        it_targetshift = anaGrid.TextMatrix(anaGrid.Row, 11) * 1 '10
        'anaGrid.TextMatrix(anaGrid.Row, 0)
        qryinside = "no_mach='" & anaGrid.TextMatrix(ygd, 0) & "' and lcd_itemdid='" & anaGrid.TextMatrix(ygd, 3) _
                            & "' and reg_mold='" & anaGrid.TextMatrix(ygd, 5) & "'"
        rsA_aks.Filter = qryinside
        it_mp = rsA_aks("mpower")
        it_leadtime = FormatNumber(anaGrid.TextMatrix(anaGrid.Row, anaGrid.Col) / (3600 / anaGrid.TextMatrix(anaGrid.Row, 8) * Int(it_cavstd)), 2)
        it_datesupply = Format(Now, "yyyy\-MM\-dd") '^7
        it_isNO = getISNO(anaGrid.TextMatrix(anaGrid.Row, 0), anaGrid.TextMatrix(anaGrid.Row, 5), anaGrid.TextMatrix(anaGrid.Row, 3))
        it_label = getTypeLabel(anaGrid.TextMatrix(anaGrid.Row, 0), anaGrid.TextMatrix(anaGrid.Row, 5), anaGrid.TextMatrix(anaGrid.Row, 3))
        it_labelbox = getTypeLabelBox(anaGrid.TextMatrix(anaGrid.Row, 0), anaGrid.TextMatrix(anaGrid.Row, 5), anaGrid.TextMatrix(anaGrid.Row, 3))
        
        it_WOno = getWOno
        it_cappday = anaGrid.TextMatrix(anaGrid.Row, 12) * 1 '11
        .Sections("Section4").Controls("lblheaderpart").Caption = it_partno
        With .Sections("Section2")
            GenerateCode128 it_WOno
            .Controls("lblIssuedDate").Caption = anaGrid.TextMatrix(1, anaGrid.Col) & "-" & Left$(cmbPeriod, 4)
            .Controls("lblwarna").Caption = anaGrid.TextMatrix(ygd, 6)
            .Controls("lbl_k_issued").Caption = .Controls("lblIssuedDate").Caption
            .Controls("lblPartNo").Caption = it_partno
            .Controls("lbl_k_partno").Caption = .Controls("lblPartNo").Caption
            .Controls("lblPartNM").Caption = it_partnm
            .Controls("lbl_k_partnm").Caption = .Controls("lblPartNM").Caption
            .Controls("lblMoldNo").Caption = it_mold
            .Controls("lblMesinNo").Caption = it_mesin
            .Controls("lbl_k_mesin").Caption = .Controls("lblMesinNo").Caption
            .Controls("lblLOT").Caption = lot
            .Controls("lbl_k_lot").Caption = .Controls("lblLOT").Caption
            .Controls("lblQTY").Caption = it_qty
            .Controls("lblcavitystd").Caption = it_cavstd
            .Controls("lblctfinishing").Caption = it_ct2
            .Controls("lblctmachine").Caption = it_ct1
            .Controls("lbltargetshift").Caption = it_targetshift
            .Controls("lblmanpower").Caption = it_mp
            .Controls("lblLeadTime").Caption = it_leadtime
            .Controls("lbldatesupply").Caption = Format(Now, "dd\-MMM\-yyyy")
            .Controls("lblISno").Caption = it_isNO
            .Controls("lblmanual").Caption = it_label
            .Controls("lblmanualbox").Caption = it_labelbox
            .Controls("lblNodoc").Caption = it_WOno
            .Controls("lbl_k_nodoc").Caption = .Controls("lblNodoc").Caption
            .Controls("lbltimeupdate").Caption = timeupdate
            .Controls("lblcavityact").Caption = it_cav
            Set .Controls("Image1").Picture = LoadPicture(App.Path & "\Templates\com.bmp")
            Set .Controls("Image3").Picture = LoadPicture(App.Path & "\Templates\comr.bmp")
            
            it_ttlqtyreq = 0
            it_qtyreq = 0
            'reinit_material

            .Controls("lbl_k_matid").Caption = ""
            .Controls("lbl_m_matid").Caption = ""
            .Controls("lbl_k_matvir_nm").Caption = ""
            .Controls("lblqtyReq").Caption = ""
            .Controls("lblqtyReq_m").Caption = ""
            
            For r = 1 To rsBOM.RecordCount
                rsBOM.AbsolutePosition = r
                .Controls("lbl_k_matid").Caption = .Controls("lbl_k_matid").Caption & vbNewLine & rsBOM("bom_com_item")
                .Controls("lbl_k_matvir_nm").Caption = .Controls("lbl_k_matvir_nm").Caption & vbNewLine & rsBOM("item_name")

                it_qtyreq = it_qty * rsBOM("bom_qty_perassy")
                it_ttlqtyreq = it_ttlqtyreq + it_qtyreq
                
            Next
            .Controls("lbl_m_matid").Caption = .Controls("lbl_k_matid").Caption
            For r = 1 To rsBOM.RecordCount
                rsBOM.AbsolutePosition = r
                it_qtyreq = it_qty * rsBOM("bom_qty_perassy")
                it_reqqtypurg = (it_qtyreq / it_ttlqtyreq) * (it_purg / 1000)
                .Controls("lblqtyReq").Caption = .Controls("lblqtyReq").Caption & vbNewLine & (it_qtyreq + it_reqqtypurg) & " " & rsBOM("um_name")
            Next
            .Controls("lblqtyReq_m").Caption = .Controls("lblqtyReq").Caption
        End With
        
        .Refresh
        Me.MousePointer = vbDefault
        lot = Replace(lot, " ", "")
        If cekWODBL(it_partno, it_mesin, it_mold, it_issue) Then
            MsgBox "we have blocked that data, please contact admin"
            Exit Sub
        End If
        If pPrintPP Then
            .PrintReport False, rptRangeAllPages
            qry = "insert into worko (wo_no,status,lotno,issudate,partno,moldno, " _
            & "mesinno,qty,cavstd,ctscnd,ctmachine,targetpshift,manpower,leadtime " _
            & ",datesupply,isno,tipelabel,matreq,mpp_doc,mpprev " _
            & ",cappday,mpp_pp,printdate,userprint,qty_prg,tipelabelboks) values('" & it_WOno & "','O','" _
            & lot & "','" & it_issue & "','" & it_partno & "','" & it_mold & "'," _
            & "'" & it_mesin & "'," & it_qty & "," & it_cavstd & "," & it_ct2 & "" _
            & "," & it_ct1 & "," & it_targetshift & "," & it_mp & "," & it_leadtime & "" _
            & ",'" & it_datesupply & "','" & it_isNO & "','" & it_label & "'" _
            & "," & it_qtyreq & "" _
            & ",'" & NoDocMPS & "','" & rev_MPS & "'," & it_cappday & "," & it_pp & "," _
            & "now(),'" & FilterIn(pUserName) & "'," & it_purg & ",'" & it_labelbox & "')"
            Con.Execute qry
            
            For k = 1 To rsBOM.RecordCount
                rsBOM.AbsolutePosition = k
                qry = "insert into worko_mat values ('" & it_WOno & "'" _
                & ",'" & RTrim(rsBOM("bom_com_item")) & "'" _
                & ",'" & RTrim(rsBOM("item_name")) & "','" & Trim(rsBOM("pfm_id")) & "'" _
                & "," & rsBOM("bom_qty_perassy") & ")"
                Con.Execute qry
            Next
        Else
            .Show
        End If
    End With
    Exit Sub
Exc:
    If Err.Number = 8542 Then
        MsgBox "Ukuruan lebar kertas yang dibutuhkan tidak memungkinkan, " & vbNewLine & " ganti tipe kertas atau Printer yang mendukung kertas A3", vbCritical, "Sorry " & Err.Number
        CommonDialog1.ShowPrinter
        MsgBox "Silahkan coba lagi", vbInformation
    Else
        MsgBox Err.Description & vbNewLine & it_partno & "_" & it_mesin & qryinside, vbCritical, "Error No. : " & Err.Number
    End If
End Sub

Function cekWODBL(pPART As String, pMCH As String, pMOLD, pTGL As String) As Boolean
    Dim qqry As String
    Dim rsAB As New ADODB.Recordset
    Dim ToReturn As Boolean
    ToReturn = False
    qqry = "SELECT count(*) from worko where partno='" & pPART & "' and mesinno='" & pMCH & "' and moldno='" & pMOLD & "' and issudate='" & pTGL & "'"
    Set rsAB = Con.Execute(qqry)
    If rsAB(0) > 0 Then
        ToReturn = True
    End If
    cekWODBL = ToReturn
End Function

Private Function Check_dbl_WO(pPART As String, pdate As String, pmesin As String) As Boolean
    qry = "SELECT count(*) FROM worko WHERE partno='" & pPART & "' AND issudate='" & pdate & "' AND mesinno='" & pmesin & "'"
    Set RsTemp = Con.Execute(qry)
    If RsTemp(0) > 0 Then
        Check_dbl_WO = True
    Else
        Check_dbl_WO = False
    End If
    Set RsTemp = Nothing
End Function

Private Function getLot(yGrid As Long, xGrid As Integer) As String
    Dim konter As Integer
    konter = 0
    For i = 21 To xGrid
        If IsNumeric(anaGrid.TextMatrix(yGrid, i)) Then
            If anaGrid.TextMatrix(yGrid, i) > 0 Then
                konter = konter + 1
            End If
        End If
    Next
    If konter > 9 Then
        getLot = konter & " " & Right$(cmbPeriod, 2) & " " & Mid$(cmbPeriod, 3, 2)
    Else
        getLot = "0" & konter & " " & Right$(cmbPeriod, 2) & " " & Mid$(cmbPeriod, 3, 2)
    End If
End Function

Private Function getCavStd(pmesin As String, pMOLD As String, pPART As String, ptgal As String) As String
    rsA_aks.Fields("no_mach").Properties("Optimize") = True
    rsA_aks.Fields("lcd_itemdid").Properties("Optimize") = True
    rsA_aks.Fields("reg_mold").Properties("Optimize") = True
    rsA_aks.Fields("plandate").Properties("Optimize") = True
    
    rsA_aks.Filter = adFilterNone
    rsA_aks.Filter = "no_mach='" & pmesin & "' and reg_mold='" & pMOLD & "' and lcd_itemdid='" & pPART & "' and plandate='" & ptgal & "'"
    If rsA_aks.RecordCount > 0 Then
        getCavStd = rsA_aks("cav_std")
        timeupdate = IIf(IsNull(rsA_aks("timeupdate")), "", "Last update : " & Format(rsA_aks("timeupdate"), "dd MMM yyyy HH:mm"))
    End If
End Function

Private Function getCTscnd(pmesin As String, pMOLD As String, pPART As String) As String
    rsA_aks.Filter = adFilterNone
    rsA_aks.Filter = "no_mach='" & pmesin & "' and reg_mold='" & pMOLD & "' and lcd_itemdid='" & pPART & "'"
    If rsA_aks.RecordCount > 0 Then
        getCTscnd = rsA_aks("ct_scnd")
    End If
End Function

Private Function getMP(pmesin As String, pMOLD As String, pPART As String) As Variant
    rsA_aks.Filter = adFilterNone
    rsA_aks.Filter = "no_mach='" & pmesin & "' and reg_mold='" & pMOLD & "' and lcd_itemdid='" & pPART & "'"
    If rsA_aks.RecordCount > 0 Then
        getMP = ceiling(rsA_aks("lcneed_mp"))
    End If
End Function

Private Function getISNO(pmesin As String, pMOLD As String, pPART As String) As String
    rsA_aks.Filter = adFilterNone
    rsA_aks.Filter = "no_mach='" & pmesin & "' and reg_mold='" & pMOLD & "' and lcd_itemdid='" & pPART & "'"
    If rsA_aks.RecordCount > 0 Then
        getISNO = rsA_aks("isno")
    End If
End Function

Private Function getTypeLabel(pmesin As String, pMOLD As String, pPART As String) As String
    rsA_aks.Filter = adFilterNone
    rsA_aks.Filter = "no_mach='" & pmesin & "' and reg_mold='" & pMOLD & "' and lcd_itemdid='" & pPART & "'"
    If rsA_aks.RecordCount > 0 Then
        If IsNull(rsA_aks!typelabel) Then
            getTypeLabel = ""
        Else
'        Label Manual Logo BPI
'Label Manual Logo ASKARA
'Tidak Pakai Label Manual

            If rsA_aks!typelabel = 0 Then
                getTypeLabel = "Label Manual Logo BPI"
            ElseIf rsA_aks!typelabel = 1 Then
                getTypeLabel = "Label Manual Logo ASKARA"
            ElseIf rsA_aks!typelabel = 2 Then
                getTypeLabel = "Tidak Pakai Label Manual"
            End If
        End If
    End If
End Function

Private Function getTypeLabelBox(pmesin As String, pMOLD As String, pPART As String) As String
    rsA_aks.Filter = adFilterNone
    rsA_aks.Filter = "no_mach='" & pmesin & "' and reg_mold='" & pMOLD & "' and lcd_itemdid='" & pPART & "'"
    If rsA_aks.RecordCount > 0 Then
        If IsNull(rsA_aks!typelabelbox) Then
            getTypeLabelBox = ""
        Else
            If rsA_aks!typelabelbox = 0 Then
                getTypeLabelBox = "Label Manual"
            ElseIf rsA_aks!typelabelbox = 1 Then
                getTypeLabelBox = "Tidak Pakai Label Manual"
            End If
        End If
    End If
End Function

Private Function getWOno() As String
    qry = "SELECT left(wo_no,3) wono from worko where substring(wo_no from 9 for 2)='" & Right$(cmbPeriod, 2) & "' and substring(wo_no from 12 for 2)='" & Mid$(cmbPeriod, 3, 2) & "'" _
    & " order by substring(wo_no from 12 for 2)='" & Mid$(cmbPeriod, 3, 2) & "' desc,substring(wo_no from 9 for 2)='" & Right$(cmbPeriod, 2) & "' desc,left(wo_no,3) desc limit 1 "

    Set RsGet = Con.Execute(qry)
    Dim iCount As String
    If RsGet.RecordCount > 0 Then
        iCount = RsGet(0) * 1 + 1
        iCount = Right$("000" & iCount, 3)
        getWOno = iCount & "/PPC/" & Right$(cmbPeriod, 2) & "/" & Mid$(cmbPeriod, 3, 2)
    Else
        getWOno = "001/PPC/" & Right$(cmbPeriod, 2) & "/" & Mid$(cmbPeriod, 3, 2)
    End If
End Function

Private Sub hitungRekapINJ()
    ReDim ar_totallc(1 To 1) As Variant
    Dim sMch As String
    Dim ttlPerHari As Single
    Dim in_MP As Single
    Dim in_NeedDay As Single
    Dim in_NeedMP As Single
    Dim ttlrows As Long
    Dim ttlCOls As Byte
    RsA.Fields("no_mach").Properties("Optimize") = True
    RsA.Fields("lcd_itemdid").Properties("Optimize") = True
    RsA.Fields("reg_mold").Properties("Optimize") = True
    With anaGrid
        ttlrows = .rows - 1
        ttlCOls = .Cols - 1
        For i = 2 To ttlrows
            If Len(.TextMatrix(i, 0)) > 1 Then
                ttlMPP = 0
                For c = COLSDATE To ttlCOls
                    If IsNumeric(.TextMatrix(i, c)) Then
                        ttlMPP = ttlMPP + (.TextMatrix(i, c) * 1)
                    End If
                Next
                .TextMatrix(i, 20) = FormatNumber(ttlMPP, 0) '19

                RsA.Filter = "no_mach='" & .TextMatrix(i, 0) & "' and lcd_itemdid='" & .TextMatrix(i, 3) _
                            & "' and reg_mold='" & .TextMatrix(i, 5) & "'"
                If RsA.RecordCount > 0 Then
                    If .TextMatrix(i, 12) <> 0 Then '11
                        in_NeedDay = ttlMPP / .TextMatrix(i, 12) '11
                        in_NeedMP = in_NeedDay / HKWs * RsA("mpower")
                        
                        .TextMatrix(i, 9) = FormatNumber(in_NeedMP, 1) '8
                    End If
                End If
                If .TextMatrix(i, 12) = 0 Then 'cap per day '11
                    .TextMatrix(i, 14) = 0 '13
                Else
                    .TextMatrix(i, 14) = ttlMPP / .TextMatrix(i, 12) ' 13  11
                End If
                
                If .TextMatrix(i, 10) = 0 Then '9
                    .TextMatrix(i, 13) = 0 '12
                Else
                    .TextMatrix(i, 13) = ttlMPP / .TextMatrix(i, 10) '12 10
                End If
                If ttlMPP <> 0 Then
                    If .TextMatrix(i, 12) * 1 > 0 Then '11
                    .TextMatrix(i, (COLSDATE - 1)) = (ttlMPP / .TextMatrix(i, 12)) / HKWs * 100 '11
                    End If
                Else
                    .TextMatrix(i, (COLSDATE - 1)) = 0
                End If
                .TextMatrix(i, 13) = FormatNumber(.TextMatrix(i, 13), 2) '12
                .TextMatrix(i, 14) = FormatNumber(.TextMatrix(i, 14), 2) '13
                If IsNumeric(.TextMatrix(i, (COLSDATE - 1))) Then
                    .TextMatrix(i, (COLSDATE - 1)) = FormatNumber(.TextMatrix(i, (COLSDATE - 1)), 2) & "%"
                End If
            End If
            If LenB(.TextMatrix(i, 0)) = 0 And LenB(.TextMatrix(i - 1, 0)) <> 0 Then
                .TextMatrix(i, 1) = "Total"
                .TextMatrix(i + 1, (COLSDATE - 1)) = "MP"
                .Row = i: .Col = (COLSDATE - 1)
                st_MPP = 0
                st_CapDay = 0
                st_ReqHour = 0
                st_ReqDay = 0
                st_PP = 0
                st_ttlmpp = 0
                st_FC = 0
                ReDim Preserve ar_totallc(1 To UBound(ar_totallc) + 1)
                For r = i - 1 To 1 Step -1
                    If IsNumeric(.TextMatrix(r, 2)) Then   'jika berubah mesin
                        If LenB(.TextMatrix(r, (COLSDATE - 1))) <> 0 Then 'jika ada presentase
                            st_MPP = st_MPP + (Left$(.TextMatrix(r, (COLSDATE - 1)), Len(.TextMatrix(r, (COLSDATE - 1))) - 1) * 1)
                        End If
                        If LenB(.TextMatrix(r, 12)) <> 0 Then '11
                            st_CapDay = st_CapDay + .TextMatrix(r, 12) * 1 '11
                        End If
                        If LenB(.TextMatrix(r, 13)) <> 0 Then '12
                            st_ReqHour = st_ReqHour + .TextMatrix(r, 13) * 1 '12
                        End If
                        If LenB(.TextMatrix(r, 14)) <> 0 Then '13
                            st_ReqDay = st_ReqDay + .TextMatrix(r, 14) * 1 '13
                        End If
                        If LenB(.TextMatrix(r, 17)) <> 0 Then '16
                            st_PP = st_PP + .TextMatrix(r, 17) * 1 '16
                        End If
                        If LenB(.TextMatrix(r, 19)) <> 0 Then '18
                            st_FC = st_FC + .TextMatrix(r, 19) * 1 '18
                        End If
                        If LenB(.TextMatrix(r, 20)) <> 0 Then '19
                            st_ttlmpp = st_ttlmpp + .TextMatrix(r, 20) * 1 '19
                        End If
                    Else
                        Exit For
                    End If
                Next
                .TextMatrix(i, (COLSDATE - 1)) = st_MPP & "%"
                .TextMatrix(i, 20) = FormatNumber(st_ttlmpp, 0) '19
                .TextMatrix(i, 12) = FormatNumber(st_CapDay, 0) '11
                .TextMatrix(i, 13) = FormatNumber(st_ReqHour, 0) '12
                .TextMatrix(i, 14) = FormatNumber(st_ReqDay, 0) '13
                .TextMatrix(i, 17) = FormatNumber(st_PP, 0) '16
                .TextMatrix(i, 19) = FormatNumber(st_FC, 0) '18
                If st_MPP > 99 Then
                    .CellBackColor = RGB(255, 155, 155) ' merah muda
                ElseIf st_MPP > 80 Then
                    .CellBackColor = RGB(255, 255, 0) 'vbYellow
                ElseIf st_MPP > 0 Then
                    .CellBackColor = RGB(111, 255, 0) 'vbGreen
                End If
            End If
        Next
       
        For k = COLSDATE To ttlCOls
            For i = 2 To ttlrows
                If LenB(.TextMatrix(i, 1)) <> 0 Then
                    If IsNumeric(.TextMatrix(i, k)) And IsNumeric(.TextMatrix(i, 10)) Then '9
'                        MsgBox " sni"
                        If .TextMatrix(i, k) * 1 > 0 Then
                            
                            RsA.Filter = adFilterNone
                            RsA.Filter = "no_mach='" & .TextMatrix(i, 0) & "' and lcd_itemdid='" & .TextMatrix(i, 3) _
                            & "' and reg_mold='" & .TextMatrix(i, 5) & "'"
                            If RsA.RecordCount > 0 Then
                                If .TextMatrix(i, 10) > 0 Then
                                    ttlPerHari = ttlPerHari + ((.TextMatrix(i, k) / .TextMatrix(i, 10)) / (RsA("shiftusg") * RsA("hourpshift")) * RsA("mpower")) '9
                                Else
                                    ttlPerHari = ttlPerHari
                                End If
                            End If
                        End If
                    End If
                Else
                    If .TextMatrix(i, (COLSDATE - 1)) = "MP" Then
                        .TextMatrix(i, k) = FormatNumber(ttlPerHari, 2)
                        ttlPerHari = 0
                    End If
                End If
            Next
        Next
    End With
End Sub

Private Sub hitungRekapSUB()
    ReDim ar_totallc(1 To 1) As Variant
    With anaSubcont
        For i = 2 To .rows - 1
            If Len(.TextMatrix(i, 0)) > 1 Then
                ttlMPP = 0
                For c = 20 To .Cols - 1
                    If IsNumeric(.TextMatrix(i, c)) Then
                        ttlMPP = ttlMPP + (.TextMatrix(i, c) * 1)
                        
                    End If
                Next
                .TextMatrix(i, 18) = FormatNumber(ttlMPP, 0)
                .TextMatrix(i, 14) = ttlMPP / .TextMatrix(i, 12)
                .TextMatrix(i, 13) = ttlMPP / .TextMatrix(i, 10)
                If ttlMPP <> 0 Then
                    .TextMatrix(i, 19) = (ttlMPP / .TextMatrix(i, 12)) / HKWs * 100
                Else
                    .TextMatrix(i, 19) = 0
                End If
                .TextMatrix(i, 13) = FormatNumber(.TextMatrix(i, 13), 2)
                .TextMatrix(i, 14) = FormatNumber(.TextMatrix(i, 14), 2)
                .TextMatrix(i, 19) = FormatNumber(.TextMatrix(i, 19), 2) & "%"
            End If
            If LenB(.TextMatrix(i, 0)) = 0 And LenB(.TextMatrix(i - 1, 0)) <> 0 Then
                .TextMatrix(i, 1) = "Total"
                .Row = i: .Col = 19
                st_MPP = 0
                st_CapDay = 0
                st_ReqHour = 0
                st_ReqDay = 0
                st_PP = 0
                st_ttlmpp = 0
                st_FC = 0
                ReDim Preserve ar_totallc(1 To UBound(ar_totallc) + 1)
                For r = i - 1 To 1 Step -1
                    If IsNumeric(.TextMatrix(r, 7)) Then 'jika berubah mesin
                        If LenB(.TextMatrix(r, 19)) <> 0 Then 'jika ada presentase
                            st_MPP = st_MPP + (Left$(.TextMatrix(r, 19), Len(.TextMatrix(r, 19)) - 1) * 1)
                        End If
                        If LenB(.TextMatrix(r, 12)) <> 0 Then
                            st_CapDay = st_CapDay + .TextMatrix(r, 12) * 1
                        End If
                        If LenB(.TextMatrix(r, 13)) <> 0 Then
                            st_ReqHour = st_ReqHour + .TextMatrix(r, 13) * 1
                        End If
                        If LenB(.TextMatrix(r, 14)) <> 0 Then
                            st_ReqDay = st_ReqDay + .TextMatrix(r, 14) * 1
                        End If
                        If LenB(.TextMatrix(r, 16)) <> 0 Then
                            st_PP = st_PP + .TextMatrix(r, 16) * 1
                        End If
                        If LenB(.TextMatrix(r, 17)) <> 0 Then
                            st_FC = st_FC + .TextMatrix(r, 17) * 1
                        End If
                        If LenB(.TextMatrix(r, 18)) <> 0 Then
                            st_ttlmpp = st_ttlmpp + .TextMatrix(r, 18) * 1
                        End If
                    Else
                        Exit For
                    End If
                Next
                .TextMatrix(i, 19) = st_MPP & "%"
                .TextMatrix(i, 18) = FormatNumber(st_ttlmpp, 0)
                .TextMatrix(i, 12) = FormatNumber(st_CapDay, 0)
                .TextMatrix(i, 13) = FormatNumber(st_ReqHour, 0)
                .TextMatrix(i, 14) = FormatNumber(st_ReqDay, 0)
                .TextMatrix(i, 16) = FormatNumber(st_PP, 0)
                .TextMatrix(i, 17) = FormatNumber(st_FC, 0)
                If st_MPP > 99 Then
                    .CellBackColor = RGB(255, 155, 155) ' merah muda
                ElseIf st_MPP > 80 Then
                    .CellBackColor = RGB(255, 255, 0) 'vbYellow
                ElseIf st_MPP > 0 Then
                    .CellBackColor = RGB(111, 255, 0) 'vbGreen
                End If
            End If
        Next
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
