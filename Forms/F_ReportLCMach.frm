VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form F_ReportLCMach 
   Caption         =   "Report Machine and Man Power"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   15960
   Begin VB.ComboBox cmbFiletype2 
      Height          =   375
      ItemData        =   "F_ReportLCMach.frx":0000
      Left            =   13920
      List            =   "F_ReportLCMach.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   76
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox cmbFiletype 
      Height          =   375
      ItemData        =   "F_ReportLCMach.frx":0023
      Left            =   6360
      List            =   "F_ReportLCMach.frx":002D
      Style           =   2  'Dropdown List
      TabIndex        =   75
      Top             =   1080
      Width           =   1935
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   10560
      ScaleHeight     =   360
      ScaleWidth      =   5295
      TabIndex        =   66
      Top             =   7540
      Width           =   5295
      Begin VB.Label LBLKMP1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   70
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLKMP2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   69
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLKMP3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   68
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLKMP4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   67
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   10560
      ScaleHeight     =   360
      ScaleWidth      =   5295
      TabIndex        =   61
      Top             =   7180
      Width           =   5295
      Begin VB.Label LBLRTMP4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   65
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLRTMP3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   64
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLRTMP2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   63
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLRTMP1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   10560
      ScaleHeight     =   360
      ScaleWidth      =   5295
      TabIndex        =   56
      Top             =   6820
      Width           =   5295
      Begin VB.Label LBLMSNOFF1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLMSNOFF2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   59
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLMSNOFF3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   58
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLMSNOFF4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   57
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   10560
      ScaleHeight     =   360
      ScaleWidth      =   5295
      TabIndex        =   51
      Top             =   6480
      Width           =   5295
      Begin VB.Label LBLMSNON4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   55
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLMSNON3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   54
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLMSNON2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   53
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLMSNON1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10560
      ScaleHeight     =   375
      ScaleWidth      =   5295
      TabIndex        =   46
      Top             =   6120
      Width           =   5295
      Begin VB.Label LBLBULAN4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   50
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLBULAN3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   49
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLBULAN2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LBLBULAN1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   8400
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   41
      Top             =   6120
      Width           =   2175
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "(person)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   74
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "(person)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   73
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "(unit)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   72
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "(unit)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   71
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Need MP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "AVG MP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Machine Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Machine On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1515
      End
   End
   Begin MSFlexGridLib.MSFlexGrid agrid 
      Height          =   1695
      Left            =   120
      TabIndex        =   40
      Top             =   6120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2990
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExportLC 
      Caption         =   "Export"
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      ToolTipText     =   "Spreadsheet"
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8400
      ScaleHeight     =   435
      ScaleWidth      =   7395
      TabIndex        =   11
      Top             =   1560
      Width           =   7455
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Need of Man Power"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   8115
      TabIndex        =   8
      Top             =   1560
      Width           =   8175
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Capacity per Machine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   13920
      TabIndex        =   7
      ToolTipText     =   "Spreadsheet"
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox CmbRevision 
      Height          =   375
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox CmbDocument 
      Height          =   375
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportLCMach.frx":0046
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6800
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   0
      OleObjectBlob   =   "F_ReportLCMach.frx":00AC
      Top             =   600
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
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
      Format          =   152502275
      CurrentDate     =   42544
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportLCMach.frx":02E0
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "F_ReportLCMach.frx":0342
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin MSComctlLib.ListView lv2 
      Height          =   3855
      Left            =   8400
      TabIndex        =   10
      Top             =   2160
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "F_ReportLCMach.frx":03A8
      TabIndex        =   14
      Top             =   6120
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "F_ReportLCMach.frx":0412
      TabIndex        =   15
      Top             =   6480
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "F_ReportLCMach.frx":0484
      TabIndex        =   16
      Top             =   6840
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "F_ReportLCMach.frx":04F8
      TabIndex        =   17
      Top             =   7200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "F_ReportLCMach.frx":055C
      TabIndex        =   18
      Top             =   7560
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel z1 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "F_ReportLCMach.frx":05CA
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel z2 
      Height          =   255
      Left            =   12120
      OleObjectBlob   =   "F_ReportLCMach.frx":062C
      TabIndex        =   20
      Top             =   6120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel z3 
      Height          =   255
      Left            =   13320
      OleObjectBlob   =   "F_ReportLCMach.frx":068E
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel z4 
      Height          =   255
      Left            =   14520
      OleObjectBlob   =   "F_ReportLCMach.frx":06F0
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel a1 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "F_ReportLCMach.frx":0752
      TabIndex        =   23
      Top             =   6480
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel a2 
      Height          =   255
      Left            =   12120
      OleObjectBlob   =   "F_ReportLCMach.frx":07C4
      TabIndex        =   24
      Top             =   6480
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel a3 
      Height          =   255
      Left            =   13320
      OleObjectBlob   =   "F_ReportLCMach.frx":0836
      TabIndex        =   25
      Top             =   6480
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel a4 
      Height          =   255
      Left            =   14520
      OleObjectBlob   =   "F_ReportLCMach.frx":08A8
      TabIndex        =   26
      Top             =   6480
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel d1 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "F_ReportLCMach.frx":091A
      TabIndex        =   27
      Top             =   6840
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel d2 
      Height          =   255
      Left            =   12120
      OleObjectBlob   =   "F_ReportLCMach.frx":098C
      TabIndex        =   28
      Top             =   6840
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel d3 
      Height          =   255
      Left            =   13320
      OleObjectBlob   =   "F_ReportLCMach.frx":09FE
      TabIndex        =   29
      Top             =   6840
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel d4 
      Height          =   255
      Left            =   14520
      OleObjectBlob   =   "F_ReportLCMach.frx":0A70
      TabIndex        =   30
      Top             =   6840
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel s1 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "F_ReportLCMach.frx":0AE2
      TabIndex        =   31
      Top             =   7200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel s2 
      Height          =   255
      Left            =   12120
      OleObjectBlob   =   "F_ReportLCMach.frx":0B54
      TabIndex        =   32
      Top             =   7200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel s3 
      Height          =   255
      Left            =   13320
      OleObjectBlob   =   "F_ReportLCMach.frx":0BC6
      TabIndex        =   33
      Top             =   7200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel s4 
      Height          =   255
      Left            =   14520
      OleObjectBlob   =   "F_ReportLCMach.frx":0C38
      TabIndex        =   34
      Top             =   7200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel e1 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "F_ReportLCMach.frx":0CAA
      TabIndex        =   35
      Top             =   7560
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel e2 
      Height          =   255
      Left            =   12120
      OleObjectBlob   =   "F_ReportLCMach.frx":0D1C
      TabIndex        =   36
      Top             =   7560
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel e3 
      Height          =   255
      Left            =   13320
      OleObjectBlob   =   "F_ReportLCMach.frx":0D8E
      TabIndex        =   37
      Top             =   7560
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel e4 
      Height          =   255
      Left            =   14520
      OleObjectBlob   =   "F_ReportLCMach.frx":0E00
      TabIndex        =   38
      Top             =   7560
      Width           =   975
   End
   Begin MSComctlLib.ListView lv3 
      Height          =   1695
      Left            =   2760
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2990
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "F_ReportLCMach"
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
Private RsA As ADODB.Recordset
Private rsB As ADODB.Recordset
Dim qry As String
Dim nmbulan() As String
Dim period1 As String
Dim period2 As String
Dim period3 As String
Dim period4 As String
Private oExcel      As Object 'Excel.Application
Private oBook       As Object 'Excel.Workbook
Private oSheet      As Object 'Excel.Worksheet
Dim i As Integer, j As Integer
Dim ttlMesin As Integer
Dim ttlMPPERIOD1 As Variant
Dim ttlMPPERIOD2 As Variant
Dim ttlMPPERIOD3 As Variant
Dim ttlMPPERIOD4 As Variant

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

Private Sub CmbDocument_DropDown()
    qry = "select distinct on (fltpp_doc) fltpp_doc from loadcap_generate_h where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "'"
    Set RsA = Con.Execute(qry)
    CmbDocument.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            CmbDocument.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub

Private Function nmAngkakeBulan(pis As String) As String
    Dim x As Integer
    For x = 1 To UBound(nmbulan)
        If x = pis Then
            nmAngkakeBulan = nmbulan(x)
            Exit For
        End If
    Next
End Function

Private Sub formatHeaderFG()
    With agrid
        .TextMatrix(1, 2) = nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, 3) = nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DateAdd("m", 1, DTPicker1.Value), "yy")
        .TextMatrix(1, 4) = nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DateAdd("m", 2, DTPicker1.Value), "yy")
        .TextMatrix(1, 5) = nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DateAdd("m", 3, DTPicker1.Value), "yy")
        
        .TextMatrix(1, 6) = nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, 7) = nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DateAdd("m", 1, DTPicker1.Value), "yy")
        .TextMatrix(1, 8) = nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DateAdd("m", 2, DTPicker1.Value), "yy")
        .TextMatrix(1, 9) = nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DateAdd("m", 3, DTPicker1.Value), "yy")
        
        .TextMatrix(1, 10) = nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
        .TextMatrix(1, 11) = nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DateAdd("m", 1, DTPicker1.Value), "yy")
        .TextMatrix(1, 12) = nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DateAdd("m", 2, DTPicker1.Value), "yy")
        .TextMatrix(1, 13) = nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DateAdd("m", 3, DTPicker1.Value), "yy")
    End With
End Sub

Private Sub formatHeader3()
    lv3.ColumnHeaders(3).Text = "Loading AVG " & nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
    lv3.ColumnHeaders(4).Text = "Total Loading " & nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
    lv3.ColumnHeaders(5).Text = "Loading AVG " & nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
    lv3.ColumnHeaders(6).Text = "Total Loading " & nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DTPicker1, "yy")
End Sub

Private Sub formatHeaderBulan(LV As ListView)
    LV.ColumnHeaders(3).Text = nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
    LV.ColumnHeaders(4).Text = nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DateAdd("m", 1, DTPicker1.Value), "yy")
    LV.ColumnHeaders(5).Text = nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DateAdd("m", 2, DTPicker1.Value), "yy")
    LV.ColumnHeaders(6).Text = nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DateAdd("m", 3, DTPicker1.Value), "yy")
    LBLBULAN1 = nmAngkakeBulan(Val(Right(period1, 2))) & "-" & Format(DTPicker1, "yy")
    LBLBULAN2 = nmAngkakeBulan(Val(Right(period2, 2))) & "-" & Format(DateAdd("m", 1, DTPicker1.Value), "yy")
    LBLBULAN3 = nmAngkakeBulan(Val(Right(period3, 2))) & "-" & Format(DateAdd("m", 2, DTPicker1.Value), "yy")
    LBLBULAN4 = nmAngkakeBulan(Val(Right(period4, 2))) & "-" & Format(DateAdd("m", 3, DTPicker1.Value), "yy")
End Sub

Private Sub CmbRevision_Click()
    If CmbRevision <> "" Then
        Dim xx As ListItem
        period1 = Format(DTPicker1.Value, "yyyyMM")
        period2 = Format(DateAdd("m", 1, DTPicker1.Value), "yyyyMM") 'Left(period1, 4) & Right("00" & Val(Right(period1, 2) + 1), 2)
        period3 = Format(DateAdd("m", 2, DTPicker1.Value), "yyyyMM") 'Left(period2, 4) & Right("00" & Val(Right(period2, 2) + 1), 2)
        period4 = Format(DateAdd("m", 3, DTPicker1.Value), "yyyyMM") 'Left(period3, 4) & Right("00" & Val(Right(period3, 2) + 1), 2)
    

        qry = "select b.no_mach,max(tonage_mach) tonage,sum(case when a.fltpp_ym='" & period1 & "' then lcvsmach end) lcvsmach," _
            & "sum(case when a.fltpp_ym='" & period2 & "' then lcvsmach end) lcvsmach2, " _
            & "sum(case when a.fltpp_ym='" & period3 & "' then lcvsmach end) lcvsmach3," _
            & "sum(case when a.fltpp_ym='" & period4 & "' then lcvsmach end) lcvsmach4,max(remark) remark," _
            & "sum(case when a.fltpp_ym='" & period1 & "' then lcneed_mp end) nmp1," _
            & "sum(case when a.fltpp_ym='" & period2 & "' then lcneed_mp end) nmp2," _
            & "sum(case when a.fltpp_ym='" & period3 & "' then lcneed_mp end) nmp3," _
            & "sum(case when a.fltpp_ym='" & period4 & "' then lcneed_mp end) nmp4,coalesce(rstate_mach,state_mach,rstate_mach) stsmsn " _
            & " from loadcap_generate_d a " _
             & " right join loadcap_mst_mach b on a.no_mach=b.no_mach and a.fltpp_rev=" & CmbRevision & "  and a.fltpp_doc='" & CmbDocument & "' " _
            & " left join v_mc_mat c on b.no_mach=c.no_mach " _
            & " " _
            & " group by b.no_mach,rstate_mach,state_mach " _
            & " order by 1"
        Set RsA = Con.Execute(qry)
        lv1.ListItems.Clear
        LV2.ListItems.Clear
        ttlMesin = 0
        ttlMPPERIOD1 = 0
        ttlMPPERIOD2 = 0
        ttlMPPERIOD3 = 0
        ttlMPPERIOD4 = 0
        If RsA.RecordCount > 0 Then
            formatHeaderBulan lv1
            formatHeaderBulan LV2
            While Not RsA.EOF
                Set xx = lv1.ListItems.Add(, , RsA("no_mach"))
                xx.SubItems(1) = IIf(IsNull(RsA("tonage")), 0, RsA("tonage")) & "T"
                xx.SubItems(2) = IIf(IsNull(RsA("lcvsmach")), 0, RsA("lcvsmach")) & "%"
                xx.SubItems(3) = IIf(IsNull(RsA("lcvsmach2")), 0, RsA("lcvsmach2")) & "%"
                xx.SubItems(4) = IIf(IsNull(RsA("lcvsmach3")), 0, RsA("lcvsmach3")) & "%"
                xx.SubItems(5) = IIf(IsNull(RsA("lcvsmach4")), 0, RsA("lcvsmach4")) & "%"
                xx.SubItems(6) = IIf(IsNull(RsA("remark")), "", RsA("remark"))
                xx.SubItems(7) = IIf(IsNull(RsA("stsmsn")), "", RsA("stsmsn"))
                
                Set xx = LV2.ListItems.Add(, , RsA("no_mach"))
                xx.SubItems(1) = IIf(IsNull(RsA("tonage")), 0, RsA("tonage")) & "T"
                xx.SubItems(2) = IIf(IsNull(RsA("nmp1")), 0, RsA("nmp1"))
                xx.SubItems(3) = IIf(IsNull(RsA("nmp2")), 0, RsA("nmp2"))
                xx.SubItems(4) = IIf(IsNull(RsA("nmp3")), 0, RsA("nmp3"))
                xx.SubItems(5) = IIf(IsNull(RsA("nmp4")), 0, RsA("nmp4"))
                xx.SubItems(6) = IIf(IsNull(RsA("stsmsn")), "", RsA("stsmsn"))
                If RsA("stsmsn") = 1 Then
                    ttlMesin = ttlMesin + 1
                End If
                RsA.MoveNext
            Wend
            LBLMSNON1.Caption = ttlMesin
            LBLMSNON2.Caption = LBLMSNON1
            LBLMSNON3.Caption = LBLMSNON1
            LBLMSNON4.Caption = LBLMSNON1
            LBLMSNOFF1.Caption = lv1.ListItems.Count - ttlMesin
            LBLMSNOFF2.Caption = LBLMSNOFF1
            LBLMSNOFF3.Caption = LBLMSNOFF1
            LBLMSNOFF4.Caption = LBLMSNOFF1
            For i = lv1.ListItems.Count To 1 Step -1
                ttlMPPERIOD1 = ttlMPPERIOD1 + LV2.ListItems(i).ListSubItems(2).Text * 1
                ttlMPPERIOD2 = ttlMPPERIOD2 + LV2.ListItems(i).ListSubItems(3).Text * 1
                ttlMPPERIOD3 = ttlMPPERIOD3 + LV2.ListItems(i).ListSubItems(4).Text * 1
                ttlMPPERIOD4 = ttlMPPERIOD1 + LV2.ListItems(i).ListSubItems(5).Text * 1
            Next
            LBLRTMP1.Caption = FormatNumber(ttlMPPERIOD1 / LBLMSNON1, 2)
            LBLRTMP2.Caption = FormatNumber(ttlMPPERIOD2 / LBLMSNON2, 2)
            LBLRTMP3.Caption = FormatNumber(ttlMPPERIOD3 / LBLMSNON3, 2)
            LBLRTMP4.Caption = FormatNumber(ttlMPPERIOD4 / LBLMSNON4, 2)
            
            LBLKMP1 = FormatNumber(LBLMSNON1 * (ttlMPPERIOD1 / LBLMSNON1), 2)
            LBLKMP2 = FormatNumber(LBLMSNON2 * (ttlMPPERIOD2 / LBLMSNON2), 2)
            LBLKMP3 = FormatNumber(LBLMSNON3 * (ttlMPPERIOD3 / LBLMSNON3), 2)
            LBLKMP4 = FormatNumber(LBLMSNON4 * (ttlMPPERIOD4 / LBLMSNON4), 2)
        End If
        qry = "select tonage_mach, COUNT(distinct(b.no_mach)) ttlmesin,sum(case when a.fltpp_ym='" & period1 & "' then lcvsmach end) avglc1, " _
            & " sum(case when a.fltpp_ym='" & period2 & "' then lcvsmach end) avglc2," _
            & " sum(case when a.fltpp_ym='" & period3 & "' then lcvsmach end) avglc3," _
            & " sum(case when a.fltpp_ym='" & period4 & "' then lcvsmach end) avglc4" _
            & " from loadcap_generate_d a " _
            & " right join loadcap_mst_mach b on a.no_mach=b.no_mach and a.fltpp_rev=" & CmbRevision & " and a.fltpp_doc='" & CmbDocument & "' " _
            & " left join v_mc_mat c on b.no_mach=c.no_mach " _
            & " group by tonage_mach " _
            & " order by 1"
        Set RsA = Con.Execute(qry)
        lv3.ListItems.Clear
        If RsA.RecordCount > 0 Then
'            formatHeader3
            formatHeaderFG
            i = 2
            agrid.rows = RsA.RecordCount + i
            While Not RsA.EOF
                With agrid
                    .TextMatrix(i, 0) = RsA("tonage_mach")
                    .TextMatrix(i, 1) = RsA("ttlmesin")
                    .TextMatrix(i, 2) = FormatNumber(IIf(IsNull(RsA("avglc1")), 0, RsA("avglc1")), 2) & "%"
                    .TextMatrix(i, 3) = FormatNumber(IIf(IsNull(RsA("avglc2")), 0, RsA("avglc2")), 2) & "%"
                    .TextMatrix(i, 4) = FormatNumber(IIf(IsNull(RsA("avglc3")), 0, RsA("avglc3")), 2) & "%"
                    .TextMatrix(i, 5) = FormatNumber(IIf(IsNull(RsA("avglc4")), 0, RsA("avglc4")), 2) & "%"
                    
                    .TextMatrix(i, 6) = FormatNumber(IIf(IsNull(RsA("avglc1")), 0, RsA("avglc1")) / RsA("ttlmesin"), 2) & "%"
                    .TextMatrix(i, 7) = FormatNumber(IIf(IsNull(RsA("avglc2")), 0, RsA("avglc2")) / RsA("ttlmesin"), 2) & "%"
                    .TextMatrix(i, 8) = FormatNumber(IIf(IsNull(RsA("avglc3")), 0, RsA("avglc3")) / RsA("ttlmesin"), 2) & "%"
                    .TextMatrix(i, 9) = FormatNumber(IIf(IsNull(RsA("avglc4")), 0, RsA("avglc4")) / RsA("ttlmesin"), 2) & "%"
                End With
                i = i + 1
                RsA.MoveNext
            Wend
            Dim k As Byte, touta As Byte, touta2 As Byte, touta3 As Byte, touta4 As Byte
            With agrid
                For i = 2 To agrid.rows - 1
                    touta = 0
                    touta2 = 0
                    touta3 = 0
                    touta4 = 0
                    For k = 1 To lv1.ListItems.Count
                        If .TextMatrix(i, 0) = Left(lv1.ListItems(k).ListSubItems(1), Len(lv1.ListItems(k).ListSubItems(1)) - 1) Then
                            If Left(lv1.ListItems(k).ListSubItems(2), Len(lv1.ListItems(k).ListSubItems(2)) - 1) * 1 > 100 Then
                                touta = touta + 1
                            End If
                           
                        End If
                        If .TextMatrix(i, 0) = Left(lv1.ListItems(k).ListSubItems(1), Len(lv1.ListItems(k).ListSubItems(1)) - 1) Then
                            If Left(lv1.ListItems(k).ListSubItems(3), Len(lv1.ListItems(k).ListSubItems(3)) - 1) * 1 > 100 Then
                                touta2 = touta2 + 1
                            End If
                           
                        End If
                        If .TextMatrix(i, 0) = Left(lv1.ListItems(k).ListSubItems(1), Len(lv1.ListItems(k).ListSubItems(1)) - 1) Then
                            If Left(lv1.ListItems(k).ListSubItems(4), Len(lv1.ListItems(k).ListSubItems(4)) - 1) * 1 > 100 Then
                                touta3 = touta3 + 1
                            End If
                           
                        End If
                        If .TextMatrix(i, 0) = Left(lv1.ListItems(k).ListSubItems(1), Len(lv1.ListItems(k).ListSubItems(1)) - 1) Then
                            If Left(lv1.ListItems(k).ListSubItems(5), Len(lv1.ListItems(k).ListSubItems(5)) - 1) * 1 > 100 Then
                                touta4 = touta4 + 1
                            End If
                        End If
                    Next
                    .TextMatrix(i, 10) = touta
                    .TextMatrix(i, 11) = touta2
                    .TextMatrix(i, 12) = touta3
                    .TextMatrix(i, 13) = touta4
                    If i Mod 2 = 0 Then
                        For j = 0 To .Cols - 1
                            .Col = j
                            .Row = i
                            .CellBackColor = RGB(255, 255, 149)
                        Next
                    End If
                    .Col = 2
                    .Row = i
                    .CellAlignment = flexAlignRightCenter
                Next
            End With
        End If
    End If
End Sub

Private Sub CmbRevision_DropDown()
    qry = "select distinct on (fltpp_rev) fltpp_rev from loadcap_generate_h where fltpp_period='" & Format(DTPicker1.Value, "yyyyMM") & "' and fltpp_doc='" & CmbDocument & "'"
    Set RsA = Con.Execute(qry)
    CmbRevision.Clear
    If RsA.RecordCount > 0 Then
        While Not RsA.EOF
            CmbRevision.AddItem RsA(0)
            RsA.MoveNext
        Wend
    End If
End Sub

Private Sub cmdExport_Click()
    Dim spreasheet      As String
    If cmbFiletype2.ListIndex = 0 Then
        spreasheet = "Excel.Application"
    Else
        spreasheet = "Ket.Application"
    End If
    If LV2.ListItems.Count < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        
        oSheet.Cells(1, 1) = "Need of Man Power"
        oSheet.Cells(2, 1) = LV2.ColumnHeaders(1).Text
        oSheet.Cells(2, 2) = LV2.ColumnHeaders(2).Text
        oSheet.Cells(2, 3) = DTPicker1.Value
        oSheet.Cells(2, 3).NumberFormat = "mmm-yy"
        oSheet.Cells(2, 4) = DateAdd("m", 1, DTPicker1.Value)
        oSheet.Cells(2, 4).NumberFormat = "mmm-yy"
        oSheet.Cells(2, 5) = DateAdd("m", 2, DTPicker1.Value)
        oSheet.Cells(2, 5).NumberFormat = "mmm-yy"
        oSheet.Cells(2, 6) = DateAdd("m", 3, DTPicker1.Value)
        oSheet.Cells(2, 6).NumberFormat = "mmm-yy"
        Dim i As Integer, baris As Integer
        baris = 3
        For i = 1 To LV2.ListItems.Count
            oSheet.Cells(baris, 1) = LV2.ListItems(i).Text
            oSheet.Cells(baris, 2) = LV2.ListItems(i).SubItems(1)
            oSheet.Cells(baris, 3) = LV2.ListItems(i).SubItems(2)
            oSheet.Cells(baris, 4) = LV2.ListItems(i).SubItems(3)
            oSheet.Cells(baris, 5) = LV2.ListItems(i).SubItems(4)
            oSheet.Cells(baris, 6) = LV2.ListItems(i).SubItems(5)
            baris = baris + 1
        Next
        
        oSheet.Cells(baris + 1, 2) = "Keterangan"
        oSheet.Cells(baris + 1, 3) = DTPicker1.Value 'LBLBULAN1.Caption
        oSheet.Cells(baris + 1, 3).NumberFormat = "mmm-yy"
        oSheet.Cells(baris + 1, 4) = DateAdd("m", 1, DTPicker1.Value) 'LBLBULAN2.Caption
        oSheet.Cells(baris + 1, 4).NumberFormat = "mmm-yy"
        oSheet.Cells(baris + 1, 5) = DateAdd("m", 2, DTPicker1.Value) 'LBLBULAN3.Caption
        oSheet.Cells(baris + 1, 5).NumberFormat = "mmm-yy"
        oSheet.Cells(baris + 1, 6) = DateAdd("m", 3, DTPicker1.Value) 'LBLBULAN4.Caption
        oSheet.Cells(baris + 1, 6).NumberFormat = "mmm-yy"
        
        
        oSheet.Cells(baris + 2, 2) = Label4.Caption & " " & Label3.Caption
        oSheet.Cells(baris + 2, 3) = LBLMSNON1.Caption
        oSheet.Cells(baris + 2, 4) = LBLMSNON2.Caption
        oSheet.Cells(baris + 2, 5) = LBLMSNON3.Caption
        oSheet.Cells(baris + 2, 6) = LBLMSNON4.Caption
        
        oSheet.Cells(baris + 3, 2) = Label5.Caption & " " & Label8.Caption
        oSheet.Cells(baris + 3, 3) = LBLMSNOFF1.Caption
        oSheet.Cells(baris + 3, 4) = LBLMSNOFF2.Caption
        oSheet.Cells(baris + 3, 5) = LBLMSNOFF3.Caption
        oSheet.Cells(baris + 3, 6) = LBLMSNOFF4.Caption
        
        oSheet.Cells(baris + 4, 2) = Label6.Caption & " " & Label9.Caption
        oSheet.Cells(baris + 4, 3) = LBLRTMP1.Caption
        oSheet.Cells(baris + 4, 4) = LBLRTMP2.Caption
        oSheet.Cells(baris + 4, 5) = LBLRTMP3.Caption
        oSheet.Cells(baris + 4, 6) = LBLRTMP4.Caption
        
        oSheet.Cells(baris + 5, 2) = Label7.Caption & " " & Label10.Caption
        oSheet.Cells(baris + 5, 3) = LBLKMP1.Caption
        oSheet.Cells(baris + 5, 4) = LBLKMP2.Caption
        oSheet.Cells(baris + 5, 5) = LBLKMP3.Caption
        oSheet.Cells(baris + 5, 6) = LBLKMP4.Caption
        
        oExcel.Range(oExcel.Cells(2, 1), oExcel.Cells(baris - 1, 6)).Borders.LineStyle = 1 ' xlContinuous
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, -4143 'xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub cmdExportLC_Click()
    Dim spreasheet      As String
    If cmbFiletype.ListIndex = 0 Then
        spreasheet = "Excel.Application"
    Else
        spreasheet = "Ket.Application"
    End If
    If lv1.ListItems.Count < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject(spreasheet)
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        oSheet.Cells(1, 1) = "Loading Capacity /Machine"
        oSheet.Cells(2, 1) = lv1.ColumnHeaders(1).Text
        oSheet.Cells(2, 2) = lv1.ColumnHeaders(2).Text
        oSheet.Cells(2, 3) = DTPicker1.Value 'lv1.ColumnHeaders(3).Text
        oSheet.Cells(2, 4) = DateAdd("m", 1, DTPicker1.Value) 'lv1.ColumnHeaders(4).Text
        oSheet.Cells(2, 5) = DateAdd("m", 2, DTPicker1.Value) 'lv1.ColumnHeaders(5).Text
        oSheet.Cells(2, 6) = DateAdd("m", 3, DTPicker1.Value) 'lv1.ColumnHeaders(6).Text
        oSheet.Cells(2, 7) = lv1.ColumnHeaders(7).Text
        Dim i As Integer, baris As Integer, k As Integer
        For i = 1 To 7
            oSheet.Cells(2, i).NumberFormat = "mmm-yy"
        Next
        baris = 3
        For i = 1 To lv1.ListItems.Count
            oSheet.Cells(baris, 1) = lv1.ListItems(i).Text
            oSheet.Cells(baris, 2) = lv1.ListItems(i).SubItems(1)
            oSheet.Cells(baris, 3) = lv1.ListItems(i).SubItems(2)
            oSheet.Cells(baris, 4) = lv1.ListItems(i).SubItems(3)
            oSheet.Cells(baris, 5) = lv1.ListItems(i).SubItems(4)
            oSheet.Cells(baris, 6) = lv1.ListItems(i).SubItems(5)
            oSheet.Cells(baris, 7) = lv1.ListItems(i).SubItems(6)
            baris = baris + 1
        Next
        oExcel.Range(oExcel.Cells(2, 1), oExcel.Cells(baris - 1, 7)).Borders.LineStyle = 1 'xlContinuous
        baris = baris + 1
        oSheet.Range("A" & baris & ":A" & baris + 1).Merge
        oSheet.Range("A" & baris & ":A" & baris + 1).HorizontalAlignment = xlCenter
        oSheet.Range("A" & baris & ":A" & baris + 1).VerticalAlignment = xlCenter
        oSheet.Range("B" & baris & ":B" & baris + 1).Merge
        oSheet.Range("B" & baris & ":B" & baris + 1).HorizontalAlignment = xlCenter
        oSheet.Range("B" & baris & ":B" & baris + 1).VerticalAlignment = xlCenter
        oSheet.Range("B" & baris & ":B" & baris + 1).WrapText = True
        oSheet.Range("C" & baris & ":F" & baris).Merge
        oSheet.Range("C" & baris & ":F" & baris).HorizontalAlignment = xlCenter
        oSheet.Range("G" & baris & ":J" & baris).Merge
        oSheet.Range("G" & baris & ":J" & baris).HorizontalAlignment = xlCenter
        oSheet.Range("K" & baris & ":N" & baris).Merge
        oSheet.Range("K" & baris & ":N" & baris).HorizontalAlignment = xlCenter
        With agrid
            For i = 0 To .rows - 1
                oSheet.Cells(baris, 1) = .TextMatrix(i, 0)
                oSheet.Cells(baris, 2) = .TextMatrix(i, 1)
                If i = 1 Then
                    oSheet.Cells(baris, 3) = DTPicker1.Value '.TextMatrix(i, 2)
                    oSheet.Cells(baris, 4) = DateAdd("m", 1, DTPicker1.Value) '.TextMatrix(i, 3)
                    oSheet.Cells(baris, 5) = DateAdd("m", 2, DTPicker1.Value) '.TextMatrix(i, 4)
                    oSheet.Cells(baris, 6) = DateAdd("m", 3, DTPicker1.Value) '.TextMatrix(i, 5)
                    oSheet.Cells(baris, 7) = DTPicker1.Value '.TextMatrix(i, 6)
                    oSheet.Cells(baris, 8) = DateAdd("m", 1, DTPicker1.Value) '.TextMatrix(i, 7)
                    oSheet.Cells(baris, 9) = DateAdd("m", 2, DTPicker1.Value) '.TextMatrix(i, 8)
                    oSheet.Cells(baris, 10) = DateAdd("m", 3, DTPicker1.Value) '.TextMatrix(i, 9)
                    oSheet.Cells(baris, 11) = DTPicker1.Value '.TextMatrix(i, 6)
                    oSheet.Cells(baris, 12) = DateAdd("m", 1, DTPicker1.Value) '.TextMatrix(i, 7)
                    oSheet.Cells(baris, 13) = DateAdd("m", 2, DTPicker1.Value) '.TextMatrix(i, 8)
                    oSheet.Cells(baris, 14) = DateAdd("m", 3, DTPicker1.Value) '.TextMatrix(i, 9)
                    For k = 3 To 14
                        oSheet.Cells(baris, k).NumberFormat = "mmm-yy"
                    Next
                Else
                    oSheet.Cells(baris, 3) = .TextMatrix(i, 2)
                    oSheet.Cells(baris, 4) = .TextMatrix(i, 3)
                    oSheet.Cells(baris, 5) = .TextMatrix(i, 4)
                    oSheet.Cells(baris, 6) = .TextMatrix(i, 5)
                    oSheet.Cells(baris, 7) = .TextMatrix(i, 6)
                    oSheet.Cells(baris, 8) = .TextMatrix(i, 7)
                    oSheet.Cells(baris, 9) = .TextMatrix(i, 8)
                    oSheet.Cells(baris, 10) = .TextMatrix(i, 9)
                    oSheet.Cells(baris, 11) = .TextMatrix(i, 10)
                    oSheet.Cells(baris, 12) = .TextMatrix(i, 11)
                    oSheet.Cells(baris, 13) = .TextMatrix(i, 12)
                    oSheet.Cells(baris, 14) = .TextMatrix(i, 13)
                End If
                baris = baris + 1
            Next
        End With
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, -4143 'xlWorkbookNormal
        MsgBox "saved !", vbInformation, "Creating Template"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    Else
        MsgBox "Canceled !", vbInformation, "Createing Template"
    End If
End Sub

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub settingLV()
    With lv1
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "MC No"
        .ColumnHeaders.Add , , "Tonage", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "Remark"
        .ColumnHeaders.Add , , "Mach State", 0
    End With
    
    With LV2
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "MC No"
        .ColumnHeaders.Add , , "Tonage", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "bulan", , lvwColumnRight
        .ColumnHeaders.Add , , "Mach State", 0
    End With
    
    With lv3
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Tonage (T)"
        .ColumnHeaders.Add , , "Total Machine", , lvwColumnRight
        .ColumnHeaders.Add , , "Loading AVG ", , lvwColumnRight
        .ColumnHeaders.Add , , "Total Loading", , lvwColumnRight
        .ColumnHeaders.Add , , "Loading AVG ", , lvwColumnRight
        .ColumnHeaders.Add , , "Total Loading", , lvwColumnRight
        .ColumnHeaders.Add , , "Loading AVG ", , lvwColumnRight
        .ColumnHeaders.Add , , "Total Loading", , lvwColumnRight
        .ColumnHeaders.Add , , "Loading AVG ", , lvwColumnRight
        .ColumnHeaders.Add , , "Total Loading", , lvwColumnRight
        .ColumnHeaders.Add , , "Loading AVG ", , lvwColumnRight
        .ColumnHeaders.Add , , "Total Loading", , lvwColumnRight
    End With
    
    Dim i As Integer
    With agrid
        .Cols = 14: .ColWidth(0) = 700: .ColWidth(1) = 780: .ColWidth(2) = 900
        .ColWidth(3) = 900: .ColWidth(4) = 900: .ColWidth(5) = 900: .ColWidth(6) = 900
        .rows = 5
        .FixedRows = 2
        .FixedCols = 0
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        
        For i = 0 To .Cols - 1
            .Col = i
            .Row = 1
            .CellBackColor = RGB(255, 255, 74)
            .Col = i
            .Row = 0
            .CellBackColor = RGB(255, 255, 74)
            .CellAlignment = flexAlignCenterCenter
        Next
        
        
        .MergeCells = flexMergeRestrictRows
        i = 0
        .TextMatrix(0, i) = "Tonage":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 1
        .TextMatrix(0, i) = "Total Machine":        .TextMatrix(1, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 2
        .TextMatrix(0, i) = "Total Loading"
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = .TextMatrix(0, 2)
        .MergeCol(i) = True
        
        i = 6
        .TextMatrix(0, i) = "Average Loading"
        .MergeCol(i) = True
        
        i = 7
        .TextMatrix(0, i) = .TextMatrix(0, 6)
        .MergeCol(i) = True
        
        i = 8
        .TextMatrix(0, i) = .TextMatrix(0, 6)
        .MergeCol(i) = True
        
        i = 9
        .TextMatrix(0, i) = .TextMatrix(0, 6)
        .MergeCol(i) = True
        
        i = 10
        .TextMatrix(0, i) = "Overload"
        .MergeCol(i) = True
        .ColWidth(i) = 800
        
        
        i = 11
        .TextMatrix(0, i) = .TextMatrix(0, 10)
        .MergeCol(i) = True
        .ColWidth(i) = 800
        
        i = 12
        .TextMatrix(0, i) = .TextMatrix(0, 10)
        .MergeCol(i) = True
        .ColWidth(i) = 800
        
        i = 13
        .TextMatrix(0, i) = .TextMatrix(0, 10)
        .MergeCol(i) = True
        .ColWidth(i) = 800
            
        .MergeRow(0) = True
    End With
    
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

Private Sub Form_Load()
 On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call settingLV
    Call activeTheme(skinFD, Me)
    Me.Height = 8475
    Me.Width = 16155
    ReDim nmbulan(1 To 12) As String
    nmbulan(1) = "Jan"
    nmbulan(2) = "Feb"
    nmbulan(3) = "Mar"
    nmbulan(4) = "Apr"
    nmbulan(5) = "May"
    nmbulan(6) = "Jun"
    nmbulan(7) = "Jul"
    nmbulan(8) = "Aug"
    nmbulan(9) = "Sep"
    nmbulan(10) = "Oct"
    nmbulan(11) = "Nov"
    nmbulan(12) = "Dec"
    Picture3.BackColor = RGB(18, 173, 233)
    Picture5.BackColor = RGB(18, 173, 233)
    Picture4.BackColor = RGB(102, 217, 255)
    Picture7.BackColor = RGB(102, 217, 255)
    Picture2.BackColor = RGB(102, 217, 255)
    DTPicker1.Value = Now
    cmbFiletype.ListIndex = 0
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    With CmbDocument
        .Left = DTPicker1.Left: .Top = SkinLabel1.Top
    End With
    With CmbRevision
        .Left = DTPicker1.Left: .Top = SkinLabel3.Top
    End With
    cmbFiletype.Width = cmdExportLC.Width
    cmbFiletype.Left = cmdExportLC.Left
    cmbFiletype.Top = CmbRevision.Top
    cmbFiletype2.Top = cmbFiletype.Top
    cmbFiletype2.Left = cmdExport.Left
    cmbFiletype2.Width = cmdExport.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DelTab Me
End Sub

Private Sub lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     With lv1
        If .SortKey <> ColumnHeader.Index - 1 Then
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        Else
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
             Else
                 .SortOrder = lvwAscending
            End If
        End If
        .Sorted = -1
    End With
End Sub
