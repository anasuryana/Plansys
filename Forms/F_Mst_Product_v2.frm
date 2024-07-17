VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form F_Mst_Product_v2 
   Caption         =   "Master Product .."
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13935
   Begin VB.PictureBox Picmore 
      BackColor       =   &H0000C000&
      Height          =   4935
      Left            =   360
      ScaleHeight     =   4875
      ScaleWidth      =   7275
      TabIndex        =   41
      Top             =   3840
      Visible         =   0   'False
      Width           =   7335
      Begin VB.PictureBox picother 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   0
         ScaleHeight     =   3975
         ScaleWidth      =   7215
         TabIndex        =   92
         Top             =   840
         Visible         =   0   'False
         Width           =   7215
         Begin VB.CommandButton cmdImport 
            BackColor       =   &H0080FF80&
            Caption         =   "Import"
            Height          =   375
            Left            =   240
            TabIndex        =   95
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdCreateTempl 
            BackColor       =   &H0080FF80&
            Caption         =   "Export"
            Height          =   375
            Left            =   960
            TabIndex        =   94
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cmbFiletype 
            Height          =   360
            ItemData        =   "F_Mst_Product_v2.frx":0000
            Left            =   1920
            List            =   "F_Mst_Product_v2.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblcmdtype2 
            Caption         =   "Label29"
            Height          =   375
            Left            =   1920
            TabIndex        =   96
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Import Detail Data"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Import Header Data"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.PictureBox PicImportHeader 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   0
         ScaleHeight     =   3975
         ScaleWidth      =   7215
         TabIndex        =   85
         Top             =   840
         Width           =   7215
         Begin VB.CommandButton Command4 
            Caption         =   "Browse file"
            Height          =   375
            Left            =   90
            TabIndex        =   88
            Top             =   0
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Start"
            Height          =   375
            Left            =   5850
            TabIndex        =   87
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox txturlfile 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   86
            Text            =   "..."
            Top             =   0
            Width           =   4215
         End
         Begin MSComctlLib.ListView lvmore 
            Height          =   3375
            Left            =   0
            TabIndex        =   89
            Top             =   480
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   5953
            View            =   3
            LabelEdit       =   1
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
               Text            =   "No"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Part Number"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Polybag Label"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Box Label"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Color"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Status"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6840
         TabIndex        =   43
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "More option"
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
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   6855
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   8055
      TabIndex        =   44
      Top             =   0
      Width           =   8055
      Begin VB.ComboBox cmbBoxlabel 
         Height          =   360
         ItemData        =   "F_Mst_Product_v2.frx":0023
         Left            =   1440
         List            =   "F_Mst_Product_v2.frx":002D
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   360
         ItemData        =   "F_Mst_Product_v2.frx":0059
         Left            =   1440
         List            =   "F_Mst_Product_v2.frx":0066
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtcolor 
         Height          =   360
         Left            =   5760
         TabIndex        =   79
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtweight 
         Height          =   360
         Left            =   5760
         TabIndex        =   72
         ToolTipText     =   "gram"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtItemId 
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtItemName 
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   55
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtShift 
         Height          =   360
         Left            =   6720
         TabIndex        =   54
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtHourPshift 
         Height          =   360
         Left            =   6720
         TabIndex        =   53
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFaktorProd 
         Height          =   360
         Left            =   6720
         TabIndex        =   52
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtTotalMold 
         Height          =   360
         Left            =   1440
         TabIndex        =   51
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtMinStock 
         Height          =   360
         Left            =   1440
         TabIndex        =   50
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtMaxStock 
         Height          =   360
         Left            =   1440
         TabIndex        =   49
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtISNo 
         Height          =   360
         Left            =   5760
         TabIndex        =   48
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   360
         ItemData        =   "F_Mst_Product_v2.frx":00B5
         Left            =   1440
         List            =   "F_Mst_Product_v2.frx":00BF
         TabIndex        =   47
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdbox 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   46
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtBox 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3000
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   2295
         Left            =   4080
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   4048
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin VB.Label Label29 
         BackColor       =   &H0080FF80&
         Caption         =   "Box Label"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackColor       =   &H0080FF80&
         Caption         =   "Polybag Label"
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackColor       =   &H0080FF80&
         Caption         =   "Color"
         Height          =   375
         Left            =   4800
         TabIndex        =   80
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label24 
         BackColor       =   &H0080FF80&
         Caption         =   "Weight"
         Height          =   255
         Left            =   4800
         TabIndex        =   74
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label25 
         BackColor       =   &H0080FF80&
         Caption         =   "Kg"
         Height          =   255
         Left            =   6840
         TabIndex        =   73
         ToolTipText     =   "gram"
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "Part Number"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Part Name"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "Shift"
         Height          =   255
         Left            =   4800
         TabIndex        =   69
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Hour / Shift"
         Height          =   255
         Left            =   4800
         TabIndex        =   68
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Productivity Factor"
         Height          =   255
         Left            =   4800
         TabIndex        =   67
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "Total Mold"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackColor       =   &H0080FF80&
         Caption         =   "Min Stock"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackColor       =   &H0080FF80&
         Caption         =   "Max Stock"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackColor       =   &H0080FF80&
         Caption         =   "IS No"
         Height          =   255
         Left            =   4800
         TabIndex        =   63
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label22 
         BackColor       =   &H0080FF80&
         Caption         =   "Category"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label28 
         BackColor       =   &H0080FF80&
         Caption         =   "Box"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label32 
         BackColor       =   &H0080FF80&
         Caption         =   "Day"
         Height          =   255
         Left            =   3360
         TabIndex        =   60
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label33 
         BackColor       =   &H0080FF80&
         Caption         =   "Day"
         Height          =   255
         Left            =   3360
         TabIndex        =   59
         Top             =   2040
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0E0FF&
      Height          =   2535
      Left            =   5280
      ScaleHeight     =   2475
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   5160
         TabIndex        =   35
         Top             =   50
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   50
         Width           =   735
      End
      Begin VB.TextBox txtItemFind 
         Height          =   360
         Left            =   840
         TabIndex        =   8
         Top             =   50
         Width           =   2055
      End
      Begin MSComctlLib.ListView lv_sc 
         Height          =   1935
         Left            =   45
         TabIndex        =   7
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Find"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13320
      Top             =   2640
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3735
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "double click to edit"
      Top             =   5040
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6588
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
      Appearance      =   1
      NumItems        =   22
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Part No"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Part Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Man Power"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time Second Process"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Machine No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ALT MCH1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ALT MCH2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ALT MCH3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ALT MCH4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ALT MCH5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ALT MCH6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ALT MCH7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Cavity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Cycle Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Subcont"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Shift"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Hour per Shift"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Total Mold"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Used Mold"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Productivity Factor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "LabelBox"
         Object.Width           =   2540
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   8280
      OleObjectBlob   =   "F_Mst_Product_v2.frx":00DA
      Top             =   -120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   8055
      TabIndex        =   1
      Top             =   4440
      Width           =   8055
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "»"
         Height          =   495
         Left            =   7080
         TabIndex        =   78
         ToolTipText     =   "More option"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   5400
         TabIndex        =   77
         Tag             =   "s"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   6120
         TabIndex        =   76
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080FF80&
         Caption         =   "Add new"
         Height          =   495
         Left            =   4440
         TabIndex        =   75
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtFind 
         Height          =   360
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox txtQty_lot 
         Height          =   360
         Left            =   0
         TabIndex        =   0
         Top             =   4560
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "Find :"
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Qty Lot"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   4560
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   8040
      ScaleHeight     =   8895
      ScaleWidth      =   5895
      TabIndex        =   11
      Top             =   0
      Width           =   5895
      Begin VB.PictureBox PicCheckContainer 
         Height          =   5775
         Left            =   120
         ScaleHeight     =   5715
         ScaleWidth      =   5595
         TabIndex        =   98
         Top             =   3000
         Visible         =   0   'False
         Width           =   5655
         Begin VB.TextBox txtCheckResult 
            Height          =   5055
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   101
            Top             =   480
            Width           =   5295
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5160
            TabIndex        =   100
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "Result"
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
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   5175
         End
      End
      Begin VB.TextBox txtrunnershoot 
         Height          =   360
         Left            =   1440
         TabIndex        =   38
         Top             =   2520
         Width           =   855
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   5175
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Do you want to see history ? right click -> Details"
         Top             =   3600
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Part No"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Part Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Man Power"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time Second Process"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Machine No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ALT MCH1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ALT MCH2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ALT MCH3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ALT MCH4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ALT MCH5"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ALT MCH6"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "ALT MCH7"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Cavity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Cycle Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Subcont"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Shift"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Hour per Shift"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Total Mold"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Used Mold"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Productivity Factor"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdFindSubc 
         Caption         =   "..."
         Height          =   375
         Left            =   5400
         TabIndex        =   37
         Tag             =   "s"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Tag             =   "s"
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton cmdAddproc 
         Caption         =   "Save"
         Height          =   495
         Left            =   840
         TabIndex        =   33
         Tag             =   "s"
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton cmdDelProc 
         Caption         =   "Delete"
         Height          =   495
         Left            =   1560
         TabIndex        =   32
         Tag             =   "s"
         ToolTipText     =   "Delete Process"
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   2400
         TabIndex        =   31
         Tag             =   "s"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtSecondProcess 
         Height          =   360
         Left            =   4200
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtMchineNo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1440
         TabIndex        =   21
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton cmdFindMachine 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   20
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtmold 
         Height          =   360
         Left            =   1440
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtCT 
         Height          =   360
         Left            =   4200
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtCavity 
         Height          =   360
         Left            =   1440
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtPriority 
         Height          =   360
         Left            =   1440
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtManPower 
         Height          =   360
         Left            =   4200
         TabIndex        =   15
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtCavityStd 
         Height          =   360
         Left            =   1440
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "SUBCONT"
         Height          =   255
         Left            =   2640
         MaskColor       =   &H0080FF80&
         TabIndex        =   13
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtSubcontDi 
         Enabled         =   0   'False
         Height          =   360
         Left            =   4200
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCheckInputCalc 
         Caption         =   "Check "
         Height          =   495
         Left            =   3360
         TabIndex        =   97
         Tag             =   "s"
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Kg"
         Height          =   255
         Left            =   2320
         TabIndex        =   40
         ToolTipText     =   "gram"
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Runner/shoot"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Man Power"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CT 2nd Process"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cycle Time (CT)"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Priority"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cavity STD"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cavity Actual"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mould"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Machine No"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "F_Mst_Product_v2"
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
Public c_machine_no As String
Public idPROC       As String
Public typebox      As String
Private qry         As String
Private kebijakanSC As String
Private lisitm      As ListItem
Private id          As String, id2 As String
Private rsProc      As ADODB.Recordset
Private rsAneHelper As ADODB.Recordset
Private rsSubc      As ADODB.Recordset
Private oExcel      As Object 'Excel.Application
Private oBook       As Object 'Excel.Workbook
Private oSheet      As Object 'Excel.Worksheet
Private spreasheet  As String



Private Sub lvToForm()
    If LV.ListItems.Count > 0 Then
        id = LV.SelectedItem.Text
        txtItemid = LV.SelectedItem.SubItems(1)
        txtItemName = LV.SelectedItem.SubItems(2)
        txtTotalMold = LV.SelectedItem.SubItems(3)
        txtShift = LV.SelectedItem.SubItems(5)
        txtHourPshift = LV.SelectedItem.SubItems(6)
        txtFaktorProd = LV.SelectedItem.SubItems(7)
        txtQty_lot = LV.SelectedItem.SubItems(8)
        txtMinStock = LV.SelectedItem.SubItems(9)
        txtMaxStock = LV.SelectedItem.SubItems(10)
        If Len(LV.SelectedItem.SubItems(11)) > 0 Then
            cmbLabel = LV.SelectedItem.SubItems(11)
        End If
        
        txtISNo = LV.SelectedItem.SubItems(4)
        cmbCategory = LV.SelectedItem.SubItems(12)
        txtweight = LV.SelectedItem.SubItems(13)
        txtBox = LV.SelectedItem.SubItems(14)
        typebox = LV.SelectedItem.SubItems(15)
        If Len(LV.SelectedItem.SubItems(16)) > 0 Then
            cmbBoxlabel = LV.SelectedItem.SubItems(16)
        End If
        txtcolor = LV.SelectedItem.SubItems(17)
        Check1.Refresh
        LoadSubDatanya
    End If
End Sub

Private Sub lvtOform_v2()
    If LV2.ListItems.Count > 0 Then
        id2 = LV2.SelectedItem.Text
        idPROC = id2
        txtManPower = LV2.SelectedItem.SubItems(2)
        txtMchineNo = LV2.SelectedItem.SubItems(3)
        txtmold = LV2.SelectedItem.SubItems(4)
        txtCavity = LV2.SelectedItem.SubItems(5)
        txtCavityStd = LV2.SelectedItem.SubItems(6)
        txtCT = LV2.SelectedItem.SubItems(7)
        txtSecondProcess = LV2.SelectedItem.SubItems(8)
        txtpriority = LV2.SelectedItem.SubItems(9)
        If LV2.SelectedItem.SubItems(10) = "yes" Then
            Check1.Value = 1
            txtSubcontDi = LV2.SelectedItem.SubItems(3)
        Else
            Check1.Value = 0
            txtSubcontDi = ""
        End If
        txtrunnershoot = LV2.SelectedItem.SubItems(11)
        Check1.Refresh
    End If
End Sub

Private Sub settingLV()
    With LV
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "id", 0
        .ColumnHeaders.Add , , "Part No", 2000
        .ColumnHeaders.Add , , "Part Name", 3000
        .ColumnHeaders.Add , , "Total Mold", 1150
        .ColumnHeaders.Add , , "IS No", 1200 '900
        .ColumnHeaders.Add , , "Shift", 0 '900
        .ColumnHeaders.Add , , "Hour per Shift", 0
        .ColumnHeaders.Add , , "Productivity Factor"
        .ColumnHeaders.Add , , "Qty Lot", 0
        .ColumnHeaders.Add , , "Min Stock", 1200
        .ColumnHeaders.Add , , "Max Stock", 1200
        .ColumnHeaders.Add , , "Label Type"
        .ColumnHeaders.Add , , "Category"
        .ColumnHeaders.Add , , "Weight"
        .ColumnHeaders.Add , , "Type Box"
        .ColumnHeaders.Add , , "Type Box id", 0
        .ColumnHeaders.Add , , "Box Label"
        .ColumnHeaders.Add , , "Color"
    End With
    With LV2
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "id", 0
        .ColumnHeaders.Add , , "Part No", 0
        .ColumnHeaders.Add , , "Man Power", 900
        .ColumnHeaders.Add , , "Machine No"
        .ColumnHeaders.Add , , "Mold No"
        .ColumnHeaders.Add , , "Cavity", 700
        .ColumnHeaders.Add , , "Cavity STD", 1000
        .ColumnHeaders.Add , , "Cycle Time (CT)"
        .ColumnHeaders.Add , , "CT 2nd", 1000
        .ColumnHeaders.Add , , "Priority", 1000
        .ColumnHeaders.Add , , "Subcont"
        .ColumnHeaders.Add , , "Runner/shoot"
    End With
    With lv_sc
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Kode Subcont", 2500
        .ColumnHeaders.Add , , "Nama Subcont", 2000
        .ColumnHeaders.Add , , "Alamat", 5000
        .ColumnHeaders.Add , , "CP", 2500
        .ColumnHeaders.Add , , "Total Mesin", 2000
        .ColumnHeaders.Add , , "Prioritas", 800
    End With
End Sub

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

Private Sub getList()
    LV.ListItems.Clear
    LV2.ListItems.Clear
    Do Until RsGet.EOF
        Set lisitm = LV.ListItems.Add(, , RTrim(RsGet!lc_idproduct))
            lisitm.SubItems(1) = RTrim(RsGet!partNo)
            lisitm.SubItems(2) = RTrim(RsGet!item_name)
            lisitm.SubItems(3) = IIf(IsNull(RsGet!ttlmold), 0, RsGet!ttlmold)
            lisitm.SubItems(4) = IIf(IsNull(RsGet!isno), "-", RsGet!isno)
            lisitm.SubItems(5) = IIf(IsNull(RsGet!shift_usg), 0, RsGet!shift_usg)
            lisitm.SubItems(6) = IIf(IsNull(RsGet!hour_p_shift), 0, RsGet!hour_p_shift)
            lisitm.SubItems(7) = IIf(IsNull(RsGet!faktor_productivity), 0, RsGet!faktor_productivity)
            lisitm.SubItems(8) = IIf(IsNull(RsGet!qtylot), 0, RsGet!qtylot)
            lisitm.SubItems(9) = IIf(IsNull(RsGet!minstock), 0, RsGet!minstock)
            lisitm.SubItems(10) = IIf(IsNull(RsGet!maxstock), 0, RsGet!maxstock)
            lisitm.SubItems(12) = IIf(IsNull(RsGet!catgory), 0, RsGet!catgory)
            lisitm.SubItems(13) = IIf(IsNull(RsGet!wght), 0, RsGet!wght)
            If IsNull(RsGet!typelabel) Then
                lisitm.SubItems(11) = ""
            Else
            'Label Manual Logo BPI
            'Label Manual Logo ASKARA
            'Tidak Pakai Label Manual
                If RsGet!typelabel = 0 Then
                    lisitm.SubItems(11) = "Label Manual Logo BPI"
                ElseIf RsGet!typelabel = 1 Then
                    lisitm.SubItems(11) = "Label Manual Logo ASKARA"
                ElseIf RsGet!typelabel = 2 Then
                    lisitm.SubItems(11) = "Tidak Pakai Label Manual"
                End If
            End If
            lisitm.SubItems(14) = IIf(IsNull(RsGet!boxname), 0, RsGet!boxname)
            lisitm.SubItems(15) = IIf(IsNull(RsGet!typeboxid), 0, RsGet!typeboxid)
            If IsNull(RsGet!typelabelbox) Then
                lisitm.SubItems(16) = ""
            Else
                If RsGet!typelabelbox = 0 Then
                    lisitm.SubItems(16) = "Label Manual"
                ElseIf RsGet!typelabelbox = 1 Then
                    lisitm.SubItems(16) = "Tidak Pakai Label Manual"
                End If
            End If
            lisitm.SubItems(17) = IIf(IsNull(RsGet!colordesc), "-", RsGet!colordesc)
        RsGet.MoveNext
    Loop
End Sub

Private Sub LoadDatanyaSubcont()
    Set rsSubc = Con.Execute("select * from loadcap_mst_subcont order by prioritas asc")
End Sub

Private Sub getListSubcont()
    lv_sc.ListItems.Clear
    Do Until rsSubc.EOF
        Set lisitm = lv_sc.ListItems.Add(, , RTrim(rsSubc!kodesubcont))
            lisitm.SubItems(1) = RTrim(rsSubc!namasubcont)
            lisitm.SubItems(2) = RTrim(rsSubc!alamat)
            lisitm.SubItems(3) = rsSubc!kontakperson
            lisitm.SubItems(4) = IIf(IsNull(rsSubc!totalmesin_tersedia), 1, rsSubc!totalmesin_tersedia)
            lisitm.SubItems(5) = IIf(IsNull(rsSubc!prioritas), 0, rsSubc!prioritas)
        rsSubc.MoveNext
    Loop
End Sub

Private Sub getSubList()
    LV2.ListItems.Clear
    Do Until rsProc.EOF
        Set lisitm = LV2.ListItems.Add(, , RTrim$(rsProc("idproclc")))
            lisitm.SubItems(1) = RTrim$(rsProc("partNo"))
            lisitm.SubItems(2) = IIf(IsNull(rsProc!manPower), 0, rsProc!manPower)
            lisitm.SubItems(3) = IIf(IsNull(rsProc!prod_nomach), "", RTrim(rsProc!prod_nomach))
            lisitm.SubItems(4) = rsProc!mold_no
            lisitm.SubItems(5) = rsProc!cavity
            lisitm.SubItems(6) = IIf(IsNull(rsProc!cavity_std), 0, rsProc!cavity_std)
            lisitm.SubItems(7) = rsProc("ct")
            lisitm.SubItems(8) = rsProc("ct_2")
            lisitm.SubItems(9) = rsProc("priorit")
            lisitm.SubItems(10) = rsProc("subcont")
            lisitm.SubItems(11) = IIf(IsNull(rsProc("runnershoot")), 0, rsProc("runnershoot"))
        rsProc.MoveNext
    Loop
End Sub

Private Sub LoadDatanya()
    Set RsGet = Con.Execute("select lc_idproduct,partno,item_name,ttlmold,isno,shift_usg,hour_p_shift,hour_p_shift,faktor_productivity,qtylot,minstock,maxstock,catgory,typelabel,coalesce(wght,0) wght,coalesce(boxname,'--') boxname,coalesce(typeboxid,'--') typeboxid " _
    & ",typelabelbox,colordesc from loadcap_mst_product_r a inner join mst_item b on a.partno=b.item_id " _
    & " left join mst_box c on a.typeboxid=c.typeid order by partno asc")
End Sub

Private Sub LoadSubDatanya()
    Set rsProc = Con.Execute("select * from loadcap_proc where partno='" & txtItemid & "' order by priorit asc")
    Call getSubList
End Sub

Private Sub LoadCheckResult()
    Set rsProc = Con.Execute("select * from loadcap_proc where partno='" & txtItemid & "' order by priorit asc")
    Call getSubList
End Sub

Private Sub kosong()
    txtItemid = ""
    txtItemName = ""
    txtTotalMold = 0
    txtManPower = ""
    txtMchineNo = ""
    txtmold = "00000"
    txtCavity = 0
    txtCT = 0
    txtSecondProcess = 0
    txtpriority = ""
    txtQty_lot = 0
    txtweight = 0
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        kebijakanSC = "yes"
    Else
        kebijakanSC = "no"
    End If
End Sub

Private Sub cmdAdd_Click()
    kosong
    cmdfind.SetFocus
    cmdSave.Tag = "s"
End Sub

Private Function PeriksaMold() As Boolean
    qry = "select * from loadcap_proc where mold_no = '" & txtmold & "' and partno <> '" & txtItemid & "' limit 1"
    Set rsAneHelper = Con.Execute(qry)
    If rsAneHelper.RecordCount > 0 Then
        PeriksaMold = True
    Else
        PeriksaMold = False
    End If
End Function

Private Function PeriksaMesinMold() As Boolean
    qry = "select * from loadcap_proc where prod_nomach='" & txtMchineNo & "' and  mold_no = '" & txtmold & "' and partno = '" & txtItemid & "' limit 1"
    Set rsAneHelper = Con.Execute(qry)
    If rsAneHelper.RecordCount > 0 Then
        PeriksaMesinMold = True
    Else
        PeriksaMesinMold = False
    End If

End Function

Private Function PeriksaMesinMold_v2() As Boolean
    qry = "select * from loadcap_proc where prod_nomach='" & txtMchineNo & "' and  mold_no = '" & txtmold & "' and partno = '" & txtItemid & "' and cavity=" & txtCavity & " limit 1"
    Set rsAneHelper = Con.Execute(qry)
    If rsAneHelper.RecordCount > 0 Then
        PeriksaMesinMold_v2 = True
    Else
        PeriksaMesinMold_v2 = False
    End If
End Function

Private Sub cmdAddproc_Click()
    On Error GoTo Ex
    If Len(txtItemid) < 1 Then MsgBox "Please choose Part number first": Exit Sub
    If IsNumeric(txtManPower) = False Or (Len(txtmold) < 1 And Len(txtMchineNo) < 1) _
        Or IsNumeric(txtHourPshift) = False _
        Or Len(txtItemName) < 2 Or IsNumeric(txtFaktorProd) = False _
    Then
        MsgBox "Please check the data", vbExclamation
        Exit Sub
    End If
    If MsgBox("Save  ?", vbQuestion + vbYesNo) = vbYes Then
        BukaKoneksi
'        If PeriksaMold Then MsgBox "Mold tersebut sudah digunakan" & vbNewLine & "untuk mesin lain": Exit Sub
        If PeriksaMesinMold Then MsgBox "Data tersebut sudah terdaftar": Exit Sub
        preparedInsertProc
        Con.Execute qry
        insertLogProc "insert", ""
        MsgBox "Saved "
        LoadSubDatanya
    End If
    Exit Sub
Ex:
    MsgBox Err.Description, vbCritical, "Maaf"
End Sub

Private Sub insertLogProc(typeProc As String, idPROC As String)
    If Len(Trim(idPROC)) < 1 Then
        idPROC = 0
    End If
    
    qry = "insert into loadcap_lg_proc values(DEFAULT,'" & typeProc & "',now(),'" & pUserId & "'," & idPROC & ",'" & txtItemid & "' " _
        & ",'" & IIf(kebijakanSC = "no", txtMchineNo, txtSubcontDi) & "','" & txtmold & "'," & txtCavity & "" _
        & "," & txtCT & "," & txtSecondProcess & "," & txtpriority & "," & txtManPower & "," & txtCavityStd & ",'" & kebijakanSC & "')"
    Con.Execute qry
End Sub

Private Sub cmdbox_Click()
    Popup_Box.Show 1
End Sub

Private Sub cmdCheckInputCalc_Click()
    PicCheckContainer.Visible = True
        
    Set rsProc = Con.Execute("select * from loadcap_proc a where coalesce(a.ct,0)=0 " & _
            "or coalesce(a.cavity,0)=0")
    Dim strTemp_ As String
    strTemp_ = strTemp_ & "FG" & vbTab & "Machine" & vbTab & "Mold" & vbNewLine
    Dim isExist As Boolean
    
    Do Until rsProc.EOF
        isExist = True
        strTemp_ = strTemp_ & rsProc!partNo & vbTab & rsProc!prod_nomach & vbTab & rsProc!mold_no & vbNewLine
        rsProc.MoveNext
    Loop
    
    If isExist = False Then
        strTemp_ = "Item-Mchine Master is OK"
    End If
    
    txtCheckResult.Text = strTemp_
End Sub

Private Sub cmdCreateTempl_Click()
On Error GoTo exCe
    Screen.MousePointer = 11
    qry = "select * from v_lc_export order by partno asc"
    Set rsAneHelper = Con.Execute(qry)
    If rsAneHelper.RecordCount < 1 Then MsgBox "nothing to be exported": Exit Sub
    CommonDialog1.Filter = ""
    CommonDialog1.CancelError = True
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        If cmbFiletype.ListIndex = 0 Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "Ket.Application"
        End If
        Set oExcel = CreateObject(spreasheet) 'New Excel.Application
        Set oBook = oExcel.Workbooks.Add
        Set oSheet = oBook.Sheets.Item(1)
        oSheet.Cells(1, 1) = "Format Upload Master Product [Loading vs Capacity]"
        oSheet.Cells(2, 1) = "No"
        oSheet.Cells(2, 2) = "Part ID"
        oSheet.Cells(2, 3) = "Part Name"
        oSheet.Cells(2, 4) = "Category"
        oSheet.Cells(2, 5) = "Min Stock"
        oSheet.Cells(2, 6) = "Max Stock"
        oSheet.Cells(2, 7) = "SUBCONT"
        oSheet.Cells(2, 8) = "IS No"
        oSheet.Cells(2, 9) = "Machine"
        oSheet.Cells(2, 10) = "Mold"
        oSheet.Cells(2, 11) = "Cavity Actual"
        oSheet.Cells(2, 12) = "Cavity Std"
        oSheet.Cells(2, 13) = "CT"
        oSheet.Cells(2, 14) = "CT2"
        oSheet.Cells(2, 15) = "Man Power"
        oSheet.Cells(2, 16) = "Priority"
        oSheet.Cells(2, 17) = "Weight"
        oSheet.Cells(2, 18) = "Runner/shoot"
        oSheet.Cells(2, 19) = "Box Type"
        oSheet.Columns(10).NumberFormat = "@"
        
        Dim i As Integer, baris As Double
        baris = 3
        progres True
        ProgressBar1.Value = 0
        While Not rsAneHelper.EOF
            oSheet.Cells(baris, 1) = baris - 2
            oSheet.Cells(baris, 2) = rsAneHelper("partno")
            oSheet.Cells(baris, 3) = rsAneHelper("partname")
            oSheet.Cells(baris, 4) = IIf(IsNull(rsAneHelper("catgory")), "-", rsAneHelper("catgory"))
            oSheet.Cells(baris, 5) = rsAneHelper("minstock")
            oSheet.Cells(baris, 6) = rsAneHelper("maxstock")
            oSheet.Cells(baris, 7) = rsAneHelper("subcont")
            oSheet.Cells(baris, 8) = rsAneHelper("isno")
            oSheet.Cells(baris, 9) = rsAneHelper("prod_nomach")
            If Len(rsAneHelper("mold_no")) <> 0 Then
                If Right$(rsAneHelper("mold_no"), 2) = vbCrLf Or Right$(rsAneHelper("mold_no"), 2) = vbNewLine Then
                    oSheet.Cells(baris, 10) = Left$(rsAneHelper("mold_no"), Len(rsAneHelper("mold_no")) - 2)
                Else
                    oSheet.Cells(baris, 10) = RTrim(rsAneHelper("mold_no"))
                End If
            End If
            oSheet.Cells(baris, 11) = rsAneHelper("cavity")
            oSheet.Cells(baris, 12) = rsAneHelper("cavity_std")
            oSheet.Cells(baris, 13) = rsAneHelper("ct")
            oSheet.Cells(baris, 14) = rsAneHelper("ct_2")
            oSheet.Cells(baris, 15) = rsAneHelper("manpower")
            oSheet.Cells(baris, 16) = rsAneHelper("priorit")
            oSheet.Cells(baris, 17) = rsAneHelper("wght")
            oSheet.Cells(baris, 18) = rsAneHelper("runnershoot")
            oSheet.Cells(baris, 19) = rsAneHelper("typeboxid")
            baris = baris + 1
            ProgressBar1.Value = FormatNumber(((baris - 3) * 100) / rsAneHelper.RecordCount, 0)
            rsAneHelper.MoveNext
        Wend
        'xlWorkbookNormal
        oExcel.ActiveWorkbook.SaveAs CommonDialog1.FileName, -4143
        MsgBox "saved !", vbInformation, "Good"
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
        progres False
        
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

Private Sub progres(bmuncul As Boolean)
'    lblwaktu.Visible = bmuncul
    ProgressBar1.Visible = bmuncul
End Sub


Private Sub cmdDelete_Click()
On Error GoTo ErrEx
    If MsgBox("Are you want to delete data ?", vbQuestion + vbYesNo) = vbYes Then
        qry = "delete from loadcap_mst_product_r where lc_idproduct=" & id
        Con.Execute qry
        MsgBox "deleted successfully", vbInformation
    End If
    If Len(txtfind) > 2 Then
        LoadDatanya
        LoadDatanya_V2
    Else
        LoadDatanya
    End If
    Call getList
    Exit Sub
ErrEx:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub cmdDelProc_Click()
On Error GoTo ErrEx
    If MsgBox("Are you want to delete data ?", vbQuestion + vbYesNo) = vbYes Then
        insertLogProc "delete", id2
        qry = "delete from loadcap_proc where idproclc=" & id2
        Con.Execute qry
        MsgBox "deleted successfully", vbInformation
    End If
    LoadSubDatanya
    Exit Sub
ErrEx:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub cmdfind_Click()
    GetForm = Me.Name
    PopUp_Item_Sup.Show 1
End Sub

Private Sub cmdFindMachine_Click()
    GetForm = Me.Name
    PopUp_machine.Show 1
    txtMchineNo = c_machine_no
End Sub

Private Sub cmdFindSubc_Click()
    If Check1.Value = vbChecked Then
        kebijakanSC = "yes"
        Picture3.Visible = True
        txtMchineNo.Text = ""
    Else
        kebijakanSC = "no"
        Picture3.Visible = False
    End If
End Sub

Private Sub cmdImport_Click()
    Dim urlFILE As String, ada As Boolean, barisX As Double
    Dim rsKU As ADODB.Recordset
    Const NamaTabel As String = "loadcap_mst_product_r"
    With CommonDialog1
        .Filter = ""
        .ShowOpen
        urlFILE = .FileName
    End With
    If urlFILE <> "" Then
        If cmbFiletype.ListIndex = 0 Then
            spreasheet = "Excel.Application"
        Else
            spreasheet = "et.Application"
        End If
        Set oExcel = CreateObject(spreasheet)
        oExcel.Workbooks.Open urlFILE
        Set oBook = oExcel.Workbooks(1)
        Set oSheet = oBook.Worksheets(1)
        ada = True
        barisX = 3
        BukaKoneksi
        Screen.MousePointer = 11
        progres True
        While ada
            If oSheet.Cells(barisX, 2) <> "" Then
                qry = "SELECT partno FROM " & NamaTabel & " WHERE partno='" & Trim(oSheet.Cells(barisX, 2)) & "'"
                Set rsAneHelper = Con.Execute(qry)
                txtItemid = Trim(oSheet.Cells(barisX, 2))
                kebijakanSC = IIf(oSheet.Cells(barisX, 7) = "", "no", oSheet.Cells(barisX, 7))
                txtISNo = oSheet.Cells(barisX, 8)
                txtMchineNo = Trim(oSheet.Cells(barisX, 9))
                txtmold = Trim(oSheet.Cells(barisX, 10))
                txtCavity = IIf(oSheet.Cells(barisX, 11) = "", 0, oSheet.Cells(barisX, 11))
                txtCavityStd = IIf(oSheet.Cells(barisX, 12) = "", 0, oSheet.Cells(barisX, 12))
                txtCT = IIf(oSheet.Cells(barisX, 13) = "", 0, oSheet.Cells(barisX, 13))
                txtSecondProcess = IIf(oSheet.Cells(barisX, 14) = "", 0, oSheet.Cells(barisX, 14))
                txtManPower = IIf(oSheet.Cells(barisX, 15) = "", 0, oSheet.Cells(barisX, 15))
                txtpriority = IIf(oSheet.Cells(barisX, 16) = "", 0, oSheet.Cells(barisX, 16))
                txtweight = IIf(oSheet.Cells(barisX, 17) = "", 0, oSheet.Cells(barisX, 17))
                txtrunnershoot = IIf(oSheet.Cells(barisX, 18) = "", 0, oSheet.Cells(barisX, 18))
                If rsAneHelper.RecordCount > 0 Then
                    Set rsKU = New ADODB.Recordset
                    qry = "select partno,isno,ttlmold,isno,minstock,maxstock,catgory,wght,typeboxid from " & NamaTabel & " where partno='" & oSheet.Cells(barisX, 2) & "'"
                    rsKU.Open qry, Con, adOpenKeyset, adLockOptimistic, adCmdText
                    rsKU("isno") = oSheet.Cells(barisX, 8)
                    rsKU("catgory") = oSheet.Cells(barisX, 4)
                    rsKU("minstock") = oSheet.Cells(barisX, 5)
                    rsKU("maxstock") = oSheet.Cells(barisX, 6)
                    rsKU("wght") = oSheet.Cells(barisX, 17)
                    rsKU("typeboxid") = oSheet.Cells(barisX, 19)
                    
                    rsKU.Update
                    If Len(oSheet.Cells(barisX, 11)) > 0 Then
                        txtCavityStd = IIf(oSheet.Cells(barisX, 12) = "", 0, oSheet.Cells(barisX, 12))
                        txtMchineNo = oSheet.Cells(barisX, 9)
                        qry = "update loadcap_proc set cavity_std=" & txtCavityStd & ",prod_nomach='" & txtMchineNo & "',timeupdate=now(),runnershoot=" & txtrunnershoot & " where " _
                            & " partno='" & oSheet.Cells(barisX, 2) & "' and mold_no='" & oSheet.Cells(barisX, 10) & "' and " _
                            & " prod_nomach='" & oSheet.Cells(barisX, 9) & "'"
                        Con.Execute qry
                        If PeriksaMesinMold = False Then
                            preparedInsertProc
                            Con.Execute qry
                        End If
                    End If
                Else
                    qry = "INSERT INTO " & NamaTabel & " (lc_idproduct,partno,partname,catgory,minstock,maxstock,isno,wght,typeboxid) values(DEFAULT,'" & oSheet.Cells(barisX, 2) & "'" _
                            & ",'" & oSheet.Cells(barisX, 3) & "','" & oSheet.Cells(barisX, 4) & "'" _
                            & "," & oSheet.Cells(barisX, 5) & "," & oSheet.Cells(barisX, 6) & ",'" & oSheet.Cells(barisX, 8) & "'," & oSheet.Cells(barisX, 17) & ",'" & oSheet.Cells(barisX, 19) & "')"
                    Set rsKU = Con.Execute(qry)
                    
                    If PeriksaMesinMold = False Then
                        preparedInsertProc
                        Con.Execute qry
                    End If
                End If
            Else
                ada = False
            End If
            barisX = 1 + barisX
        Wend
        Screen.MousePointer = 0
        MsgBox "Uploaded !", vbInformation, "Upload Status"
        progres False
        
        oExcel.Quit
        Set oSheet = Nothing
        Set oBook = Nothing
        Set oExcel = Nothing
    End If
End Sub

Private Sub preparedInsertProc()
    If Len(txtMchineNo) > 0 Then
        If kebijakanSC = "no" Then
            qry = "insert into loadcap_proc(idproclc,partno,mold_no,cavity,prod_nomach,ct,ct_2, " _
            & " priorit,manpower,cavity_std,subcont,submch,timeupdate,runnershoot) values(DEFAULT,'" & txtItemid & "' " _
            & ",'" & txtmold & "'," & txtCavity & ",'" & txtMchineNo & "'" _
            & "," & txtCT & "," & txtSecondProcess & "," & txtpriority & "," & txtManPower & "," & txtCavityStd & ",'" & kebijakanSC & "',FALSE,now()," & txtrunnershoot & ")"
        Else
            qry = "insert into loadcap_proc(idproclc,partno,mold_no,cavity,prod_nomach,ct,ct_2, " _
            & " priorit,manpower,cavity_std,subcont,submch,timeupdate,runnershoot) values(DEFAULT,'" & txtItemid & "' " _
            & ",'" & txtmold & "'," & txtCavity & ",'" & txtSubcontDi & "'" _
            & "," & txtCT & "," & txtSecondProcess & "," & txtpriority & "," & txtManPower & "," & txtCavityStd & ",'" & kebijakanSC & "',TRUE,now()," & txtrunnershoot & ")"
        End If
    Else
        If kebijakanSC = "no" Then
            qry = "insert into loadcap_proc(idproclc,partno,mold_no,cavity,ct,ct_2,priorit,manpower,cavity_std,subcont,submch,timeupdate,runnershoot) values(DEFAULT,'" & txtItemid & "' " _
            & ",'" & txtmold & "'," & txtCavity & "" _
            & "," & txtCT & "," & txtSecondProcess & "," & txtpriority & "," & txtManPower & "," & txtCavityStd & ",'" & kebijakanSC & "',FALSE,now()," & txtrunnershoot & ")"
        Else
            qry = "insert into loadcap_proc(idproclc,partno,mold_no,cavity,prod_nomach,ct,ct_2,priorit,manpower,cavity_std,subcont,submch,timeupdate,runnershoot) values(DEFAULT,'" & txtItemid & "' " _
            & ",'" & txtmold & "'," & txtCavity & ",'" & txtSubcontDi & "'" _
            & "," & txtCT & "," & txtSecondProcess & "," & txtpriority & "," & txtManPower & "," & txtCavityStd & ",'" & kebijakanSC & "',TRUE,now()," & txtrunnershoot & ")"
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    txtSubcontDi = lv_sc.SelectedItem.Text
    txtMchineNo = txtSubcontDi
    Picture3.Visible = False
End Sub

Private Function checkPrimaryPN() As Boolean
    qry = "select count(*) from loadcap_mst_product_r where partno='" & Trim(txtItemid) & "'"
    Set RsBantu = Con.Execute(qry)
    If RsBantu(0) > 0 Then
        checkPrimaryPN = True
    Else
        checkPrimaryPN = False
    End If
End Function

Private Sub cmdSave_Click()
On Error GoTo ER_exc
    If cmbLabel.ListIndex < 0 Then cmbLabel.SetFocus: Exit Sub
    If cmdSave.Tag = "s" Then
        Const kolomI As String = "lc_idproduct,partno,partname,isno,shift_usg,hour_p_shift,faktor_productivity,ttlmold,qtylot,minstock,maxstock,typelabel,catgory,wght,typeboxid,typelabelbox,colordesc"
        If IsNumeric(txtTotalMold) = False _
            Or IsNumeric(txtHourPshift) = False _
            Or Len(txtItemName) < 2 Or IsNumeric(txtFaktorProd) = False _
        Then
            Exit Sub
        End If
        If checkPrimaryPN Then
            MsgBox "The Part Number is already registered", vbInformation
            Exit Sub
        End If
        If MsgBox("Save  ?", vbQuestion + vbYesNo) = vbYes Then
            BukaKoneksi
            qry = "insert into loadcap_mst_product_r (" & kolomI & ") values(DEFAULT,'" & txtItemid & "' " _
                & ",'" & txtItemName & "','" & txtISNo & "'" _
                & "," & txtShift & "," & txtHourPshift & "," & txtFaktorProd & "," & txtTotalMold & "" _
                & ",0,'" & txtMinStock & "','" & txtMaxStock & "'" _
                & ",'" & cmbLabel.ListIndex & "','" & cmbCategory & "'," & txtweight * 1 & ",'" & typebox & "','" & cmbBoxlabel.ListIndex & "','" & txtcolor & "')"
                Con.Execute qry
                MsgBox "Saved "
        End If
    Else
        If MsgBox("Update  ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        BukaKoneksi
        Set rsAneHelper = New ADODB.Recordset
        rsAneHelper.Open "select partno,isno,ttlmold,qtylot,minstock,maxstock,typelabel,catgory,partname,wght,typeboxid,typelabelbox,colordesc from loadcap_mst_product_r where lc_idproduct=" & id, Con, adOpenKeyset, adLockOptimistic, adCmdText
        If rsAneHelper("partno") = txtItemid Then
            rsAneHelper("isno") = txtISNo
            rsAneHelper("ttlmold") = txtTotalMold
            rsAneHelper("qtylot") = txtQty_lot
            rsAneHelper("minstock") = txtMinStock
            rsAneHelper("maxstock") = txtMaxStock
            rsAneHelper("typelabel") = cmbLabel.ListIndex
            rsAneHelper("typelabelbox") = cmbBoxlabel.ListIndex
            rsAneHelper("catgory") = cmbCategory
            rsAneHelper("partname") = txtItemName
            rsAneHelper("wght") = txtweight
            rsAneHelper("typeboxid") = typebox
            rsAneHelper("colordesc") = txtcolor
            
            rsAneHelper.Update
        End If
        MsgBox "Updated ", vbInformation, "Good !"
    End If
    updatesHIFT
    LoadDatanya
    Call getList
    Exit Sub
ER_exc:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub updatesHIFT()
    Dim aqry As String
    aqry = "update loadcap_mst_product_r set shift_usg=" & txtShift & ",hour_p_shift=" & txtHourPshift & ",faktor_productivity=" & txtFaktorProd
    Con.Execute aqry
End Sub

Private Sub LoadDatanya_V2()
    If Len(Trim(txtfind)) > 0 Then
        RsGet.Filter = "partno LIKE '*" & txtfind & "*'"
    Else
        RsGet.Filter = adFilterNone
    End If
    If RsGet.RecordCount > 0 Then
        Call getList
    Else
        RsGet.Filter = adFilterNone
        RsGet.Filter = "item_name LIKE '*" & txtfind & "*'"
        Call getList
    End If
    
End Sub

Private Sub LoadDatanya_v3()
    If Len(Trim(txtItemFind)) > 0 Then
        rsSubc.Filter = "namasubcont LIKE '*" & txtItemFind & "*'"
    Else
        rsSubc.Filter = adFilterNone
    End If
    Call getListSubcont
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrEx
    If MsgBox("Are you sure want to update ? ", vbQuestion + vbYesNo) = vbYes Then
        insertLogProc "update", id2
        qry = "update loadcap_proc set partno='" & txtItemid & "',priorit=" & txtpriority _
            & ",prod_nomach='" & txtMchineNo & "',mold_no='" & txtmold & "',timeupdate=now()" _
            & ",cavity=" & txtCavity & ",ct=" & txtCT & ",ct_2=" & txtSecondProcess _
            & ",manpower=" & txtManPower & ",cavity_std=" & txtCavityStd & ",runnershoot=" & txtrunnershoot & " where idproclc=" & id2
        Con.Execute qry
        qry = "update loadcap_proc set subcont='" & kebijakanSC & "' where mold_no='" & txtmold & "' AND submch=FALSE" 'partno='" & txtItemId & "' and
        Con.Execute qry
        MsgBox "Updated...", vbInformation, "Good !"
        With LV2
            .SelectedItem.SubItems(2) = txtManPower
            .SelectedItem.SubItems(3) = txtMchineNo
            .SelectedItem.SubItems(4) = txtmold
            .SelectedItem.SubItems(5) = txtCavity
            .SelectedItem.SubItems(6) = txtCavityStd
            .SelectedItem.SubItems(7) = txtCT
            .SelectedItem.SubItems(8) = txtSecondProcess
            .SelectedItem.SubItems(9) = txtpriority
            .SelectedItem.SubItems(10) = kebijakanSC
        End With
    End If
    LoadSubDatanya
    Exit Sub
ErrEx:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub Command1_Click()
    Picture3.Visible = False
End Sub

Private Sub Command2_Click()
    txtMchineNo = ""
    txtmold = ""
    txtCavity = ""
    txtCavityStd = ""
    txtpriority = ""
    txtCT = ""
    txtSecondProcess = ""
    txtManPower = ""
    txtSubcontDi = ""
    txtrunnershoot = 0
End Sub

Private Sub Command3_Click()
    Picmore.Visible = True
End Sub

Private Sub Command4_Click()
    On Error GoTo handlErr
    With CommonDialog1
        .Filter = ""
        .ShowOpen
        txturlfile = .FileName
    End With
    Dim file_name As String
    Dim fnum As Integer
    Dim whole_file As String
    Dim lines As Variant
    Dim one_line As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    Dim the_array() As String
    Dim r As Long
    Dim c As Long

    file_name = txturlfile
    If file_name = "" Then Exit Sub
    ' Load the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    ' Break the file into lines.
    lines = Split(whole_file, vbLf)

    ' Dimension the array.
    num_rows = UBound(lines)
    one_line = Split(lines(0), ",")
    num_cols = UBound(one_line)
    ReDim the_array(num_rows, num_cols)
    ' Copy the data into the array.
    Dim rskol() As String
    Dim LV As ListItem
    lvmore.ListItems.Clear
    For r = 1 To num_rows - 1
        Set LV = lvmore.ListItems.Add(, , r)
        rskol = Split(lines(r), ",")
        LV.SubItems(1) = rskol(0)
        If Len(rskol(1)) <> 0 Then
            If Right$(rskol(1), 2) = vbLf Or Right$(rskol(1), 2) = vbNewLine Then
                LV.SubItems(2) = Trim(Left$(rskol(1), Len(rskol(1)) - 2))
            Else
                LV.SubItems(2) = Trim(rskol(1))
            End If
        End If
        Dim kolx As String
        kolx = rskol(2)
        If Len(kolx) <> 0 Then
            If InStr(rskol(2), vbLf) Or InStr(rskol(2), vbCr) Then
                LV.SubItems(3) = Trim(Left$(kolx, Len(kolx) - 1))
            Else
                LV.SubItems(3) = Trim(kolx)
            End If
        End If
        kolx = rskol(3)
        If Len(kolx) <> 0 Then
            If InStr(rskol(3), vbLf) Or InStr(rskol(3), vbCr) Then
                LV.SubItems(4) = Trim(Left$(kolx, Len(kolx) - 1))
            Else
                LV.SubItems(4) = Trim(kolx)
            End If
        End If

    Next r
    Exit Sub
handlErr:
    MsgBox Err.Description

End Sub

Private Sub Command5_Click()
    If MsgBox("Are you sure ? ", vbQuestion + vbYesNo) = vbYes Then
        Dim u As Long
        Dim Ret As Byte
        With lvmore
            For u = 1 To .ListItems.Count
                qry = "UPDATE loadcap_mst_product_r set typelabel='" & .ListItems(u).SubItems(2) & "', typelabelbox='" & .ListItems(u).SubItems(3) & "', " _
                & " colordesc='" & .ListItems(u).SubItems(4) & "'" _
                & " WHERE partno='" & .ListItems(u).SubItems(1) & "'"
                Con.Execute qry, Ret
                .ListItems(u).SubItems(5) = Ret & " row updated"
            Next
        End With
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
    settingLV
    Me.Height = 9500
    Me.Width = 14160
    kebijakanSC = "no"
    LoadDatanya
    LoadDatanyaSubcont
    Call getList
    Call getListSubcont
    cmdSave.Tag = "s"
    idPROC = 0
    cmbCategory.ListIndex = 1
    cmbFiletype.ListIndex = 0
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub load()

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

Private Sub Form_Resize()
    ResizeControls
    cmbLabel.Left = txtItemid.Left
    cmbLabel.Top = Label21.Top
    cmbLabel.Width = txtItemid.Width
    
    cmbCategory.Left = txtMaxStock.Left
    cmbCategory.Top = Label22.Top
    cmbFiletype.Top = lblcmdtype2.Top
    cmbFiletype.Left = lblcmdtype2.Left
    cmbFiletype.Width = lblcmdtype2.Width
    cmbBoxlabel.Left = cmbLabel.Left
    cmbBoxlabel.Width = cmbLabel.Width
    cmbBoxlabel.Top = Label29.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If Cancel = 0 Then
        DelTab Me
    End If
End Sub

Private Sub Label31_Click()
    Picmore.Visible = False
End Sub

Private Sub Label35_Click()
PicCheckContainer.Visible = False
End Sub

Private Sub LV_Click()
    cmdSave.Tag = "s"
    lvToForm
    Picture3.Visible = False
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With LV
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

Private Sub LV_DblClick()
    cmdSave.Tag = "u"
    txtTotalMold.SetFocus
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 3
            Clipboard.Clear
            Clipboard.SetText LV.SelectedItem.SubItems(1)
    End Select
End Sub

Private Sub LV_KeyUp(KeyCode As Integer, Shift As Integer)
    cmdSave.Tag = "s"
    lvToForm
End Sub

Private Sub lv_sc_DblClick()
    cmdOK_Click
End Sub

Private Sub LV2_Click()
    lvtOform_v2
End Sub

Private Sub LV2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        PopupMenu MDI_Parent.mnuPopDet, vbPopupMenuRightButton
    End If
End Sub

Private Sub LV2_KeyUp(KeyCode As Integer, Shift As Integer)
    lvtOform_v2
End Sub

Private Sub lv2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDI_Parent.mnuPopDet, vbPopupMenuRightButton
    End If
End Sub



Private Sub Option1_Click()
    PicImportHeader.Visible = True
    picother.Visible = False
    
End Sub

Private Sub Option2_Click()
    PicImportHeader.Visible = False
    picother.Visible = True
End Sub

Private Sub txtCavity_Validate(Cancel As Boolean)
    If IsNumeric(txtCavity) = False Then Cancel = True
End Sub

Private Sub txtCavityStd_Validate(Cancel As Boolean)
     If IsNumeric(txtCavityStd) = False Then Cancel = True
End Sub

Private Sub txtCT_Validate(Cancel As Boolean)
    If IsNumeric(txtCT) = False Then Cancel = True
End Sub

Private Sub txtFaktorProd_Validate(Cancel As Boolean)
    If IsNumeric(txtFaktorProd) = False Then Cancel = True
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtfind = FilterIn(txtfind)
        LoadDatanya
        LoadDatanya_V2
    End If
End Sub

Private Sub txtHourPshift_Validate(Cancel As Boolean)
    If IsNumeric(txtHourPshift) = False Then Cancel = True
End Sub

Private Sub txtISNo_KeyPress(KeyAscii As Integer)
    If InStr(1, KARAKTERBAHAYA, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtItemFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtItemFind = FilterIn(txtItemFind)
        LoadDatanyaSubcont
        LoadDatanya_v3
    End If
End Sub

Private Sub txtManPower_Validate(Cancel As Boolean)
    If IsNumeric(txtManPower) = False Then Cancel = True
End Sub

Private Sub txtMaxStock_Validate(Cancel As Boolean)
    If IsNumeric(txtMaxStock) = False Then Cancel = True
End Sub

Private Sub txtMinStock_Validate(Cancel As Boolean)
    If IsNumeric(txtMinStock) = False Then Cancel = True
End Sub

Private Sub txtPriority_Validate(Cancel As Boolean)
   If IsNumeric(txtpriority) = False Then Cancel = True
End Sub

Private Sub txtQty_lot_Validate(Cancel As Boolean)
    If IsNumeric(txtQty_lot) = False Then Cancel = True
End Sub

Private Sub txtSecondProcess_Validate(Cancel As Boolean)
    If IsNumeric(txtSecondProcess) = False Then Cancel = True
End Sub

Private Sub txtShift_Validate(Cancel As Boolean)
    If IsNumeric(txtShift) = False Then Cancel = True
End Sub

Private Sub txtTotalMold_Validate(Cancel As Boolean)
    If IsNumeric(txtTotalMold) = False Then Cancel = True
End Sub

Private Sub txtweight_Validate(Cancel As Boolean)
    If IsNumeric(txtweight) = False Then Cancel = True
End Sub
