VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form_Users 
   Caption         =   "User Permission"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   Icon            =   "Form_Users.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin skn 
      Left            =   9960
      OleObjectBlob   =   "Form_Users.frx":000C
      Top             =   5280
   End
   Begin VB.PictureBox FrameUser 
      BackColor       =   &H000040C0&
      Height          =   9375
      Left            =   120
      ScaleHeight     =   9315
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtEmpNo 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtEmpName 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   4335
      End
      Begin VB.CommandButton cmdPopID 
         Caption         =   "---"
         Height          =   495
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox FramePermission 
         BackColor       =   &H00FFFFFF&
         Height          =   7935
         Left            =   -120
         ScaleHeight     =   7875
         ScaleWidth      =   6435
         TabIndex        =   1
         Top             =   1440
         Width           =   6495
         Begin VB.CheckBox chkMnLoadCap 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Loading Capacity"
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
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   2175
         End
         Begin VB.PictureBox LPermission 
            BackColor       =   &H000040C0&
            Height          =   615
            Left            =   120
            ScaleHeight     =   555
            ScaleWidth      =   2355
            TabIndex        =   5
            Top             =   120
            Width           =   2415
            Begin VB.Label Label 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "PERMISSION"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   120
               TabIndex        =   6
               Top             =   120
               Width           =   2055
            End
         End
         Begin VB.CheckBox chkMnLTPP 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LTPP"
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
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox chkSubLTPP 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Generate LTPP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "UPDATE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4920
            TabIndex        =   2
            Top             =   7080
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000040C0&
            BorderWidth     =   3
            X1              =   2280
            X2              =   6480
            Y1              =   360
            Y2              =   360
         End
      End
      Begin ACTIVESKINLibCtl.Skin FinsDevSkinner 
         Left            =   -240
         OleObjectBlob   =   "Form_Users.frx":0240
         Top             =   -1000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo errLoad
    Call activeTheme(skn, Me)
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub
