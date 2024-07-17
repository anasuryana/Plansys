VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form_ODBC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting..."
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "Form_ODBC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.Skin FinsSkinner 
      Left            =   10320
      OleObjectBlob   =   "Form_ODBC.frx":BB04
      Top             =   360
   End
   Begin VB.PictureBox LODBC 
      BackColor       =   &H000040C0&
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   360
      Width           =   2775
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ODBC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox FrameODBC 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   360
      ScaleHeight     =   2475
      ScaleWidth      =   4995
      TabIndex        =   5
      Top             =   600
      Width           =   5055
      Begin VB.CommandButton cmdChange 
         BackColor       =   &H8000000B&
         Caption         =   "CHANGE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtNewODBC 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   0
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdTest2 
         BackColor       =   &H8000000B&
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtODBC 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdTest1 
         BackColor       =   &H8000000B&
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H8000000B&
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   " New ODBC"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   " ODBC"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form_ODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChange_Click()
    Call SaveINI("SETTING", "odbc", txtNewODBC.Text)
    MsgBox "Saved!", vbInformation, "Information..."
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdTest1_Click()
    On Error GoTo errHandler
    
    Set Con = New ADODB.Connection
    Con.Open txtODBC.Text
    If Con.State = 1 Then
        MsgBox "Database is Connected!", vbInformation, "Connected..."
    End If
    Exit Sub
errHandler:
    If Err.Number = -2147467259 Then
        MsgBox "Database is not connected!", vbCritical, "Not Connected..."
    End If
End Sub

Private Sub cmdTest2_Click()
    On Error GoTo errHandler
    
    Set Con = New ADODB.Connection
    Con.Open txtNewODBC.Text
    If Con.State = 1 Then
        MsgBox "Database is Connected!", vbInformation, "Connected..."
    End If
    Exit Sub
errHandler:
    If Err.Number = -2147467259 Then
        MsgBox "Database is not connected!", vbCritical, "Not Connected..."
    End If
End Sub

Private Sub Form_Load()
    Call activeTheme(FinsSkinner, Me)
    txtODBC.Text = GetINI("SETTING", "odbc", vbNullString)
End Sub

Private Sub txtNewODBC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdTest2_Click
End Sub
