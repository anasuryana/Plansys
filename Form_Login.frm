VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Login 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Login - PLANSYS"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   Icon            =   "Form_Login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6735
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox LBLOGIN 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6735
      TabIndex        =   6
      Top             =   0
      Width           =   6735
      Begin VB.Label LBMD 
         BackStyle       =   0  'Transparent
         Caption         =   "PLANNING SYSTEM - INJECTION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label CLOSEFORM 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5880
         TabIndex        =   17
         Top             =   0
         Width           =   735
      End
      Begin VB.Label tlogin 
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.PictureBox cmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   5040
      ScaleHeight     =   540
      ScaleWidth      =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   1440
   End
   Begin VB.PictureBox picBANSHU 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   5160
      Picture         =   "Form_Login.frx":4C4A
      ScaleHeight     =   1335
      ScaleWidth      =   1335
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.PictureBox FrameLogin 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2535
      ScaleWidth      =   6495
      TabIndex        =   5
      Top             =   600
      Width           =   6495
      Begin VB.PictureBox picPASS 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4320
         Picture         =   "Form_Login.frx":596E
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox picUSER 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4320
         Picture         =   "Form_Login.frx":6299
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         Height          =   615
         Left            =   1080
         ScaleHeight     =   555
         ScaleWidth      =   3075
         TabIndex        =   13
         Top             =   480
         Width           =   3135
         Begin VB.TextBox txtUsername 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   480
            Left            =   120
            TabIndex        =   0
            Top             =   30
            Width           =   2835
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFC0C0&
         Height          =   615
         Left            =   1080
         ScaleHeight     =   555
         ScaleWidth      =   3075
         TabIndex        =   14
         Top             =   1080
         Width           =   3135
         Begin VB.TextBox txtPassword 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   30
            Width           =   2865
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PLANSYS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label msgLOGIN 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PASS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.PictureBox FrameLine 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2775
      ScaleWidth      =   6735
      TabIndex        =   7
      Top             =   480
      Width           =   6735
   End
   Begin MSComctlLib.ImageList ImageListButton 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   96
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Login.frx":6ACD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Login.frx":72B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   480
      Picture         =   "Form_Login.frx":7B07
      ScaleHeight     =   2535
      ScaleWidth      =   4575
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " MRP Barcode System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   2280
         Width           =   1815
      End
   End
   Begin ACTIVESKINLibCtl.Skin FinsDeveloperSkin 
      Left            =   -360
      OleObjectBlob   =   "Form_Login.frx":9758
      Top             =   -360
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSukses As Boolean

Private Sub CLOSEFORM_Click()
Unload Me
End Sub
Private Sub Form_Initialize()
    If App.PrevInstance Then
        End
    Else
        Me.Show
    End If
End Sub
Private Sub cmdLogin_Click()
    On Error GoTo ErrHandler
    
    If txtUsername <> "" And txtPassword <> "" Then
        If txtUsername = "setodbc" And txtPassword = "setodbc" Then
            Form_ODBC.Show 1
        Else
            Call BukaKoneksi
            Call selectDB
            RsDB.Open "select u.empno, u.password, m.acc_admin, m.acc_por, m.acc_sin from cd_users u, cd_users_me m where u.empno = m.empno", Con, adOpenDynamic, adLockOptimistic
            RsDB.MoveFirst
            RsDB.Find "empno = '" & txtUsername.Text & "'", , adSearchForward, 1
            If Not RsDB.EOF Then
                If txtUsername = RsDB.Fields(0) And txtPassword = RsDB.Fields(1) Then
                    LoginSukses = True
                    Me.Hide
                    MDI_Parent.Show
                Else
                    msgLOGIN = "Username dan Password Tidak Valid!"
                    txtUsername.SetFocus
                End If
            Else
                msgLOGIN = "Username Tidak Ditemukan atau Tidak Memiliki Akses."
                txtUsername.SetFocus
            End If
        End If
    Else
        'MsgBox "Isi username dan password anda!", vbInformation, "Informasi..."
        msgLOGIN = "Isi username dan password anda!"
    End If
    Exit Sub
ErrHandler:
    If Err.Number = -2147467259 Then
        MsgBox "Database is not connected. Please contact your administrator!", vbCritical, "Fatal Error..."
        End
    Else
        MsgBox Err.Description, vbCritical, "Error Login... [" & Err.Number & "]"
    End If
End Sub

Private Sub cmdLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set cmdLogin.Picture = ImageListButton.ListImages(2).Picture
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    'Call ExplodeForm(Me, 10000)
    myTemplates
    Set cmdLogin.Picture = ImageListButton.ListImages(1).Picture
    Call activeTheme(finsDeveloperSkin, Me)
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load [" & Err.Number & "]"
        Unload Me
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set cmdLogin.Picture = ImageListButton.ListImages(1).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
       LoginSukses = False
End Sub

Private Sub FrameLine_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set cmdLogin.Picture = ImageListButton.ListImages(1).Picture
End Sub

Private Sub FrameLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set cmdLogin.Picture = ImageListButton.ListImages(1).Picture
End Sub

Private Sub LBLOGIN_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim ReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   ReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   txtUsername.SetFocus
End If
End Sub

Private Sub LBMD_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim ReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   ReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   txtUsername.SetFocus
End If
End Sub


Private Sub picBANSHU_DblClick()
    Form_ODBC.Show 1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set cmdLogin.Picture = ImageListButton.ListImages(1).Picture
End Sub

Private Sub tlogin_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim ReturnValue As Long
If Button = 1 Then
   Call ReleaseCapture
   ReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   txtUsername.SetFocus
End If
End Sub

Private Sub txtPassword_Change()
    msgLOGIN = ""
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdLogin_Click
    End If
End Sub

Private Sub txtUsername_Change()
    msgLOGIN = ""
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        txtPassword.SetFocus
    End If
End Sub
