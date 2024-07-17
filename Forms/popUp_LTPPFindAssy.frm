VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form popUp_LTPPFindAssy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Assy"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "popUp_LTPPFindAssy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAssyNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "popUp_LTPPFindAssy.frx":000C
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin skn 
      Left            =   0
      OleObjectBlob   =   "popUp_LTPPFindAssy.frx":0072
      Top             =   120
   End
End
Attribute VB_Name = "popUp_LTPPFindAssy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
On Error GoTo errFind
    Form_GenerateLTPP.txtFindAssy.Text = RTrim(txtAssyNo)
    Unload Me
Exit Sub
errFind:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Find: " & Err.Number
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    activeTheme skn, Me
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub txtAssyNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdFind_Click
End If
End Sub
