VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form PopUp_PrinterPrint 
   Caption         =   "Choose Printer"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   Icon            =   "PopUp_PrinterPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPrinter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1400
      Left            =   8280
      Picture         =   "PopUp_PrinterPrint.frx":000C
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   2400
      Width           =   1400
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
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
      Left            =   8280
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin finsDeveloperSkin 
      Left            =   8400
      OleObjectBlob   =   "PopUp_PrinterPrint.frx":095F
      Top             =   1680
   End
   Begin VB.ListBox listPrinter 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT"
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
      Left            =   8280
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "PopUp_PrinterPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    pStPrinter = False
    Unload Me
End Sub

Private Sub cmdprint_Click()
On Error GoTo errSet
    If listPrinter.Text <> "" Then
        pGetPrinter = listPrinter.Text
'        Call setDataDefaultPrinter(pGetPrinter)
        Dim w As New WshNetwork
        w.SetDefaultPrinter pGetPrinter
        Set w = Nothing
        pStPrinter = True
        Unload Me
    Else
        MsgBox "Please, Choose Printer First!", vbExclamation, "Warning..."
    End If
Exit Sub
errSet:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error [" & Err.Number & "]"
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Dim r As Long
    Dim Buffer As String
    
    Call activeTheme(finsDeveloperSkin, Me)
    
    ' Get the list of available printers from WIN.INI
    Buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", _
       Buffer, Len(Buffer))

    ' Display the list of printer in the ListBox listPrinter
    ParseList listPrinter, Buffer
    listPrinter.Text = Printer.DeviceName
ErrHandler:
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub listPrinter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdprint_Click
    End If
End Sub
