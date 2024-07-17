VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form F_GenerateLoadCap 
   Caption         =   "Generate Loadcap"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
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
   ScaleHeight     =   7185
   ScaleWidth      =   11280
   Begin MSFlexGridLib.MSFlexGrid anaGrid 
      Height          =   5415
      Left            =   45
      TabIndex        =   6
      Top             =   1680
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   9551
      _Version        =   393216
      MergeCells      =   1
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
   Begin VB.ComboBox CmbDocument 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txtRevision 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
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
      Format          =   145489923
      CurrentDate     =   42544
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_GenerateLoadCap.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin skinFD 
      Left            =   9240
      OleObjectBlob   =   "F_GenerateLoadCap.frx":0064
      Top             =   1080
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_GenerateLoadCap.frx":0298
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "F_GenerateLoadCap.frx":02FE
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "F_GenerateLoadCap"
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
Private idS As String
Private rsA As ADODB.Recordset
Private rsB As ADODB.Recordset
Private nm_msn_full() As String

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

Private Sub settingFG()
    Dim i As Integer
    With anaGrid
        .Cols = 56: .ColWidth(0) = 700: .ColWidth(1) = 2800: .ColWidth(2) = 3000: .ColWidth(3) = 3000
        .rows = 5
        .FixedRows = 3
        .FixedCols = 4
        .WordWrap = True
        .ColAlignment(2) = flexAlignLeftCenter
        .ColWidth(7) = 3000
        
        .MergeCells = flexMergeFree
        i = 0
        .TextMatrix(0, i) = "No":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 1
        .TextMatrix(0, i) = "Customer":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 2
        .TextMatrix(0, i) = "Assy no":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 3
        .TextMatrix(0, i) = "Assy Desc":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 4
        .TextMatrix(0, i) = "STOCK FG":         .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 5
        .TextMatrix(0, i) = "STOCK WIP":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 6
        .TextMatrix(0, i) = "FC":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 7
        .TextMatrix(0, i) = "ITO":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 8
        .TextMatrix(0, i) = "SUBCONT":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 9
        .TextMatrix(0, i) = "PROD PLAN 1":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 10
        .TextMatrix(0, i) = "PROD PLAN 2":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 11
        .TextMatrix(0, i) = "PROD PLAN 3":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 12
        .TextMatrix(0, i) = "PROD PLAN 4":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 13
        .TextMatrix(0, i) = "Cav":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 14
        .TextMatrix(0, i) = "C/T":        .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 15
        .TextMatrix(0, i) = "Man Power":          .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 16
        .TextMatrix(0, i) = "2nd Proses":          .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 17
        .TextMatrix(0, i) = "Cap/day":         .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 18
        .TextMatrix(0, i) = "Cap/Month":         .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 19
        .TextMatrix(0, i) = "Need day": .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 20
        .TextMatrix(0, i) = "MC No":        .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 21
        .TextMatrix(0, i) = "Tonage":        .TextMatrix(1, i) = .TextMatrix(0, i):         .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 22
        .TextMatrix(0, i) = "%": .TextMatrix(1, i) = .TextMatrix(0, i):     .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 23
        .TextMatrix(0, i) = "Ovrday": .TextMatrix(1, i) = .TextMatrix(0, i):     .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 24
        .TextMatrix(0, i) = "Alternative 1": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 25
        .TextMatrix(0, i) = "Tonage": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 26
        .TextMatrix(0, i) = "%": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 27
        .TextMatrix(0, i) = "Ovrday": .TextMatrix(1, i) = .TextMatrix(0, i):     .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
         i = 28
        .TextMatrix(0, i) = "Alternative 2": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 29
        .TextMatrix(0, i) = "Tonage": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 30
        .TextMatrix(0, i) = "%": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        i = 31
        .TextMatrix(0, i) = "Ovrday": .TextMatrix(1, i) = .TextMatrix(0, i):     .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

         i = 32
        .TextMatrix(0, i) = "Alternative 3": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 33
        .TextMatrix(0, i) = "Tonage": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 34
        .TextMatrix(0, i) = "%": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

         i = 35
        .TextMatrix(0, i) = "Ovrday": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 36
        .TextMatrix(0, i) = "Alternative 4": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 37
        .TextMatrix(0, i) = "Tonage": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

         i = 38
        .TextMatrix(0, i) = "%": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 39
        .TextMatrix(0, i) = "Ovrday": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 40
        .TextMatrix(0, i) = "Alternative 5": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

         i = 41
        .TextMatrix(0, i) = "Tonage": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 42
        .TextMatrix(0, i) = "%": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 43
        .TextMatrix(0, i) = "Ovrday": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

         i = 44
        .TextMatrix(0, i) = "Alternative 6": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

         i = 45
        .TextMatrix(0, i) = "Tonage": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 46
        .TextMatrix(0, i) = "%": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True

        i = 47
        .TextMatrix(0, i) = "Ovrday": .TextMatrix(1, i) = .TextMatrix(0, i):        .TextMatrix(2, i) = .TextMatrix(0, i)
        .MergeCol(i) = True
        
        
        .TextMatrix(0, 48) = "Load Vs Cap Mahine"
        .TextMatrix(0, 49) = .TextMatrix(0, 48)
        .TextMatrix(0, 50) = .TextMatrix(0, 48)
        .TextMatrix(0, 51) = .TextMatrix(0, 48)
        .MergeCol(48) = True
        .MergeCol(49) = True
        .MergeCol(50) = True
        .MergeCol(51) = True
'        .TextMatrix(1, 41) = "April"
'        .TextMatrix(1, 42) = "Mei"
'        .TextMatrix(1, 43) = "Juni"
'        .TextMatrix(1, 44) = "Juli"
'        .TextMatrix(2, 41) = "April"
'        .TextMatrix(2, 42) = "Mei"
'        .TextMatrix(2, 43) = "Juni"
'        .TextMatrix(2, 44) = "Juli"
        .TextMatrix(0, 52) = "Need of Operator"
        .TextMatrix(0, 53) = .TextMatrix(0, 52)
        .TextMatrix(0, 54) = .TextMatrix(0, 52)
        .TextMatrix(0, 55) = .TextMatrix(0, 52)
        .MergeCol(52) = True
        .MergeCol(53) = True
        .MergeCol(54) = True
        .MergeCol(55) = True
'        .TextMatrix(1, 45) = "April"
'        .TextMatrix(1, 46) = "Mei"
'        .TextMatrix(1, 47) = "Juni"
'        .TextMatrix(1, 48) = "Juli"
'        .TextMatrix(2, 45) = "April"
'        .TextMatrix(2, 46) = "Mei"
'        .TextMatrix(2, 47) = "Juni"
'        .TextMatrix(2, 48) = "Juli"
        
        .MergeRow(0) = True
        .MergeRow(2) = True
    End With
End Sub

Private Sub settingGridName()
    Dim nmBulan() As String, it As Integer
    ReDim nmBulan(1 To 12) As String
    nmBulan(1) = "Januari"
    nmBulan(2) = "Februari"
    nmBulan(3) = "Maret"
    nmBulan(4) = "April"
    nmBulan(5) = "Mei"
    nmBulan(6) = "Juni"
    nmBulan(7) = "Juli"
    nmBulan(8) = "Agustus"
    nmBulan(9) = "September"
    nmBulan(10) = "Oktober"
    nmBulan(11) = "November"
    nmBulan(12) = "Desember"
    
    With anaGrid
        For it = 0 To 3
            .TextMatrix(1, 48 + it) = nmBulan(Val(Format(DTPicker1, "M")) + it)
            .TextMatrix(2, 48 + it) = nmBulan(Val(Format(DTPicker1, "M")) + it)
            
            .TextMatrix(1, 52 + it) = nmBulan(Val(Format(DTPicker1, "M")) + it)
            .TextMatrix(2, 52 + it) = nmBulan(Val(Format(DTPicker1, "M")) + it)
        Next
    End With
End Sub

Public Function DaysInMonth(ByVal dDate As Date) As Integer
    DaysInMonth = Day(DateAdd("m", 1, dDate - Day(dDate) + 1) - 1)
End Function

Private Sub prosesSisa(barisKe As Integer, present1 As Variant, pneday As Variant, phkw As Variant, pcapday As Variant)
'    MsgBox nilaiNeedDay
    Dim ulang As Integer
    Dim inSisa As Variant
        For ulang = 3 To barisKe
            If anaGrid.TextMatrix(barisKe, 20) = anaGrid.TextMatrix(ulang, 20) Then
                If anaGrid.TextMatrix(barisKe, 20) <> "" Then
                    MsgBox "sama nih_" & "(" & barisKe & "," & 20 & ")" & "dengan" & "(" & ulang & "," & "xx)"
                    
                    If present1 > 100 Then
                        MsgBox "tah"
                        anaGrid.TextMatrix(barisKe, 22) = FormatNumber(100, 2)
                        
                        anaGrid.TextMatrix(barisKe, 23) = FormatNumber(pneday - phkw, 2)
                        anaGrid.TextMatrix(barisKe, 26) = FormatNumber((pneday - phkw) / phkw * 100, 2)
                    Else
                           
                    End If
                End If
            ElseIf anaGrid.TextMatrix(barisKe, 20) = anaGrid.TextMatrix(ulang, 24) Then
                MsgBox "aha"
            ElseIf anaGrid.TextMatrix(barisKe, 20) = anaGrid.TextMatrix(ulang, 28) Then
            
            ElseIf anaGrid.TextMatrix(barisKe, 20) = anaGrid.TextMatrix(ulang, 32) Then
            
            ElseIf anaGrid.TextMatrix(barisKe, 20) = anaGrid.TextMatrix(ulang, 36) Then
            
            ElseIf anaGrid.TextMatrix(barisKe, 20) = anaGrid.TextMatrix(ulang, 40) Then
            
            ElseIf anaGrid.TextMatrix(barisKe, 20) = anaGrid.TextMatrix(ulang, 44) Then
            
            End If
            
            
'            If anaGrid.TextMatrix(barisKe, 24) = anaGrid.TextMatrix(ulang, 24) Or _
'            anaGrid.TextMatrix(barisKe, 24) = anaGrid.TextMatrix(ulang, 20) Or _
'            anaGrid.TextMatrix(barisKe, 24) = anaGrid.TextMatrix(ulang, 28) Or _
'            anaGrid.TextMatrix(barisKe, 24) = anaGrid.TextMatrix(ulang, 32) Or _
'            anaGrid.TextMatrix(barisKe, 24) = anaGrid.TextMatrix(ulang, 36) Or _
'            anaGrid.TextMatrix(barisKe, 24) = anaGrid.TextMatrix(ulang, 40) Or _
'            anaGrid.TextMatrix(barisKe, 24) = anaGrid.TextMatrix(ulang, 44) _
'            Then
'                If anaGrid.TextMatrix(barisKe, 24) <> "" Then
'                    MsgBox "sama nih_" & "(" & barisKe & "," & 24 & ")" & "dengan" & "(" & ulang & "," & "xx)"
'                End If
''                anaGrid.TextMatrix(barisKe, 22) = 100
''                anaGrid.TextMatrix(barisKe, 23) = FormatNumber(pneday - phkw, 2)
''                anaGrid.TextMatrix(barisKe, 26) = FormatNumber((pneday - phkw) / phkw * 100, 2)
'            End If
'            If anaGrid.TextMatrix(barisKe, 28) = anaGrid.TextMatrix(ulang, 28) Or _
'            anaGrid.TextMatrix(barisKe, 28) = anaGrid.TextMatrix(ulang, 24) Or _
'            anaGrid.TextMatrix(barisKe, 28) = anaGrid.TextMatrix(ulang, 20) Or _
'            anaGrid.TextMatrix(barisKe, 28) = anaGrid.TextMatrix(ulang, 32) Or _
'            anaGrid.TextMatrix(barisKe, 28) = anaGrid.TextMatrix(ulang, 36) Or _
'            anaGrid.TextMatrix(barisKe, 28) = anaGrid.TextMatrix(ulang, 40) Or _
'            anaGrid.TextMatrix(barisKe, 28) = anaGrid.TextMatrix(ulang, 44) _
'            Then
'                If anaGrid.TextMatrix(barisKe, 28) <> "" Then
'                    MsgBox "sama nih_" & "(" & barisKe & "," & 25 & ")" & "dengan" & "(" & ulang & "," & "xx)"
'                End If
''                anaGrid.TextMatrix(barisKe, 22) = 100
''                anaGrid.TextMatrix(barisKe, 23) = FormatNumber(pneday - phkw, 2)
''                anaGrid.TextMatrix(barisKe, 26) = FormatNumber((pneday - phkw) / phkw * 100, 2)
'            End If
        Next
    
End Sub

Private Sub gridFormatNum()
    Dim v As Integer
    For v = 3 To anaGrid.rows - 1
        With anaGrid
            .TextMatrix(v, 6) = FormatNumber(.TextMatrix(v, 6), 0)
            .TextMatrix(v, 7) = FormatNumber(.TextMatrix(v, 7), 4)
            .TextMatrix(v, 9) = FormatNumber(.TextMatrix(v, 9), 0)
            .TextMatrix(v, 10) = FormatNumber(.TextMatrix(v, 10), 0)
            .TextMatrix(v, 11) = FormatNumber(.TextMatrix(v, 11), 0)
            .TextMatrix(v, 12) = FormatNumber(.TextMatrix(v, 12), 0)
            .TextMatrix(v, 17) = FormatNumber(.TextMatrix(v, 17), 0)
            .TextMatrix(v, 18) = FormatNumber(.TextMatrix(v, 18), 2)
            .TextMatrix(v, 19) = FormatNumber(.TextMatrix(v, 19), 2)
            .TextMatrix(v, 22) = FormatNumber(.TextMatrix(v, 22), 2)
            .TextMatrix(v, 23) = FormatNumber(.TextMatrix(v, 23), 2)
            .TextMatrix(v, 26) = FormatNumber(.TextMatrix(v, 26), 2)
            .TextMatrix(v, 27) = FormatNumber(.TextMatrix(v, 27), 2)
            .TextMatrix(v, 30) = FormatNumber(.TextMatrix(v, 30), 2)
            .TextMatrix(v, 31) = FormatNumber(.TextMatrix(v, 31), 2)
            .TextMatrix(v, 34) = FormatNumber(.TextMatrix(v, 34), 2)
            .TextMatrix(v, 35) = FormatNumber(.TextMatrix(v, 35), 2)
            .TextMatrix(v, 38) = FormatNumber(.TextMatrix(v, 38), 2)
            .TextMatrix(v, 39) = FormatNumber(.TextMatrix(v, 39), 2)
            .TextMatrix(v, 42) = FormatNumber(.TextMatrix(v, 42), 2)
            .TextMatrix(v, 43) = FormatNumber(.TextMatrix(v, 43), 2)
            .TextMatrix(v, 46) = FormatNumber(.TextMatrix(v, 46), 2)
            .TextMatrix(v, 47) = FormatNumber(.TextMatrix(v, 47), 2)
        End With
    Next
End Sub

Private Sub CmbDocument_Click()
    Dim i As Integer, c_wip As Variant, c_cap_p_day As Variant, j As Integer, x As Integer, k As Integer, t As Integer
    Dim c_part() As String, tem_mesin As String
    Dim totalp_mesin As Variant
    Dim presentMesinUse As Variant
    Dim ar_mesin_present() As String
    Dim ar_mesin() As String, ar_mesin_alt1() As String
    Dim ar_mesin_ovl() As String
    Dim msnutama As Variant, msnalt1 As Variant, msnalt2 As Variant, msnalt3 As Variant, msnalt4 As Variant
    Dim msnalt5 As Variant, msnalt6 As Variant
    Dim ovrd_msnutama As Variant, ovrd_msnalt1 As Variant, ovrd_msnalt2 As Variant, ovrd_msnalt3 As Variant, ovrd_msnalt4 As Variant
    Dim ovrd_msnalt5 As Variant, ovrd_msnalt6 As Variant
    Dim hkw1 As Variant
    qry = "select cust_name,assy_no,a.item_name,fg,p1,p2,p3,fc1,kebijkan_subc " _
        & " ,prod_plan_1,prod_plan_2,prod_plan_3,prod_plan_4 " _
        & " ,d.cavity,cycletime,manpower,time_sec_proc,prod_nomach " _
        & " ,e.tonage_mach,a.hkw_1,alt1_prod_nomach,ma1.tonage_mach talt1,(prod_plan_1/((60 / cycletime) * d.cavity * 7 * 3 * 60 )*faktor_productivity)/a.hkw_1*100 presenku " _
        & " ,faktor_productivity,alt2_prod_nomach,ma2.tonage_mach talt2 " _
        & " ,alt3_prod_nomach,ma3.tonage_mach talt3,alt4_prod_nomach,ma4.tonage_mach talt4 " _
        & " ,alt5_prod_nomach,ma5.tonage_mach talt5,alt6_prod_nomach,ma6.tonage_mach talt6 " _
        & " ,e.state_mach smch,coalesce(ma1.state_mach,FALSE,ma1.state_mach) sma1" _
        & " ,coalesce(ma2.state_mach,FALSE,ma2.state_mach) sma2,coalesce(ma3.state_mach,FALSE,ma3.state_mach) sma3" _
        & " ,coalesce(ma4.state_mach,FALSE,ma4.state_mach) sma4,coalesce(ma5.state_mach,FALSE,ma5.state_mach) sma5" _
        & " ,coalesce(ma6.state_mach,FALSE,ma6.state_mach) sma6" _
        & " from ltpp_generate a " _
        & " inner join mst_item b on a.assy_no=b.item_id " _
        & " inner join r_customer c on b.cust_id=c.cust_id " _
        & " inner join loadcap_mst_product d on a.assy_no=d.partno" _
        & " left join loadcap_mst_mach e on d.prod_nomach=e.no_mach" _
        & " inner join ltpp_header f on a.ltpp_doc=f.ltpp_doc" _
        & " left join loadcap_mst_mach ma1 on d.alt1_prod_nomach=ma1.no_mach" _
        & " left join loadcap_mst_mach ma2 on d.alt2_prod_nomach=ma2.no_mach" _
        & " left join loadcap_mst_mach ma3 on d.alt3_prod_nomach=ma3.no_mach" _
        & " left join loadcap_mst_mach ma4 on d.alt4_prod_nomach=ma4.no_mach" _
        & " left join loadcap_mst_mach ma5 on d.alt5_prod_nomach=ma5.no_mach" _
        & " left join loadcap_mst_mach ma6 on d.alt6_prod_nomach=ma6.no_mach" _
        & " where a.ltpp_doc='" & CmbDocument & "' " _
        & " order by kebijkan_subc asc,23 desc, prod_nomach asc"
    Set rsB = Con.Execute(qry)
    ReDim nm_msn_full(1) As String
    If rsB.RecordCount > 0 Then
        settingGridName
        anaGrid.rows = 3 + rsB.RecordCount
        ReDim ar_mesin_present(1 To rsB.RecordCount)
        ReDim ar_mesin(1 To rsB.RecordCount)
        ReDim ar_mesin_alt1(1 To rsB.RecordCount)
        While Not rsB.EOF
            c_wip = rsB("p1") + rsB("p2") + rsB("p3")
            c_cap_p_day = ((60 / rsB("cycletime")) * rsB("cavity") * 7 * 3 * 60) * rsB("faktor_productivity")
            hkw1 = rsB("hkw_1")
'            If tem_mesin = rsB("prod_nomach") Then
'                presentMesinUse = ((rsB("prod_plan_1") / c_cap_p_day) / rsB("hkw_1") * 100) + presentMesinUse
''                MsgBox "oke tambahkan ke " & rsB("assy_no")
'            Else
                presentMesinUse = rsB("presenku") '(rsB("prod_plan_1") / c_cap_p_day) / rsB("hkw_1") * 100
'            End If
'            MsgBox presentMesinUse & "_" & (rsB("prod_plan_1") / c_cap_p_day) / hkw1
            If presentMesinUse > 100 Then
                If rsB("smch") = 1 Then
                    msnutama = 100
                    ovrd_msnutama = (rsB("prod_plan_1") / c_cap_p_day) - hkw1
                Else
                    msnutama = 0
                    ovrd_msnutama = (rsB("prod_plan_1") / c_cap_p_day)
                End If
            Else
                If rsB("smch") = 1 Then
                    msnutama = presentMesinUse
                    ovrd_msnutama = 0
                Else
                    msnutama = 0
                    ovrd_msnutama = (rsB("prod_plan_1") / c_cap_p_day)
                End If
            End If
            If (ovrd_msnutama / hkw1) * 100 > 100 Then
                If rsB("sma1") = 1 Then
                    msnalt1 = 100
                    ovrd_msnalt1 = ovrd_msnutama - hkw1
                Else
                    msnalt1 = 0
                    ovrd_msnalt1 = ovrd_msnutama
                End If
            Else
                If rsB("sma1") = 1 Then
                    msnalt1 = (ovrd_msnutama / hkw1) * 100
                    ovrd_msnalt1 = ovrd_msnutama - hkw1
                Else
                    msnalt1 = 0
                    ovrd_msnalt1 = ovrd_msnutama
                End If
            End If
            If (ovrd_msnalt1 / hkw1) * 100 > 100 Then
                If rsB("sma2") = 1 Then
                    msnalt2 = 100
                    ovrd_msnalt2 = ovrd_msnalt1 - hkw1
                Else
                    msnalt2 = 0
                    ovrd_msnalt2 = ovrd_msnalt1
                End If
            Else
                If rsB("sma2") = 1 Then
                    msnalt2 = (ovrd_msnalt1 / hkw1) * 100
                    ovrd_msnalt2 = ovrd_msnalt1 - hkw1
                Else
                    msnalt2 = 0
                    ovrd_msnalt2 = ovrd_msnalt1
                End If
            End If
            If (ovrd_msnalt2 / hkw1) * 100 > 100 Then
                If rsB("sma3") = 1 Then
                    msnalt3 = 100
                    ovrd_msnalt3 = ovrd_msnalt2 - hkw1
                Else
                    msnalt3 = 0
                    ovrd_msnalt3 = ovrd_msnalt2
                End If
            Else
                If rsB("sma3") = 1 Then
                    msnalt3 = (ovrd_msnalt2 / hkw1) * 100
                    ovrd_msnalt3 = ovrd_msnalt2 - hkw1
                Else
                    msnalt3 = 0
                    ovrd_msnalt3 = ovrd_msnalt2
                End If
            End If
            If (ovrd_msnalt3 / hkw1) * 100 > 100 Then
                If rsB("sma4") = 1 Then
                    msnalt4 = 100
                    ovrd_msnalt4 = ovrd_msnalt3 - hkw1
                Else
                    msnalt4 = 0
                    ovrd_msnalt4 = ovrd_msnalt3
                End If
            Else
                If rsB("sma4") = 1 Then
                    msnalt4 = (ovrd_msnalt3 / hkw1) * 100
                    ovrd_msnalt4 = ovrd_msnalt3 - hkw1
                Else
                    msnalt4 = 0
                    ovrd_msnalt4 = ovrd_msnalt3
                End If
            End If
            If (ovrd_msnalt4 / hkw1) * 100 > 100 Then
                If rsB("sma5") = 1 Then
                    msnalt5 = 100
                    ovrd_msnalt5 = ovrd_msnalt4 - hkw1
                Else
                    msnalt5 = 0
                    ovrd_msnalt5 = ovrd_msnalt4
                End If
            Else
                If rsB("sma5") = 1 Then
                    msnalt5 = (ovrd_msnalt4 / hkw1) * 100
                    ovrd_msnalt5 = ovrd_msnalt4 - hkw1
                Else
                    msnalt5 = 0
                    ovrd_msnalt5 = ovrd_msnalt4
                End If
            End If
            If (ovrd_msnalt5 / hkw1) * 100 > 100 Then
                If rsB("sma6") = 1 Then
                    msnalt6 = 100
                    ovrd_msnalt6 = ovrd_msnalt5 - hkw1
                Else
                    msnalt6 = 0
                    ovrd_msnalt6 = ovrd_msnalt5
                End If
            Else
                If rsB("sma6") = 1 Then
                    msnalt6 = (ovrd_msnalt5 / hkw1) * 100
                    ovrd_msnalt6 = ovrd_msnalt5 - hkw1
                Else
                    msnalt6 = 0
                    ovrd_msnalt6 = ovrd_msnalt5
                End If
            End If
            anaGrid.TextMatrix(3 + i, 0) = i + 1
            anaGrid.TextMatrix(3 + i, 1) = rsB(0)
            anaGrid.TextMatrix(3 + i, 2) = rsB(1)
            anaGrid.TextMatrix(3 + i, 3) = rsB(2)
            anaGrid.TextMatrix(3 + i, 4) = rsB(3)
            anaGrid.TextMatrix(3 + i, 5) = c_wip
            anaGrid.TextMatrix(3 + i, 6) = rsB("fc1")
            
            If rsB("fc1") = 0 Then
                anaGrid.TextMatrix(3 + i, 7) = 0
            Else
                anaGrid.TextMatrix(3 + i, 7) = (rsB(3) + c_wip) / rsB("fc1")
            End If
            
            anaGrid.TextMatrix(3 + i, 8) = rsB("kebijkan_subc")
            anaGrid.TextMatrix(3 + i, 9) = rsB("prod_plan_1")
            anaGrid.TextMatrix(3 + i, 10) = rsB("prod_plan_2")
            anaGrid.TextMatrix(3 + i, 11) = rsB("prod_plan_3")
            anaGrid.TextMatrix(3 + i, 12) = rsB("prod_plan_4")
            anaGrid.TextMatrix(3 + i, 13) = rsB("cavity")
            anaGrid.TextMatrix(3 + i, 14) = rsB("cycletime")
            anaGrid.TextMatrix(3 + i, 15) = rsB("manpower")
            anaGrid.TextMatrix(3 + i, 16) = rsB("time_sec_proc")
            anaGrid.TextMatrix(3 + i, 17) = c_cap_p_day
            anaGrid.TextMatrix(3 + i, 18) = c_cap_p_day * hkw1
            
            
            anaGrid.TextMatrix(3 + i, 19) = rsB("prod_plan_1") / c_cap_p_day
            anaGrid.TextMatrix(3 + i, 20) = rsB("prod_nomach")
            anaGrid.TextMatrix(3 + i, 21) = rsB("tonage_mach")
            anaGrid.TextMatrix(3 + i, 22) = IIf(msnutama < 0, 0, msnutama)
            anaGrid.TextMatrix(3 + i, 23) = IIf(ovrd_msnutama < 0, 0, ovrd_msnutama)
            
            anaGrid.TextMatrix(3 + i, 24) = rsB("alt1_prod_nomach")
            anaGrid.TextMatrix(3 + i, 25) = IIf(IsNull(rsB("talt1")), "", rsB("talt1"))
            anaGrid.TextMatrix(3 + i, 26) = IIf(msnalt1 < 0, 0, msnalt1)
            anaGrid.TextMatrix(3 + i, 27) = IIf(ovrd_msnalt1 < 0, 0, ovrd_msnalt1)
            
            anaGrid.TextMatrix(3 + i, 28) = rsB("alt2_prod_nomach")
            anaGrid.TextMatrix(3 + i, 29) = IIf(IsNull(rsB("talt2")), "", rsB("talt2"))
            anaGrid.TextMatrix(3 + i, 30) = IIf(msnalt2 < 0, 0, msnalt2)
            anaGrid.TextMatrix(3 + i, 31) = IIf(ovrd_msnalt2 < 0, 0, ovrd_msnalt2)
            
            anaGrid.TextMatrix(3 + i, 32) = rsB("alt3_prod_nomach")
            anaGrid.TextMatrix(3 + i, 33) = IIf(IsNull(rsB("talt3")), "", rsB("talt3"))
            anaGrid.TextMatrix(3 + i, 34) = IIf(msnalt3 < 0, 0, msnalt3)
            anaGrid.TextMatrix(3 + i, 35) = IIf(ovrd_msnalt3 < 0, 0, ovrd_msnalt3)
            
            anaGrid.TextMatrix(3 + i, 36) = rsB("alt4_prod_nomach")
            anaGrid.TextMatrix(3 + i, 37) = IIf(IsNull(rsB("talt4")), "", rsB("talt4"))
            anaGrid.TextMatrix(3 + i, 38) = IIf(msnalt4 < 0, 0, msnalt4)
            anaGrid.TextMatrix(3 + i, 39) = IIf(ovrd_msnalt4 < 0, 0, ovrd_msnalt4)
            
            anaGrid.TextMatrix(3 + i, 40) = rsB("alt5_prod_nomach")
            anaGrid.TextMatrix(3 + i, 41) = IIf(IsNull(rsB("talt5")), "", rsB("talt5"))
            anaGrid.TextMatrix(3 + i, 42) = IIf(msnalt5 < 0, 0, msnalt5)
            anaGrid.TextMatrix(3 + i, 43) = IIf(ovrd_msnalt5 < 0, 0, ovrd_msnalt5)
            
            anaGrid.TextMatrix(3 + i, 44) = rsB("alt6_prod_nomach")
            anaGrid.TextMatrix(3 + i, 45) = IIf(IsNull(rsB("talt6")), "", rsB("talt6"))
            anaGrid.TextMatrix(3 + i, 46) = IIf(msnalt6 < 0, 0, msnalt6)
            anaGrid.TextMatrix(3 + i, 47) = IIf(ovrd_msnalt6 < 0, 0, ovrd_msnalt6)
            anaGrid.TextMatrix(3 + i, 48) = FormatNumber((rsB("prod_plan_1") / c_cap_p_day) / hkw1 * 100)
'            prosesSisa 3 + i, presentMesinUse, rsB("prod_plan_1") / c_cap_p_day, hkw1, c_cap_p_day
            
            i = i + 1
            ar_mesin_present(i) = rsB("prod_plan_1") / c_cap_p_day 'presentMesinUse
            ar_mesin(i) = rsB("prod_nomach")
            ar_mesin_alt1(i) = IIf(IsNull(rsB("alt1_prod_nomach")), "nm", rsB("alt1_prod_nomach"))
            rsB.MoveNext
        Wend
        
'        For j = 1 To rsB.RecordCount
'            MsgBox ar_mesin_present(j) & " " & ar_mesin(j)
'        Next
        
        '#re think main machine
'        For j = 3 To anaGrid.rows - 1
'            If ar_mesin(j - 2) = anaGrid.TextMatrix(j, 20) Then
'                totalp_mesin = 0
'                For x = 1 To j - 2
'                    If ar_mesin(x) = ar_mesin(j - 2) Then
'                        totalp_mesin = totalp_mesin + Val(ar_mesin_present(x))
'                    End If
'                Next
'                totalp_mesin = totalp_mesin / hkw1 * 100
'                anaGrid.TextMatrix(j, 22) = FormatNumber(totalp_mesin, 2)
'            End If
'        Next
'
'        For j = 3 To anaGrid.rows - 1
'            If ar_mesin(j - 2) = anaGrid.TextMatrix(j, 20) And (Val(ar_mesin_present(j - 2))) > 24 Then
'                anaGrid.TextMatrix(j, 22) = FormatNumber(100, 2)
'                For k = j + 1 To anaGrid.rows - 1
'                    If anaGrid.TextMatrix(k, 20) = ar_mesin(j - 2) Then
'                        anaGrid.TextMatrix(k, 22) = FormatNumber(0, 2)
'                    End If
'                Next
'            End If
'        Next
        
        '#falt1
'        For j = 3 To anaGrid.rows - 1
'            If ar_mesin(j - 2) = anaGrid.TextMatrix(j, 20) And (Val(ar_mesin_present(j - 2))) > 24 Then
'                If (Val(ar_mesin_present(j - 2)) - 24) >= 24 Then
'                    anaGrid.TextMatrix(j, 26) = 100
'                Else
'                    anaGrid.TextMatrix(j, 26) = FormatNumber((Val(ar_mesin_present(j - 2)) - 24) / hkw1 * 100, 2)
'                    'MsgBox anaGrid.TextMatrix(j, 26)
'                End If
'                anaGrid.TextMatrix(j, 23) = FormatNumber(Val(ar_mesin_present(j - 2)) - 24, 2)
'            Else
'                If anaGrid.TextMatrix(j, 22) = 0 Then
'                    anaGrid.TextMatrix(j, 26) = FormatNumber(Val(ar_mesin_present(j - 2)) / hkw1 * 100, 2)
'                    anaGrid.TextMatrix(j, 23) = FormatNumber(Val(ar_mesin_present(j - 2)), 2)
'                End If
'            End If
'        Next
        
'        For j = 3 To anaGrid.rows - 1
'            If ar_mesin_alt1(j - 2) = anaGrid.TextMatrix(j, 24) Then
'                totalp_mesin = 0
'                For x = 1 To j - 2
'                    If ar_mesin_alt1(x) = ar_mesin_alt1(j - 2) Then
'                        If x = 1 Then
'                            MsgBox ar_mesin_alt1(x)
''                            For k = 3 To anaGrid.rows - 1
''                                MsgBox "if " & ar_mesin_alt1(x) & "=" & anaGrid.TextMatrix(k, 20)
''                                If ar_mesin_alt1(x) = anaGrid.TextMatrix(k, 20) Then
''                                    MsgBox anaGrid.TextMatrix(k, 20)
''                                End If
''                            Next
'                        End If
'                        totalp_mesin = totalp_mesin + Val(anaGrid.TextMatrix(x + 2, 23))
'                    End If
'                Next
'                totalp_mesin = totalp_mesin / hkw1 * 100
'                anaGrid.TextMatrix(j, 26) = FormatNumber(totalp_mesin, 2)
'            End If
'        Next
'        For j = 3 To anaGrid.rows - 1
'            For x = 3 To anaGrid.rows - 1
'                If anaGrid.TextMatrix(j, 20) = anaGrid.TextMatrix(x, 24) Then
'                    MsgBox
'                End If
'            Next
'        Next
'        For j = 3 To anaGrid.rows - 1
'            For k = 1 To anaGrid.rows - 1
'                If anaGrid.TextMatrix(j, 23) = anaGrid.TextMatrix(k, 20) Then
'                    For t = 1 To anaGrid.rows - 1
'                        If anaGrid.TextMatrix(t, 20) = anaGrid.TextMatrix(j, 23) Then
'                            MsgBox anaGrid.TextMatrix(t, 22)
'                        End If
'                    Next
'                End If
'            Next
'        Next
    Else
        anaGrid.rows = 3
    End If
    gridFormatNum
    anaGrid.Refresh
End Sub

Private Sub CmbDocument_DropDown()
    If IsNumeric(txtRevision) = False Then txtRevision.SetFocus: Exit Sub
    qry = "select distinct on (ltpp_doc) ltpp_doc from ltpp_generate where rev=" & txtRevision & " and period='" & Format(DTPicker1.value, "yyyyMM") & "'"
    Set rsA = Con.Execute(qry)
    CmbDocument.Clear
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            CmbDocument.AddItem rsA(0)
            rsA.MoveNext
        Wend
    End If
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

Private Sub Form_Activate()
    FocusTab Me
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    AddTab Me
    Call BukaKoneksi
    Call activeTheme(skinFD, Me)
    Call settingFG
    Me.Height = 7755
    Me.Width = 13545
Exit Sub
errLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error Load: " & Err.Number
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls
    CmbDocument.Left = txtRevision.Left
    CmbDocument.Top = SkinLabel3.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        DelTab Me
    End If
End Sub

