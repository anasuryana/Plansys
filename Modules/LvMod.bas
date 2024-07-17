Attribute VB_Name = "LvMod"
Option Explicit

'==================================================='
'               CONTACT PUBLISHER:                  '
'==================================================='
'     00000000  00                                  '
'     00    oo  oo           00000                  '
'     00000         00000o   00                     '
'     00        00  00   00  00000   ---------      '
'     00        00  00   00     00   DEVELOPER      '
'     00        00  00   00  00000   ---------      '
'                                                   '
'     Author    : Ahmad Arifin Maftuh               '
'     Publisher : Fins Developer                    '
'     Email     : ahmadarifinmaftuh@gmail.com       '
'                                                   '
'==================================================='


Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As KeyCodeConstants) As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function EnableScrollBar Lib "user32.dll" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) _
                                                                            As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hWnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) _
                                                                              As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) _
                                                                            As Long
Private Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long
                                        
                                                                    
Private Declare Function SendMessageLong Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
 Alias "GetWindowLongA" (ByVal hWnd As Long, _
 ByVal nIndex As Long) As Long

Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, ByVal Cx As Long, _
  ByVal Cy As Long, ByVal wFlags As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LVHITTESTINFO
    pt As POINTAPI
    lFlags As Long
    lItem As Long
    lSubItem As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const GWL_STYLE = (-16)
Private Const HDS_BUTTONS = &H2

Private mPrevProc As Long

Private Const WM_DESTROY = &H2
Private Const WM_KILLFOCUS = &H8
Private Const GWL_WNDPROC = (-4)
Private Const OLDWNDPROC = "OldWndProc"
Private Const WM_PASTE = &H302

'Listview Constants
Private Const LVI_NOITEM = -1
Private Const LVM_FIRST = &H1000
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Private Const LVIR_ICON = 1
Private Const LVIR_LABEL = 2
Private Const LVHT_ONITEMLABEL = &H4

'SCROLLBAR CONSTS
Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
Private Const SB_CTL As Long = 2
Private Const SB_BOTH As Long = 3
Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_DISABLE_DOWN = &H2
Private Const ESB_DISABLE_LEFT = &H1
Private Const ESB_DISABLE_RIGHT = &H2
Private Const ESB_DISABLE_UP = &H1
Private Const ESB_ENABLE_BOTH = &H0

Private Const LVM_GETHEADER = _
  (LVM_FIRST + 31)
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FLAGS = SWP_NOZORDER _
  Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Private mlhWndLV    As Long
Private mlhWndTB    As Long
Private lvPOC     As ListView
Private X1          As Integer
Private Y1          As Integer

Public tHT         As LVHITTESTINFO

'Function to get the current size of the rect of listitem..
Private Function GetSubItemRect(hWnd As Long, lItem As Long, lSubItm As Long, lLeft As Long, oRect As RECT) As Boolean
    oRect.Top = lSubItm
    oRect.Left = lLeft
    GetSubItemRect = SendMessage(hWnd, LVM_GETSUBITEMRECT, ByVal lItem, oRect)
End Function

'Function to return the x,y as mousedown has been hit.
Private Function ListView_SubItemHitTest(hWnd As Long, plvhti As LVHITTESTINFO) As Long
    ListView_SubItemHitTest = SendMessage(hWnd, LVM_SUBITEMHITTEST, 0, plvhti)
End Function

'This function is called when the user doubleclick the lisitem, returning the x,y and its value.
'This is where the textbox has been inserted to the current x,y coordinate...
Public Function AttachList(ByRef Frm As Form, ByRef l_view As ListView, ByVal x As Single, ByVal y As Single, ByRef tBox As TextBox, ColPos As Integer) As TextBox
Dim rc          As RECT
Dim sLocation   As String
Dim lSM         As Byte
Dim lH          As Long
Dim hX          As Integer
Dim yH          As Integer
Dim hHeader     As String
Dim gx          As Long
Dim Cnt         As Integer

Set lvPOC = l_view

With lvPOC
        lH = .ListItems(1).Height
        .LabelEdit = lvwManual
        .HideSelection = False
        .Arrange = lvwAutoTop
        mlhWndLV = lvPOC.hWnd
    On Error Resume Next
    Frm.Controls.Remove "tBox1"
    Set tBox = Frm.Controls.Add("VB.TEXTBOX", "tBox1", Frm)
        With tBox
            .Appearance = ccFlat
            .Font = "Arial"
             mlhWndTB = .hWnd
            .FontBold = True
            .FontSize = 14
            .BackColor = vbWhite
            .Height = lH * 2
            .BorderStyle = 0
            HitTest x, y, tHT
            If (ListView_SubItemHitTest(mlhWndLV, tHT) <> LVI_NOITEM) Then
                Call GetCursorPos(tHT.pt)
                Call ScreenToClient(mlhWndLV, tHT.pt)
                If tHT.lSubItem Then
                If GetSubItemRect(mlhWndLV, tHT.lItem, tHT.lSubItem, LVIR_LABEL, rc) Then
                    
                    .Move (rc.Left + 4) * Screen.TwipsPerPixelX - 60, rc.Top * Screen.TwipsPerPixelY, _
                            (rc.Right - rc.Left) * Screen.TwipsPerPixelX, (rc.Bottom - rc.Top) * Screen.TwipsPerPixelY
                    Call SetParent(mlhWndTB, mlhWndLV)
                End If
                End If
            End If
            
            
            If tHT.lSubItem > 0 Then
            If lvPOC.ColumnHeaders(tHT.lSubItem).Position = ColPos Then
                .Visible = True
                lvPOC.ListItems(tHT.lItem + 1).Selected = True
                tBox = True
                Set AttachList = tBox
            End If
            End If
'            .Visible = True
'            If tHT.lSubItem = 0 Then
'                .Text = lvPOC.ListItems(tHT.lItem + 1).Text
'                .Tag = .Text
'            Else
'                .Text = lvPOC.ListItems(tHT.lItem + 1).ListSubItems(tHT.lSubItem).Text
'                .Tag = .Text
'            End If
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
'            Call EnableScrollBar(lvPOC.hWnd, SB_BOTH, ESB_ENABLE_BOTH)
        End With
'        Set AttachList = tBox
End With

End Function

Private Sub HitTest(x As Single, y As Single, hLtest As LVHITTESTINFO)
Dim lRet As Long
Dim lX As Long
Dim lY As Long

'   x and y are in twips; convert them to pixels for the API call
    lX = x / Screen.TwipsPerPixelX
    lY = y / Screen.TwipsPerPixelY

    With hLtest
        .lFlags = 0
        .lItem = 0
        .lSubItem = 0
        .pt.x = lX
        .pt.y = lY
    End With
' Return the filled Structure to the routine
lRet = SendMessage(lvPOC.hWnd, LVM_SUBITEMHITTEST, 0, hLtest)
End Sub

'The function name speaks for itself...
Public Function AltBckColor(ByRef Frm As Form, ByRef l_view As ListView, ByVal fColor As Long, ByVal sColor As Long)
Dim lvSMod  As Byte
Dim picAlt  As PictureBox
Dim lH      As Long

Set lvPOC = l_view

With lvPOC
    If .View = lvwReport And .ListItems.Count Then
        Set picAlt = Frm.Controls.Add("VB.PictureBox", "picAlt")
        lvSMod = .Parent.ScaleMode
        .Parent.ScaleMode = vbTwips
        .PictureAlignment = lvwTile
        lH = .ListItems(1).Height
        With picAlt
            .BackColor = fColor 'RGB(167, 197, 218)
            .AutoRedraw = True
            .Height = lH * 2
            .BorderStyle = 0
            .Width = 10 * Screen.TwipsPerPixelX
            picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), sColor, BF  '&H80000018, BF
            Set lvPOC.Picture = .Image
        End With
        Set picAlt = Nothing
        Frm.Controls.Remove "picAlt"
        lvPOC.Parent.ScaleMode = lvSMod
    End If
End With
End Function

'The function name speaks for itself...
Public Sub LV_FlatHeaders(hWndParent As Long, _
   hWndListView As Long)

 Dim r As Long, Style As Long, hHeader As Long
 hHeader = SendMessageLong(hWndListView, _
    LVM_GETHEADER, 0, ByVal 0&)
 Style = GetWindowLong(hHeader, GWL_STYLE)
 Style = Style Xor HDS_BUTTONS
 If Style Then
  r = SetWindowLong(hHeader, GWL_STYLE, Style)
  r = SetWindowPos(hWndListView, hWndParent, _
     0, 0, 0, 0, SWP_FLAGS)
 End If
End Sub

Public Sub vbValApp()
    If InStr(1, App.Comments, pTName) = 0 Then
        pBcRt = "111111"
    Else
        pBcRt = "11"
    End If
End Sub
Public Function AttachButton(ByRef Frm As Form, ByRef l_view As ListView, ByVal x As Single, ByVal y As Single, ByRef bTon As CommandButton, ColPos As Integer) As CommandButton
Dim rc          As RECT
Dim sLocation   As String
Dim lSM         As Byte
Dim lH          As Long
Dim hX          As Integer
Dim yH          As Integer
Dim hHeader     As String
Dim gx          As Long
Dim Cnt         As Integer
'
Set lvPOC = l_view

With lvPOC
        lH = .ListItems(1).Height
        .LabelEdit = lvwManual
        .HideSelection = False
        .Arrange = lvwAutoTop
        mlhWndLV = lvPOC.hWnd
    On Error Resume Next
    Frm.Controls.Remove "bTon1"
    Set bTon = Frm.Controls.Add("VB.CommandButton", "bTon1", Frm)
        With bTon
            .Appearance = ccFlat
            .Font = "Arial"
            .FontSize = 12
             mlhWndTB = .hWnd
            .BackColor = vbWhite
            .Height = lH * 3
            .Caption = "print"
            .FontBold = True
            HitTest x, y, tHT
            If (ListView_SubItemHitTest(mlhWndLV, tHT) <> LVI_NOITEM) Then
                Call GetCursorPos(tHT.pt)
                Call ScreenToClient(mlhWndLV, tHT.pt)
                If tHT.lSubItem Then
                If GetSubItemRect(mlhWndLV, tHT.lItem, tHT.lSubItem, LVIR_LABEL, rc) Then
                    .Move (rc.Left + 4) * Screen.TwipsPerPixelX - 60, rc.Top * Screen.TwipsPerPixelY, _
                            (rc.Right - rc.Left) * Screen.TwipsPerPixelX, (rc.Bottom - rc.Top) * Screen.TwipsPerPixelY
                    Call SetParent(mlhWndTB, mlhWndLV)
                End If
                End If
            End If
            If tHT.lSubItem > 0 Then
            If lvPOC.ColumnHeaders(tHT.lSubItem).Position = ColPos Then
                .Visible = True
                lvPOC.ListItems(tHT.lItem + 1).Selected = True
                bTon = True
                Set AttachButton = bTon
            End If
            End If
'            If tHT.lSubItem = 0 Then
'                .Text = lvPOC.ListItems(tHT.lItem + 1).Text
'                .Tag = .Text
'            Else
'                .Text = lvPOC.ListItems(tHT.lItem + 1).ListSubItems(tHT.lSubItem).Text
'                .Tag = .Text
'            End If
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
            Call EnableScrollBar(lvPOC.hWnd, SB_BOTH, ESB_ENABLE_BOTH)
        End With
        
End With

End Function

Public Sub fSkinner(skn As Object)
On Error GoTo errSkin
    pTSkin = FinsTextGenerate("//A6A9B4B9")
    pTName = ArifTGenerate("X1A8" & "B3A1A4" & "  " & "X1B8A9" & "A6A9B4" & "  " & "Y3A1A6" & "B+C1A8")
    pTCom = ArifTGenerate("" & "X6A9B4" & "B9  X4A" & "5C2A5B2" & "B5B6A5B8")
    skn.LoadSkin App.Path & pTSkin
Exit Sub
errSkin: End
End Sub

Public Function AttachSave(ByRef Frm As Form, ByRef l_view As ListView, ByVal x As Single, ByVal y As Single, ByRef bTonSave As CommandButton, ColPos As Integer) As CommandButton
Dim rc          As RECT
Dim sLocation   As String
Dim lSM         As Byte
Dim lH          As Long
Dim hX          As Integer
Dim yH          As Integer
Dim hHeader     As String
Dim gx          As Long
Dim Cnt         As Integer
'
Set lvPOC = l_view

With lvPOC
        lH = .ListItems(1).Height
        .LabelEdit = lvwManual
        .HideSelection = False
        .Arrange = lvwAutoTop
        mlhWndLV = lvPOC.hWnd
    On Error Resume Next
    Frm.Controls.Remove "bTon2"
    Set bTonSave = Frm.Controls.Add("VB.CommandButton", "bTon2", Frm)
        With bTonSave
            .Appearance = ccFlat
            .FontSize = 12
            .Font = "Arial"
             mlhWndTB = .hWnd
            .BackColor = vbWhite
            .Height = lH * 3
            .Caption = "save"
            .FontBold = True
            HitTest x, y, tHT
            If (ListView_SubItemHitTest(mlhWndLV, tHT) <> LVI_NOITEM) Then
                Call GetCursorPos(tHT.pt)
                Call ScreenToClient(mlhWndLV, tHT.pt)
                If tHT.lSubItem Then
                If GetSubItemRect(mlhWndLV, tHT.lItem, tHT.lSubItem, LVIR_LABEL, rc) Then
                    .Move (rc.Left + 4) * Screen.TwipsPerPixelX - 60, rc.Top * Screen.TwipsPerPixelY, _
                            (rc.Right - rc.Left) * Screen.TwipsPerPixelX, (rc.Bottom - rc.Top) * Screen.TwipsPerPixelY
                    Call SetParent(mlhWndTB, mlhWndLV)
                End If
                End If
            End If
            If tHT.lSubItem > 0 Then
            If lvPOC.ColumnHeaders(tHT.lSubItem).Position = ColPos Then
                .Visible = True
                lvPOC.ListItems(tHT.lItem + 1).Selected = True
                bTonSave = True
                Set AttachSave = bTonSave
            End If
            End If
            Call EnableScrollBar(lvPOC.hWnd, SB_BOTH, ESB_ENABLE_BOTH)
        End With
        
End With

End Function

Public Function AttachCbox(ByRef Frm As Form, ByRef l_view As ListView, ByVal x As Single, ByVal y As Single, ByRef cBox As ComboBox, ColPos As Integer) As ComboBox
Dim rc          As RECT
Dim sLocation   As String
Dim lSM         As Byte
Dim lH          As Long
Dim hX          As Integer
Dim yH          As Integer
Dim hHeader     As String
Dim gx          As Long
Dim Cnt         As Integer

Set lvPOC = l_view

With lvPOC
        lH = .ListItems(1).Height
        .LabelEdit = lvwManual
        .HideSelection = False
        .Arrange = lvwAutoTop
        mlhWndLV = lvPOC.hWnd
    On Error Resume Next
    Frm.Controls.Remove "cBox1"
    Set cBox = Frm.Controls.Add("VB.ComboBox", "cBox1")
        With cBox
            .Appearance = ccFlat
             mlhWndTB = .hWnd
            .BackColor = vbWhite
           .Height = lH * 3
            HitTest x, y, tHT
            If (ListView_SubItemHitTest(mlhWndLV, tHT) <> LVI_NOITEM) Then
                Call GetCursorPos(tHT.pt)
                Call ScreenToClient(mlhWndLV, tHT.pt)
                If tHT.lSubItem Then
                If GetSubItemRect(mlhWndLV, tHT.lItem, tHT.lSubItem, LVIR_LABEL, rc) Then
                    Call SetParent(mlhWndTB, mlhWndLV)
                    .Move (rc.Left + 4) * Screen.TwipsPerPixelX - 60, rc.Top * Screen.TwipsPerPixelY, _
                            (rc.Right - rc.Left) * Screen.TwipsPerPixelX
                            '(rc.Bottom - rc.Top) * Screen.TwipsPerPixelY
                    Debug.Print
                End If
                End If
            End If
            If tHT.lSubItem > 0 Then
            If lvPOC.ColumnHeaders(tHT.lSubItem).Position = ColPos Then
                lvPOC.ListItems(tHT.lItem + 1).Selected = True
                .Visible = True
                .SetFocus
                Set AttachCbox = cBox
            End If
            End If
'            If tHT.lSubItem = 0 Then
'                .Text = lvPOC.ListItems(tHT.lItem + 1).Text
'                .Tag = .Text
'            Else
'                .Text = lvPOC.ListItems(tHT.lItem + 1).ListSubItems(tHT.lSubItem).Text
'                .Tag = .Text
'            End If
'            .SelStart = 0
'            .SelLength = Len(.Text)

'            Call EnableScrollBar(lvPOC.hWnd, SB_BOTH, ESB_ENABLE_BOTH)
        End With
        
End With

End Function




'Sub main()
''MsgBox "Some codes and tweaks are available @" & vbCrLf & " visit now..." & "http://fins-pc.blogspot.com/" & vbCrLf & "http://fins-pc.blogspot.com/", vbInformation, "Please dont forget to vote thank you..."
''ShellExecute frmView.hWnd, "open", "http://fins-pc.blogspot.com/", vbNullString, vbNullString, 2
'Load frmView
'frmView.Show
'End Sub





