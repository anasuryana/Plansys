Attribute VB_Name = "ListViewSroll"
Option Explicit

Private Type LV_ITEM
   Mask As Long
   iItem As Long
   iSubItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   lParam As Long
   iIndent As Long
End Type

Private Const LVM_FIRST              As Long = &H1000
Private Const LVM_GETTOPINDEX        As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE    As Long = (LVM_FIRST + 40)
Private Const LVM_SETITEMSTATE       As Long = (LVM_FIRST + 43)
Private Const LVIS_FOCUSED           As Long = &H1
Private Const LVIS_SELECTED          As Long = &H2
Private Const LVIF_STATE             As Long = &H8

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function SetItemFocusA(ByRef ctlListview As MSComctlLib.ListView, ByVal iIndex As Long, Optional ByVal iVisibleIndex = 3) As Boolean
On Error GoTo Hell

Dim LV As LV_ITEM
Dim lvItemsPerPage As Long
Dim lvNeededItems As Long
Dim lvCurrentTopIndex As Long

    With ctlListview
        ' Since this is a multi-select list, we want to unselect all items before selecting the current track.
        With LV
            .Mask = LVIF_STATE
            .State = False
            .stateMask = LVIS_SELECTED
        End With
        Call SendMessage(.hWnd, LVM_SETITEMSTATE, -1, LV)  ' Poof
        
        ' Select and set the focus rectangle on the item.
        With LV
            .Mask = LVIF_STATE
            .State = True
            .stateMask = LVIS_SELECTED Or LVIS_FOCUSED
        End With
        Call SendMessage(.hWnd, LVM_SETITEMSTATE, iIndex - 1, LV)  ' Listview index is 0-based in the API world
        
        ' Determine if desired index + number of items in view will exceed total items in the control
        lvCurrentTopIndex = SendMessage(.hWnd, LVM_GETTOPINDEX, 0&, ByVal 0&)
        lvItemsPerPage = SendMessage(.hWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
        
        ' Do we even need to scroll? Not if the selected track is already in view
        If (lvCurrentTopIndex >= iIndex) Or (iIndex > lvCurrentTopIndex + lvItemsPerPage) Then
        
            ' Is 'x' above or below target index?
            If lvCurrentTopIndex >= iIndex Then  ' Going UP
                If iIndex > iVisibleIndex Then
                    .ListItems((iIndex - iVisibleIndex + 1)).EnsureVisible ' Drops the highlighted item down a few so it's not hidden
                                                            ' behind the Column header.
                Else
                    .ListItems((iIndex)).EnsureVisible
                End If
            
            Else ' Going DOWN
                ' Are there sufficient items to set to the topindex
                If (iIndex + lvItemsPerPage) > .ListItems.Count Then
               
                   ' Can't be set to the top as the control has insufficient
                   ' items, so just scroll to the end of listview
                   .ListItems(.ListItems.Count).EnsureVisible
                   
                Else
                
                  ' It is below, and since a listview always moves the item just into view,
                  ' have it instead move to the top by faking item we want to 'EnsureVisible'
                  ' the item lvItemsPerPage -1(or -3) below the actual index of interest.
                    If iIndex > iVisibleIndex Then
                        .ListItems((iIndex + lvItemsPerPage) - iVisibleIndex).EnsureVisible
                    Else
                        .ListItems((iIndex + lvItemsPerPage) - 1).EnsureVisible
                    End If
                End If
            End If
        End If
    End With

    SetItemFocusA = True

Hell:
End Function

