Attribute VB_Name = "ListViewColor"
Public Sub SetListviewItemColour(lvControl As ListView, intRow As Long, ByVal intCol As Integer, lngColour As Long)

    Dim liItem As ListItem
    Dim liSubItem As ListSubItem
    Dim intIndex As Integer

    On Error GoTo errHand

    Set liItem = lvControl.ListItems(intRow)
    If intCol = 0 Then
        liItem.ForeColor = lngColour
        GoTo CleanUp
    End If

    For intIndex = 1 To lvControl.ColumnHeaders.Count - 1
    If intIndex = intCol Then
        Set liSubItem = liItem.ListSubItems(intIndex)
        liSubItem.ForeColor = lngColour
        GoTo CleanUp
        End If
    Next

CleanUp:
    Set liItem = Nothing
    Set liSubItem = Nothing

    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

