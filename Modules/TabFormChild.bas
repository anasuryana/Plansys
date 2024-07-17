Attribute VB_Name = "TabFormChild"
Public IDForm As Integer

Public Sub AddTab(ByVal Frm As Form)
    IDForm = IDForm + 1
    MDI_Parent.tabWindow.Tabs.Add(, , Frm.Caption).Tag = IDForm
    Frm.Tag = IDForm
End Sub

Public Sub FocusTab(ByVal Frm As Form)
    Dim i As Integer
    With MDI_Parent.tabWindow
    
    For i = 1 To .Tabs.Count
        If .Tabs(i).Tag = Frm.Tag Then
            .Tabs(i).Selected = True
            Exit For
        End If
    Next
    
    End With
End Sub

Public Sub DelTab(ByVal Frm As Form)
    Dim i As Integer
    With MDI_Parent.tabWindow
    
    For i = 1 To .Tabs.Count
        If .Tabs(i).Tag = Frm.Tag Then
            .Tabs.Remove (i)
            Exit For
        End If
    Next
    
    End With
End Sub


