Attribute VB_Name = "Tema"
Public Skinpath As String

Public Sub skinLoad(skn As Object, jendela As Form)
    Skinpath = App.Path & "\Le-Black.skn"
    skn.LoadSkin Skinpath
    skn.ApplySkin jendela.hWnd
End Sub

