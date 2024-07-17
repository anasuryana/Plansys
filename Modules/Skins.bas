Attribute VB_Name = "Skins"
Option Explicit

Public Sub activeTheme(skn As Object, Frm As Form)
    fSkinner skn
    skn.ApplySkin Frm.hWnd
End Sub

