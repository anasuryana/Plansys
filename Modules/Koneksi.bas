Attribute VB_Name = "Koneksi"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public stsFormPor As Boolean
Public GetForm As String
Public pUserName As String
Public pUserId As String
Public pTemplateLTPP As String
Public pTemplateMPP As String
Public pTSkin As String
Public pTName As String
Public pTCom    As String
Public Const KARAKTERBAHAYA As String = "'`"""
Public Const HURUFCEGAH     As String = "!@#$%^&*()_+-=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Con As New ADODB.Connection
Public RsDB As New ADODB.Recordset
Public RsBantu As New ADODB.Recordset
Public RsTemp As New ADODB.Recordset
Public RsGet As New ADODB.Recordset
Public RsDoc As New ADODB.Recordset

Private Function FileName() As String
    FileName = App.Path & "\setting.conf"
End Function
 
Public Function GetINI(ByVal Section As String, ByVal Key As String, ByVal ValueDefault)
    Dim s As String, l As Long
    s = String(255, 0)
    l = GetPrivateProfileString(Section, Key, Default, s, 255, FileName)
    GetINI = Left(s, l)
    vbValApp
    vbValCom
End Function
 
Public Function SaveINI(ByVal Section As String, ByVal Key As String, ByVal SaveValue As String)
    If SaveValue = "" Then SaveValue = vbNullChar
    SaveINI = WritePrivateProfileString(Section, Key, SaveValue, FileName)
End Function

Public Sub myTemplates()
    pTemplateLTPP = App.Path & "\Templates\LTPP.fdxl"
    pTemplateMPP = App.Path & "\Templates\MPP.fdxl"
End Sub

Public Sub BukaKoneksi()
Set Con = New ADODB.Connection
'Con.CommandTimeout = 1000
Con.Open GetINI("SETTING", "odbc", vbNullString)
Con.CursorLocation = adUseClient
ConCheck
End Sub

Public Sub selectDB()
    Set RsDB = New ADODB.Recordset
    RsDB.CursorLocation = adUseClient
End Sub

Public Sub selectBantu()
    Set RsBantu = New ADODB.Recordset
    RsBantu.CursorLocation = adUseClient
End Sub

Public Sub DBtemp()
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
End Sub



Public Sub selectGet()
    Set RsGet = New ADODB.Recordset
    RsGet.CursorLocation = adUseClient
End Sub

Public Sub selectDoc()
    Set RsDoc = New ADODB.Recordset
    RsDoc.CursorLocation = adUseClient
End Sub





