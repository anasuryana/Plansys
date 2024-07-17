Attribute VB_Name = "StyleText"
Option Explicit
Public debugitem As String

Public Function FinsTextGenerate(codeAlf As String)
On Error Resume Next
    Dim iAlf As Integer
    Dim tLen As Long
    Dim gText As String
    Dim gAlf As String
    
    tLen = Len(Trim(codeAlf))
    If tLen Mod 2 = 0 Then
        tLen = tLen / 2
        gText = ""
        For iAlf = 1 To tLen
            gAlf = FinsDevAlfGenerate(Mid("//A6A9B4B9", (iAlf * 2) - 1, 2))
            gText = gText & gAlf
        Next
    Else
        gText = ""
    End If
    FinsTextGenerate = gText
End Function

Public Function ArifTGenerate(codeAlf As String)
On Error Resume Next
    Dim iAlf As Integer
    Dim tLen As Long
    Dim gText As String
    Dim gAlf As String
    
    tLen = Len(Trim(codeAlf))
    If tLen Mod 2 = 0 Then
        tLen = tLen / 2
        gText = ""
         For iAlf = 1 To tLen
            gAlf = FinsDevAlfGenerate(Mid(codeAlf, (iAlf * 2) - 1, 2))
            gText = gText & gAlf
        Next
    Else
        gText = ""
    End If
    ArifTGenerate = gText
End Function

Private Function FinsDevAlfGenerate(codeAlf As String)
    Dim tAlf As String
    Select Case codeAlf
        Case "A1"
            tAlf = "a"
        Case "A2"
            tAlf = "b"
        Case "A3"
            tAlf = "c"
        Case "A4"
            tAlf = "d"
        Case "A5"
            tAlf = "e"
        Case "A6"
            tAlf = "f"
        Case "A7"
            tAlf = "g"
        Case "A8"
            tAlf = "h"
        Case "A9"
            tAlf = "i"
        Case "A+"
            tAlf = "j"
        Case "B1"
            tAlf = "k"
        Case "B2"
            tAlf = "l"
        Case "B3"
            tAlf = "m"
        Case "B4"
            tAlf = "n"
        Case "B5"
            tAlf = "o"
        Case "B6"
            tAlf = "p"
        Case "B7"
            tAlf = "q"
        Case "B8"
            tAlf = "r"
        Case "B9"
            tAlf = "s"
        Case "B+"
            tAlf = "t"
        Case "C1"
            tAlf = "u"
        Case "C2"
            tAlf = "v"
        Case "C3"
            tAlf = "w"
        Case "C4"
            tAlf = "x"
        Case "C5"
            tAlf = "y"
        Case "C6"
            tAlf = "z"
        Case "X1"
            tAlf = "A"
        Case "X2"
            tAlf = "B"
        Case "X3"
            tAlf = "C"
        Case "X4"
            tAlf = "D"
        Case "X5"
            tAlf = "E"
        Case "X6"
            tAlf = "F"
        Case "X7"
            tAlf = "G"
        Case "X8"
            tAlf = "H"
        Case "X9"
            tAlf = "I"
        Case "X+"
            tAlf = "J"
        Case "Y1"
            tAlf = "K"
        Case "Y2"
            tAlf = "L"
        Case "Y3"
            tAlf = "M"
        Case "Y4"
            tAlf = "N"
        Case "Y5"
            tAlf = "O"
        Case "Y6"
            tAlf = "P"
        Case "Y7"
            tAlf = "Q"
        Case "Y8"
            tAlf = "R"
        Case "Y9"
            tAlf = "S"
        Case "Y+"
            tAlf = "T"
        Case "Z1"
            tAlf = "U"
        Case "Z2"
            tAlf = "V"
        Case "Z3"
            tAlf = "W"
        Case "Z4"
            tAlf = "X"
        Case "Z5"
            tAlf = "Y"
        Case "Z6"
            tAlf = "Z"
        Case ".."
            tAlf = "."
        Case "//"
            tAlf = "/"
        Case "\\"
            tAlf = "\"
        Case "??"
            tAlf = "?"
        Case "!!"
            tAlf = "!"
        Case "@@"
            tAlf = "@"
        Case "##"
            tAlf = "#"
        Case "$$"
            tAlf = "$"
        Case "%%"
            tAlf = "%"
        Case "^^"
            tAlf = "^"
        Case "**"
            tAlf = "*"
        Case "--"
            tAlf = "-"
        Case "  "
            tAlf = " "
        Case Else
            tAlf = "?"
    End Select
    FinsDevAlfGenerate = tAlf
End Function

'Auto Round
Public Function RoundNumber(ByVal Value As Double, Optional PlacesAfterDecimal As Integer = 0) As Double
  Dim nMultiplier As Long
  nMultiplier = 10 ^ PlacesAfterDecimal
  RoundNumber = Int((Value * nMultiplier) + 0.5) / CDbl(nMultiplier)
End Function

Public Function FilterIn(inText As String) As String
    Dim i As Integer, outText As String
    For i = 1 To Len(inText)
        If Mid(inText, i, 1) <> """" And Mid(inText, i, 1) <> "'" And Mid(inText, i, 1) <> "`" Then
        outText = outText & Mid(inText, i, 1)
        End If
    Next
    FilterIn = outText
End Function

Public Function ceiling(nomor As Double) As Long

    ceiling = -Int(-nomor)
End Function

Public Function FKelipatan(pMPQ As Double, pCapPDay As Double, atasBawah As String) As Long
    Dim bReach As Boolean
    Dim MPQ As Double
    bReach = True
    MPQ = pMPQ
    While bReach
        If MPQ > pCapPDay Then
            If atasBawah = "a" Then
                FKelipatan = MPQ '- pMPQ
            Else
                FKelipatan = MPQ - pMPQ
            End If
            bReach = False
        Else
            If MPQ = pCapPDay Then
                FKelipatan = pCapPDay
                bReach = False
            Else
                FKelipatan = MPQ
            End If
        End If
        MPQ = MPQ + pMPQ
    Wend
End Function
