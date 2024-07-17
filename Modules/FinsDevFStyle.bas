Attribute VB_Name = "FinsDevFStyle"
 '----------------------------------------------------------------------------
#If Win16 Then
  Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
  End Type
#Else
  Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type
#End If

#If Win16 Then
  Declare Sub GetWindowRect Lib "User" (ByVal hwnd As Integer, lpRect As RECT)
  Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
  Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hdc As _
  Integer) As Integer
  Declare Sub SetBkColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
  Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, _
  ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
  Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
  Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject _
  As Integer) As Integer
  Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
  Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, _
  lpRect As RECT) As Long
  Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
  Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal _
  hdc As Long) As Long
  Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal _
  crColor As Long) As Long
  Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
  ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
  Declare Function SelectObject Lib "User32" (ByVal hdc As Long, ByVal hObject _
  As Long) As Long
  Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

Sub ExplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, Cx%, Cy%
Dim TheScreen As Long
Dim Brush As Long
  GetWindowRect f.hwnd, myRect
  formWidth = (myRect.Right - myRect.Left)
  formHeight = myRect.Bottom - myRect.Top
  TheScreen = GetDC(0)
  Brush = CreateSolidBrush(f.BackColor)
  For i = 1 To Movement
    Cx = formWidth * (i / Movement)
    Cy = formHeight * (i / Movement)
    X = myRect.Left + (formWidth - Cx) / 2
    Y = myRect.Top + (formHeight - Cy) / 2
    Rectangle TheScreen, X, Y, X + Cx, Y + Cy
  Next i
  X = ReleaseDC(0, TheScreen)
  DeleteObject (Brush)
End Sub

Public Sub ImplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, Cx%, Cy%
Dim TheScreen As Long
Dim Brush As Long
  GetWindowRect f.hwnd, myRect
  formWidth = (myRect.Right - myRect.Left)
  formHeight = myRect.Bottom - myRect.Top
  TheScreen = GetDC(0)
  Brush = CreateSolidBrush(f.BackColor)
  For i = Movement To 1 Step -1
    Cx = formWidth * (i / Movement)
    Cy = formHeight * (i / Movement)
    X = myRect.Left + (formWidth - Cx) / 2
    Y = myRect.Top + (formHeight - Cy) / 2
    Rectangle TheScreen, X, Y, X + Cx, Y + Cy
  Next i
  X = ReleaseDC(0, TheScreen)
  DeleteObject (Brush)
End Sub
'----------------------------------------------------------------------------

Private Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512, X As Long

    X = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    r = WriteProfileString("windows", "Device", DeviceLine)
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Public Sub Win95SetDefaultPrinter(namaPrinter As String)
    Dim Handle As Long
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim X As Long
    Dim need As Long
    Dim pi5 As PRINTER_INFO_5
    Dim LastError As Long

    PrinterName = namaPrinter
    If PrinterName = "" Then
        Exit Sub
    End If

    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

    X = OpenPrinter(PrinterName, Handle, pd)
    If X = False Then
        Exit Sub
    End If

    X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ReDim t((need \ 4)) As Long

    X = GetPrinter(Handle, 5, t(0), need, need)
    If X = False Then
        Exit Sub
    End If

    pi5.pPrinterName = PtrCtoVbString(t(0))
    pi5.pPortName = PtrCtoVbString(t(1))
    pi5.Attributes = t(2)
    pi5.DeviceNotSelectedTimeout = t(3)
    pi5.TransmissionRetryTimeout = t(4)

    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

       X = SetPrinter(Handle, 5, pi5, 0)

       If X = False Then
           MsgBox "SetPrinter Failed. Error code: " & Err.LastDllError
           Exit Sub
       Else
           If Printer.DeviceName <> namaPrinter Then
                SelectPrinter (namaPrinter)
           End If
       End If

    ClosePrinter (Handle)
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

        DriverName = Left(Buffer, iDriver - 1)
        
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Public Sub ParseList(lstCtl As Control, ByVal Buffer As String)
    Dim i As Integer
    Dim s As String

    Do
        i = InStr(Buffer, Chr(0))
        If i > 0 Then
            s = Left(Buffer, i - 1)
            If Len(Trim(s)) Then lstCtl.AddItem s
            Buffer = Mid(Buffer, i + 1)
        Else
            If Len(Trim(Buffer)) Then lstCtl.AddItem Buffer
            Buffer = ""
        End If
    Loop While i > 0
End Sub

Public Sub WinNTSetDefaultPrinter(namaPrinter As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim PrinterName As String
    Dim r As Long
        Buffer = Space(1024)
        PrinterName = namaPrinter
        r = GetProfileString("PrinterPorts", PrinterName, "", _
            Buffer, Len(Buffer))

        GetDriverAndPort Buffer, DriverName, PrinterPort

        If DriverName <> "" And PrinterPort <> "" Then
            SetDefaultPrinter namaPrinter, DriverName, PrinterPort
            If Printer.DeviceName <> namaPrinter Then
               SelectPrinter (namaPrinter)
            End If
        End If
End Sub

Public Sub setDataDefaultPrinter(namaPrinter As String)
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        Call Win95SetDefaultPrinter(namaPrinter)
    Else
        Call WinNTSetDefaultPrinter(namaPrinter)
    End If
    
End Sub

