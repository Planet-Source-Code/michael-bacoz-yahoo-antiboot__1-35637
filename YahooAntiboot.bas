Attribute VB_Name = "YahooAntiboot"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Const WM_CLOSE = &H10

Public Function Yahoo_Antiboot()
Dim imclass As Long, richedit As Long
Dim X As Integer

Dim atlfba(0 To 5) As Long


imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "RICHEDIT", vbNullString)
richedit = FindWindowEx(imclass, richedit, "RICHEDIT", vbNullString)

atlfba(0) = FindWindowEx(imclass, 0&, "atl:004eeb20", vbNullString)
atlfba(1) = FindWindowEx(imclass, 0&, "atl:004ebb50", vbNullString)
atlfba(2) = FindWindowEx(imclass, 0&, "atl:004eeb68", vbNullString)
atlfba(3) = FindWindowEx(imclass, 0&, "atl:004efb68", vbNullString)
atlfba(4) = FindWindowEx(imclass, 0&, "atl:004f0b88", vbNullString)
atlfba(5) = FindWindowEx(imclass, 0&, "atl:004f0ba8", vbNullString)

Call SendMessageLong(richedit, WM_CLOSE, 0&, 0&)

For X = 0 To 5
   Call SendMessageLong(atlfba(X), WM_CLOSE, 0&, 0&)
Next X

Current = Timer
Do While Timer - Current < Val(0.2)
    DoEvents
Loop
End Function

