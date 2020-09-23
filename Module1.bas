Attribute VB_Name = "odule1"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global X As Integer
Global NEWMESSAGES As Integer
Global T As Integer
Option Explicit

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Function OnTop(Frm As Form)
SetWindowPos Frm.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Function
Sub Wait(Interval As Long)
Dim time
time = Timer
Do While Timer - time < Val(Interval)
DoEvents
Loop
End Sub


