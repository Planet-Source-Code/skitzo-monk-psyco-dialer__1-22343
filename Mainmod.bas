Attribute VB_Name = "Module2"
' Programmed and designed by
'  Nagalla Anil Choudary
'  D.K.Pallem
'  Bapatla-522 101
'  A.P, India
'  you can redistribute reproduce the source code as u like
'  but mail your comments/updations to anilfriend@hotmail.com
'

Type node
     first As String
     second As String
     relation As String
     place As String
     town As String
     pin As String
     district As String
     state As String
     country As String
     telno As String
     fax As String
     email As String
     web As String
     notes As String
     photo As String
End Type
Public wait As Integer
Public ppath As String
Public smax As Integer
Public sair() As node
Function aniswap(s1 As Integer, s2 As Integer)
    Dim temp As node
    temp = sair(s1)
    sair(s1) = sair(s2)
    sair(s2) = temp
End Function


