VERSION 5.00
Begin VB.Form sort 
   BackColor       =   &H80000012&
   Caption         =   "Sort"
   ClientHeight    =   750
   ClientLeft      =   1860
   ClientTop       =   2100
   ClientWidth     =   4125
   Icon            =   "Sort.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   750
   ScaleWidth      =   4125
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   45
      TabIndex        =   2
      Text            =   "First Name"
      Top             =   210
      Width           =   2670
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   2925
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Accept"
      Height          =   255
      Left            =   2925
      TabIndex        =   0
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
 Combo2.Clear
  Dim p As Integer
        If Combo1.Text = "First Name" Then
            While p <> smax
                Combo2.AddItem sair(p).first
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Second Name" Then
            While p <> smax
                Combo2.AddItem sair(p).second
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Relation" Then
            
            While p <> smax
                Combo2.AddItem sair(p).relation
                p = p + 1
            Wend
       ElseIf Combo1.Text = "Place" Then
            While p <> smax
                Combo2.AddItem sair(p).place
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Town" Then
            While p <> smax
                Combo2.AddItem sair(p).town
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Pin" Then
            
            While p <> smax
                Combo2.AddItem sair(p).pin
                p = p + 1
            Wend
        ElseIf Combo1.Text = "District" Then
            While p <> smax
                Combo2.AddItem sair(p).district
                p = p + 1
            Wend
        ElseIf Combo1.Text = "State" Then
            While p <> smax
                Combo2.AddItem sair(p).state
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Country" Then
            
            While p <> smax
                Combo2.AddItem sair(p).country
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Tel.No" Then
            While p <> smax
                Combo2.AddItem sair(p).telno
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Fax" Then
            While p <> smax
                Combo2.AddItem sair(p).fax
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Email" Then
            
            While p <> smax
                Combo2.AddItem sair(p).email
                p = p + 1
            Wend
        ElseIf Combo1.Text = "Email" Then
            
            While p <> smax
                Combo2.AddItem sair(p).Web
                p = p + 1
            Wend
        End If
End Sub

Private Sub Command1_Click()
    Unload Me
    
End Sub

Private Sub Command2_Click()
   If smax < 2 Then
        Exit Sub
    End If
    Dim ki As Integer
    Dim li As Integer
    If Combo1.Text = "First Name" Then
        While ki <> smax - 1
            li = ki
            While li <> smax
                If StrComp(sair(ki).first, sair(li).first) > 0 Then
                    Call aniswap(ki, li)
                End If
                li = li + 1
            Wend
            ki = ki + 1
        Wend
    End If
    Unload Me
End Sub


