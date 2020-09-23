VERSION 5.00
Begin VB.Form find 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1260
   ClientLeft      =   1425
   ClientTop       =   2040
   ClientWidth     =   4125
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1260
   ScaleWidth      =   4125
   Begin VB.CommandButton Command2 
      Caption         =   "&Accept"
      Height          =   255
      Left            =   2925
      TabIndex        =   3
      Top             =   930
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   210
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   2925
      TabIndex        =   1
      Top             =   630
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Text            =   "First Name"
      Top             =   210
      Width           =   1635
   End
End
Attribute VB_Name = "find"
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
On Error Resume Next
        Dim x, flag
        x = 0
        While x <> smax
            flag = StrComp(Combo2.Text, Combo2.List(x))
            If flag = 0 Then
                GoTo anil
            Else
                flag = -1
            End If
            x = x + 1
        Wend
anil:
        If flag = -1 Then
            MsgBox ("SORRY NOT MATCHED")
            Exit Sub
        End If
    On Error Resume Next
    address.Text1.Text = sair(x).first
    address.Text2.Text = sair(x).second

    address.Text4.Text = sair(x).town
    address.Text8.Text = sair(x).country
    address.Text9.Text = sair(x).telno
    address.Text10.Text = sair(x).fax
    address.Text11.Text = sair(x).email
    address.Text12.Text = sair(x).Web
    address.Text13.Text = sair(x).notes
    Unload Me
End Sub

Private Sub Form_Load()
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

