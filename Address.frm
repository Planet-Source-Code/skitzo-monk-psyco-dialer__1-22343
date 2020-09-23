VERSION 5.00
Begin VB.Form address 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Psyco Softwarez - Address Book"
   ClientHeight    =   6150
   ClientLeft      =   1050
   ClientTop       =   1125
   ClientWidth     =   8265
   Icon            =   "Address.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   8265
   Begin VB.CommandButton Command13 
      Caption         =   "Dial"
      Height          =   255
      Left            =   4800
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   4800
      TabIndex        =   29
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1800
      TabIndex        =   28
      Top             =   2880
      Width           =   3345
   End
   Begin VB.CommandButton Command11 
      Caption         =   "new/clear"
      Height          =   285
      Left            =   6480
      TabIndex        =   27
      Top             =   5340
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1050
      Width           =   3765
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   3765
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   3060
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   3075
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   2685
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   3600
      Width           =   3420
   End
   Begin VB.TextBox Text13 
      Height          =   1125
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   3930
      Width           =   5985
   End
   Begin VB.CommandButton Command2 
      Caption         =   "l <"
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Top             =   5340
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "< <"
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Top             =   5340
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   " > >"
      Height          =   285
      Left            =   4440
      TabIndex        =   16
      Top             =   5340
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "> l"
      Height          =   285
      Left            =   5520
      TabIndex        =   17
      Top             =   5340
      Width           =   1050
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Update"
      Height          =   285
      Left            =   6945
      TabIndex        =   13
      Top             =   5640
      Width           =   1005
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Exit"
      Height          =   285
      Left            =   5820
      TabIndex        =   12
      Top             =   5640
      Width           =   1050
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Sort"
      Height          =   285
      Left            =   4680
      TabIndex        =   11
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Find"
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Delete"
      Height          =   285
      Left            =   2355
      TabIndex        =   9
      Top             =   5640
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   285
      Left            =   1260
      TabIndex        =   8
      Top             =   5640
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Address Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   2280
      TabIndex        =   31
      Top             =   360
      Width           =   2460
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   1110
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Second Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   25
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Town"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   24
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   23
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Tel.No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   22
      Top             =   2520
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   3240
      Width           =   465
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   3600
      Width           =   405
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   3960
      Width           =   690
   End
   Begin VB.Image Image4 
      Height          =   1080
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmed and designed by
'  Nagalla Anil Choudary
'  D.K.Pallem
'  Bapatla-522 101
'  A.P, India
'  you can redistribute reproduce the source code as u like
'  but mail your comments/updations to anilfriend@hotmail.com
'

Dim i As Integer
Private Sub img_Click()
browse.Show
img.Visible = False
Image3.Visible = True
End Sub

Private Sub Command1_Click()
   ReDim Preserve sair(smax + 1)
        sair(smax).first = Text1.Text
        sair(smax).second = Text2.Text
    
        sair(smax).town = Text4.Text
        sair(smax).country = Text8.Text
        sair(smax).telno = Text9.Text
        sair(smax).fax = Text10.Text
        sair(smax).email = Text11.Text
        sair(smax).Web = Text12.Text
        sair(smax).notes = Text13.Text
        sair(smax).photo = ppath
        smax = smax + 1
        i = smax - 1
End Sub

Private Sub Command10_Click()
 sair(i).first = Text1.Text
        sair(i).second = Text2.Text


        sair(i).town = Text4.Text
        sair(i).country = Text8.Text
        sair(i).telno = Text9.Text
        sair(i).fax = Text10.Text
        sair(i).email = Text11.Text
        sair(i).Web = Text12.Text
        sair(i).notes = Text13.Text
        sair(i).photo = ppath
        ppath = sair(i).photo
'        If (ppath <> "") Then
'            Image3.Picture = LoadPicture(ppath)
'        Else
'            Image3.Picture = image4.Picture
'        End If

End Sub

Private Sub Command11_Click()
Text1.Text = sair(i).first
        Text1.Text = ""
        Text2.Text = ""
       
      
        Text4.Text = ""
      
  
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
End Sub

Private Sub Command12_Click()
Dim RetVal As Long
RetVal = Shell("start mailto:" + Text11.Text, 0)

Exit Sub
End Sub

Private Sub Command13_Click()
Form1.Text1.Text = address.Text9.Text
address.WindowState = vbMinimized
Form1.Show
End Sub

Private Sub Command2_Click()
  If smax = 0 Then Exit Sub
  On Error Resume Next
      i = 0
      
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).Web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
anil:
End Sub

Private Sub Command3_Click()

If smax = 0 Then Exit Sub
On Error Resume Next
 i = (i - 1 + smax) Mod smax
'        i = i - 1
'        If i < 0 Then
'            i = smax - 1
'        End If
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).Web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
anil:
End Sub


Private Sub Command4_Click()
If smax = 0 Then Exit Sub
On Error Resume Next
       i = (i + 1) Mod smax
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).Web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If

anil:
               'i = (i + 1) Mod smax
End Sub


Private Sub Command5_Click()
 
If smax = 0 Then Exit Sub
On Error Resume Next
   i = smax - 1
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).Web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
        
        
anil:

End Sub


Private Sub Command6_Click()
If smax < 0 Then
        MsgBox ("NOTHING  TO DELETE ")
        Exit Sub
    
    ElseIf smax = 1 Or smax < 1 Then
        i = 0
        Text1.Text = ""
        Text2.Text = ""
        Combo1.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        ppath = ""
       
'        Exit Sub
        smax = 0
        Exit Sub
    End If
    Dim z
    z = i
    While z <> smax - 1
        sair(z).first = sair((z + 1)).first  '*****
        sair(z).second = sair((z + 1)).second
        sair(z).relation = sair((z + 1)).relation
        sair(z).place = sair((z + 1)).place
        sair(z).town = sair((z + 1)).town
        sair(z).pin = sair((z + 1)).pin
        sair(z).district = sair((z + 1)).district
        sair(z).state = sair((z + 1)).state
        sair(z).country = sair((z + 1)).country
        sair(z).telno = sair((z + 1)).telno
        sair(z).fax = sair((z + 1)).fax
        sair(z).email = sair((z + 1)).email
        sair(z).Web = sair((z + 1)).Web
        sair(z).notes = sair((z + 1)).notes
        sair(z).photo = sair((z + 1)).photo
        
        z = z + 1
    Wend
        smax = smax - 1
    i = (i) Mod smax
    
    Text1.Text = sair(i).first
    Text2.Text = sair(i).second
    Combo1.Text = sair(i).relation
    Text3.Text = sair(i).place
    Text4.Text = sair(i).town
    Text5.Text = sair(i).pin
    Text6.Text = sair(i).district
    Text7.Text = sair(i).state
    Text8.Text = sair(i).country
    Text9.Text = sair(i).telno
    Text10.Text = sair(i).fax
    Text11.Text = sair(i).email
    Text12.Text = sair(i).Web
    Text13.Text = sair(i).notes
    ppath = sair(i).photo
    If (ppath <> "") Then
    On Error GoTo anil
       Image3.Picture = LoadPicture(ppath)
    Else
       Image3.Picture = Image4.Picture
    End If


    

anil:
End Sub

Private Sub Command7_Click()
 
    find.Show
End Sub

Private Sub Command8_Click()
   sort.Show
End Sub

Private Sub Command9_Click()
  Dim j As Integer
    j = 0
    s = App.Path + "\address1.dat"
    Open s For Output As #2
       Do Until j = smax
           Write #2, sair(j).first
           Write #2, sair(j).second
           Write #2, sair(j).relation
           Write #2, sair(j).place
           Write #2, sair(j).town
           Write #2, sair(j).pin
           Write #2, sair(j).district
           Write #2, sair(j).state
           Write #2, sair(j).country
           Write #2, sair(j).telno
           Write #2, sair(j).fax
           Write #2, sair(j).email
           Write #2, sair(j).Web
           Write #2, sair(j).notes
           Write #2, sair(j).photo
           j = j + 1
          
           Loop
        Close #2
        Unload Me
      End Sub


Private Sub Form_Load()
    i = 0
    smax = 0
    ppath = ""
'    GoTo anil
    On Error GoTo anil
    s = App.Path + "\address1.dat"
    Open s For Input As #1
       Do Until EOF(1)
           ReDim Preserve sair(smax + 1)
           Input #1, sair(smax).first
           Input #1, sair(smax).second
           Input #1, sair(smax).relation
           Input #1, sair(smax).place
           Input #1, sair(smax).town
           Input #1, sair(smax).pin
           Input #1, sair(smax).district
           Input #1, sair(smax).state
           Input #1, sair(smax).country
           Input #1, sair(smax).telno
           Input #1, sair(smax).fax
           Input #1, sair(smax).email
           Input #1, sair(smax).Web
           Input #1, sair(smax).notes
           Input #1, sair(smax).photo
           smax = smax + 1
        Loop

        Close #1
        i = smax - 1

    If i > 0 Or i = 0 Then
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).Web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If StrComp(ppath, "") <> 0 Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
    End If
anil:
        
End Sub


Private Sub Image3_Click()
brow.Show
If (ppath <> "") Then
     On Error GoTo anil
    Image3.Picture = LoadPicture(ppath)
Else
  '  MsgBox ("anil")
    Image3.Picture = Image4.Picture
End If
anil:
End Sub


