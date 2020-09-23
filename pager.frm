VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86E75BE0-83F1-11CF-A8A0-444553540000}#1.0#0"; "CSRAS32.OCX"
Begin VB.Form pager 
   BackColor       =   &H00000000&
   Caption         =   "Psyco Softwarez - Meeting Escaper"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   Icon            =   "pager.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3660
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin RasDialerCtrl.Dialer Dialer1 
      Left            =   960
      Top             =   960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoUpdate      =   0   'False
      Callback        =   0   'False
      CallbackNumber  =   ""
      DialogState     =   0
      Interval        =   1000
      PhoneBook       =   ""
      PhoneEntry      =   ""
      PhoneNumber     =   ""
      Timeout         =   60
      UserDomain      =   ""
      UserName        =   ""
      AutoConnect     =   0   'False
      AutoDisconnect  =   -1  'True
      Blocking        =   0   'False
   End
   Begin VB.CommandButton Command99 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "*"
      DownPicture     =   "pager.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "#"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox phonenumber 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   5175
   End
   Begin VB.CommandButton Command 
      Caption         =   "Start"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtcurrent 
      Height          =   345
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txttimetocall 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   960
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   960
      Top             =   960
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   960
      Top             =   960
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   960
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time To Call"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number To Call"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Time Now"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "pager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
phonenumber.Text = phonenumber.Text & "1"
End Sub

Private Sub Command10_Click()
phonenumber.Text = phonenumber.Text & "0"
End Sub

Private Sub Command11_Click()
phonenumber.Text = phonenumber.Text & "*"
End Sub

Private Sub Command12_Click()
phonenumber.Text = phonenumber.Text & "#"
End Sub

Private Sub Command13_Click()
phonenumber.Text = phonenumber.Text & " "
End Sub
Private Sub Command2_Click()
phonenumber.Text = phonenumber.Text & "2"
End Sub


Private Sub Command3_Click()
phonenumber.Text = phonenumber.Text & "3"
End Sub

Private Sub Command4_Click()
phonenumber.Text = phonenumber.Text & "4"
End Sub

Private Sub Command5_Click()
phonenumber.Text = phonenumber.Text & "5"
End Sub

Private Sub Command6_Click()
phonenumber.Text = phonenumber.Text & "6"
End Sub

Private Sub Command7_Click()
phonenumber.Text = phonenumber.Text & "7"
End Sub

Private Sub Command8_Click()
phonenumber.Text = phonenumber.Text & "8"
End Sub

Private Sub Command9_Click()
phonenumber.Text = phonenumber.Text & "9"
End Sub


Private Sub Command99_Click()
phonenumber.Text = phonenumber.Text & " "
End Sub

Private Sub Command_Click()

If phonenumber.Text = "" Then   'If the user forgets to enter their phone
                                'number, they get a message box that says:
MsgBox ("You need to enter a PHONE NUMBER TO CALL")
phonenumber.SetFocus    'This returns the user back to the Phone Number text box
Else    'If the user does have a phone number entered
pager.Hide   'Hide Personal Pager from pointy haired boss, nosy
                'co-workers, etc.  This could be changed to also hide the app from
                'CTL+ALT+DEL but I didn't think it was necessary.  My thought is that
                'most people won't CTL+ALT+DEL someone else's box, especially at work.
                
MsgBox ("Be sure your cell phone/pager is TURNED ON and NOT SET TO VIBRATE!") 'self explanatory
Timer1.Enabled = True   'Start the timer that compares the time to call(AlarmTime)
                        'with the current time displayed in txtCurrent.text
End If
End Sub



Private Sub Timer1_Timer()
AlarmTime = Val(txttimetocall.Text)         'Change the text in txtTimeToCall into a value
If AlarmTime = Val(txtcurrent.Text) Then    'Compare AlarmTime to the value of
                                            'txtCurrent.text.  If they are the
                                            'same,
Timer1.Enabled = False                      'turn the timer off and dial the number

    Dialer1.phonenumber = Trim$(phonenumber.Text) + ",,,," + phonenumber.Text 'you may need
                                                                               'to add or
                                                                               'subtract commas
                                                                               'to adjust the
                                                                               'time between when
                                                                               'the dialer connects
                                                                               'and the dialer redials
                                                                               'the number.  Each
                                                                               'comma is = 1 second.
        Dialer1.Connect
Timer3.Enabled = True   'Start timer3, which has the code to disconnect the dialer
End If
End Sub

Private Sub Timer2_Timer()
'This code puts the current system time into txtCurrent.text
If txtcurrent.Text <> CStr(time) Then
    txtcurrent.Text = Format(time, "hh:mm:ss")  'Read the VB help to learn more about Format
End If
End Sub

Private Sub Timer3_Timer()
'You may have to adjust the interval of the timer so that it doesn't hang up too soon!
Dialer1.Disconnect  'This line of code disconnects the RAS Dialer
Unload Me   'Unload Personal Pager
End Sub

