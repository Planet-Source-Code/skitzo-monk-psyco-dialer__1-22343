VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Psyco Softwarez  X-Treme Dialer"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6030
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   2160
      Top             =   5280
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2520
      TabIndex        =   22
      Top             =   240
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   5640
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2160
      Top             =   6000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   21
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Address Book"
      DisabledPicture =   "Form1.frx":030A
      Height          =   375
      Left            =   2880
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   20
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Reset Dial"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Hang Up"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Dial"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Web Browser"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Timed Dial"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Alarm Clock"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Hide All"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "#"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "*"
      DownPicture     =   "Form1.frx":0614
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
      TabIndex        =   11
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   5175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   960
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Current Time"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   23
      Top             =   240
      Width           =   1095
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuwhat 
         Caption         =   "What This Does"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

DefInt A-Z
Dim GMT As String
Dim GMTD As String
Dim AlarmOn As String
Dim AlarmOff As String
Dim Hours As String
Dim Minutes As String
Dim DayNight As String
Dim AlarmTime As String
Dim Result As Long
Dim msg, Style, Title, Response, MyString
Public SnoozeTime As Integer
Public Snooze  As Boolean
Public AlarmOn1 As Boolean
Public Audio As Boolean
Public Bell As Boolean

Const EnableAlarm = "ENABLE  ALARM"
Const AlarmEnabled = "ALARM  ENABLED"
Const AlarmDisabled = "ALARM DISABLED"
Const StopAlarm = "STOP  ALARM"
' This flag is set when the user chooses Cancel.
Dim CancelFlag

Private Sub Command1_Click()
Text1.Text = Text1.Text & "1"
End Sub

Private Sub Command10_Click()
Text1.Text = Text1.Text & "0"
End Sub

Private Sub Command11_Click()
Text1.Text = Text1.Text & "*"
End Sub

Private Sub Command12_Click()
Text1.Text = Text1.Text & "#"
End Sub

Private Sub Command13_Click()
Text1.Text = Text1.Text & " "
End Sub


Private Sub Command14_Click()
address.Hide
Form1.Hide
frmAlarmClock.Hide
pager.Hide
frmBrowser.Hide
Me.Hide
        ShowIcon Me
End Sub

Private Sub Command15_Click()
frmAlarmClock.Show
End Sub

Private Sub Command23_Click()
address.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
Dim msg As Long
'the value of X will vary depending upon the scalemode setting
     If Me.ScaleMode = vbPixels Then
         msg = x
     Else
         msg = x / Screen.TwipsPerPixelX
     End If
     Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
     End Select
End Sub

Private Sub Command17_Click()
pager.Show
End Sub

Private Sub Command18_Click()
frmBrowser.Show
End Sub

Private Sub Command20_Click()
CancelFlag = True
End Sub

Private Sub Command666_Click()
Hours = InputBox("What hour of the day do you want to set the alarm for?", "Alarm Clock Setup")
     Minutes = InputBox("How many minutes after the hour?", "Alarm Clock Setup")
     DayNight = InputBox("AM or PM", "Alarm Clock Setup")
     AlarmTime = Hours & ":" & Minutes & ":00"
     MsgBox ("Time set")
        
End Sub

Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuwhat_Click()
frmwhat.Show
End Sub

Private Sub Timer7_Timer()
'This code puts the current system time into txtCurrent.text
If Text3.Text <> CStr(time) Then
    Text3.Text = Format(time, "hh:mm:ss")  'Read the VB help to learn more about Format
End If
End Sub

Private Sub Form_Load()
MSComm1.InputLen = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Remove the icon when this form unload. Don't forget to unload this form!
    RemoveIcon Me 'Add your form's name here for the sub to work.
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        ' This code hides the Form and puts the icon in the tray. Feel free to move
        ' it around if you like.
        Me.Hide
        ShowIcon Me
    End If
End Sub


Private Sub Timer1_Timer()
If Text1.Text = "" Then
Command11.Enabled = False
Command12.Enabled = False
Else
Command11.Enabled = True
Command12.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
Dim port, x, instring
port = 1
PortinG:
MSComm1.CommPort = port
MSComm1.PortOpen = True


Form1.MSComm1.Settings = "9600,N,8,1"
    MSComm1.Output = "AT" + Chr$(13)
    x = 1


    Do: DoEvents
        x = x + 1
        If x = 1000 Then MSComm1.Output = "AT" + Chr$(13)
        If x = 2000 Then MSComm1.Output = "AT" + Chr$(13)
        If x = 3000 Then MSComm1.Output = "AT" + Chr$(13)
        If x = 4000 Then MSComm1.Output = "AT" + Chr$(13)
        If x = 5000 Then MSComm1.Output = "AT" + Chr$(13)
        If x = 6000 Then MSComm1.Output = "AT" + Chr$(13)


        If x = 7000 Then
            MSComm1.PortOpen = False
            port = port + 1
            GoTo PortinG:


            If MSComm1.CommPort >= 5 Then
errr:
                MsgBox "Can't Find Modem!"
                GoTo done:
            End If
        End If
    Loop Until MSComm1.InBufferCount >= 2
    instring = MSComm1.Input
    MSComm1.PortOpen = False

  Text2.Text = port



done:
Timer2.Enabled = False
End Sub
Sub Dial(Number$)
 Dim DialString$, FromModem$, dummy

    
    DialString$ = "ATDT" + Number$ + ";" + vbCr

   
    MSComm1.CommPort = Text2.Text
    MSComm1.Settings = "9600,N,8,1"
    
   
    On Error Resume Next
    MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2: not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    
    MSComm1.InBufferCount = 0
    
   
    MSComm1.Output = DialString$
    
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, "OK") Then
             ' Notify the user to pick up the phone.
             Beep
          
             Exit Do
          End If
       End If
        
       ' Did the user choose Cancel?
       If CancelFlag Then
          CancelFlag = False
          Exit Do
       End If
    Loop
    
    ' Disconnect the modem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Close the port.
    MSComm1.PortOpen = False
End Sub

Private Sub Command19_Click()
Dial Text1.Text
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & "2"
End Sub

Private Sub Command22_Click()
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & "3"
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text & "4"
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text & "5"
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text & "6"
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text & "7"
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text & "8"
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text & "9"
End Sub

