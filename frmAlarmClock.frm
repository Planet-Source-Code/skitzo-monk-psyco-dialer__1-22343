VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmAlarmClock 
   BackColor       =   &H00000000&
   Caption         =   "Psyco Softwarez - Alarm Clock "
   ClientHeight    =   4185
   ClientLeft      =   2265
   ClientTop       =   2745
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlarmClock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSnooze 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer tmrTimer3 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1080
      Top             =   0
   End
   Begin MCI.MMControl MMControl2 
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrTimer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   0
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   3840
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrTimer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer tmrTimer0 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdSnooze 
      Caption         =   "&Snooze"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton cmdAlarmOnOff 
      Caption         =   "&Off"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Frame frmCommands 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   2880
      Width           =   6375
   End
   Begin VB.CommandButton cmdUpdateSettings 
      Caption         =   "&Update Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.OptionButton optSound 
      BackColor       =   &H00000000&
      Caption         =   "S&ound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optSetAlarm 
      BackColor       =   &H00000000&
      Caption         =   "&Set Alarm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Alarm Options"
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   6375
   End
   Begin VB.TextBox txtAlarmClock 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Alarm Clock Display"
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmAlarmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub cmdAlarmOnOff_Click()
     cmdAlarmOnOff.Caption = AlarmDisabled
     cmdAlarmOnOff.BackColor = vbButtonFace
     txtStatus.Text = "Alarm is Disabled"
     tmrTimer1.Enabled = False
     tmrTimer3.Enabled = False
     tmrSnooze.Enabled = False
End Sub
Private Sub cmdSnooze_Click()
     tmrTimer1.Enabled = False
     tmrTimer3.Enabled = False
     txtStatus.Text = "Snoozing for 10 Minutes"
     tmrSnooze.Enabled = True
     Snooze = True
     cmdAlarmOnOff.Caption = "Snoozing"
     cmdAlarmOnOff.BackColor = vbGreen
End Sub
Private Sub cmdUpdateSettings_Click()
     If optSetAlarm.Value = True Then
     Hours = InputBox("What hour of the day do you want to set the alarm for?", "Alarm Clock Setup")
     Minutes = InputBox("How many minutes after the hour?", "Alarm Clock Setup")
     DayNight = InputBox("AM or PM", "Alarm Clock Setup")
     AlarmTime = Hours & ":" & Minutes & ":00"
     txtStatus.Text = "AlarmOn"
         If txtStatus.Text = "AlarmOn" Then
             cmdAlarmOnOff.Caption = AlarmEnabled
         End If
         ElseIf optSound.Value = True Then
             msg = "Do you want the clock to tick?"
             Style = vbYesNo + vbQuestion
             Title = "Sound Settings"
             Response = MsgBox(msg, Style, Title)
         If Response = vbYes Then
             MyString = "Yes"
             tmrTimer2.Enabled = True
             Audio = True
         Else
             MyString = "No"
             tmrTimer2.Enabled = False
             Audio = False
         End If
             msg = "Do you want a silent alarm?"
             Style = vbYesNo + vbQuestion
             Title = "Sound Settings"
             Response = MsgBox(msg, Style, Title)
         If Response = vbYes Then
             MyString = "Yes"
             Bell = False
             Else
             MyString = "No"
             Bell = True
         End If
     End If
End Sub
Private Sub Form_Load()
     Me.Show
     Me.Refresh
     With nid
         .cbSize = Len(nid)
         .hwnd = Me.hwnd
         .uId = vbNull
         .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         .uCallBackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon
         .szTip = "Programming Exercise 5" & vbNullChar
     End With
         Shell_NotifyIcon NIM_ADD, nid
         tmrTimer0.Enabled = True
         tmrTimer0.Interval = 1000
         AlarmOn = "Alarm is Enabled"
         AlarmOff = "Alarm is Disabled"
         txtStatus.Text = AlarmOff
         tmrTimer1.Enabled = False
         tmrTimer2.Enabled = False
         tmrTimer3.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
         MMControl1.Command = "Close"
         MMControl2.Command = "Close"
'this removes the icon from the system tray
         Shell_NotifyIcon NIM_DELETE, nid
End Sub






Private Sub tmrTimer0_Timer()
     GMT = time$
     GMTD = Date$
     txtAlarmClock.Text = GMT
     txtDate.Text = GMTD
     If Audio = True Then
         tmrTimer2.Enabled = True
     ElseIf Audio = False Then
         tmrTimer2.Enabled = False
     End If
     If Bell = True Then
         If txtAlarmClock.Text = AlarmTime Then
         tmrTimer1.Enabled = True
         tmrTimer3.Enabled = True
         txtStatus.Text = "Alarm Ringing!!!"
         cmdAlarmOnOff.Caption = StopAlarm
         cmdAlarmOnOff.BackColor = vbRed
         End If
     ElseIf Bell = False Then
         If txtAlarmClock.Text = AlarmTime Then
         tmrTimer3.Enabled = True
         txtStatus.Text = "Alarm Ringing!!!"
         cmdAlarmOnOff.Caption = StopAlarm
         cmdAlarmOnOff.BackColor = vbRed
         End If
     End If
End Sub
Private Sub tmrTimer1_Timer()
     MMControl1.Command = "Close"
     MMControl1.Notify = False
     MMControl1.wait = True
     MMControl1.Shareable = False
     MMControl1.DeviceType = "WaveAudio"
     MMControl1.FileName = App.Path & "\WAKEUP.WAV"
     MMControl1.Command = "Open"
     MMControl1.Command = "Play"
End Sub
Private Sub tmrTimer2_Timer()
     MMControl2.Command = "Close"
     MMControl2.Notify = False
     MMControl2.wait = True
     MMControl2.Shareable = False
     MMControl2.DeviceType = "WaveAudio"
     MMControl2.FileName = App.Path & "\TICK.WAV"
     MMControl2.Command = "Open"
     MMControl2.Command = "Play"
End Sub
Private Sub tmrTimer3_Timer()
     Static Flash As Boolean
         Flash = Not Flash
     Select Case Flash
        Case True
         cmdAlarmOnOff.BackColor = vbRed
        Case False
         cmdAlarmOnOff.BackColor = vbYellow
     End Select
End Sub
Private Sub tmrSnooze_Timer()
     If SnoozeTime = 10 Then
         tmrTimer1.Enabled = True
         tmrTimer3.Enabled = True
         SnoozeTime = 0
         cmdAlarmOnOff.Caption = StopAlarm
     Else
         SnoozeTime = SnoozeTime + 1
         tmrSnooze.Enabled = True
     End If
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
Private Sub Form_Resize()
'this is necessary to assure that the minimized window is hidden
     If Me.WindowState = vbMinimized Then Me.Hide
End Sub
Private Sub mPopExit_Click()
'called when user clicks the popup menu Exit command
     Unload Me
End Sub
Private Sub mPopRestore_Click()
'called when the user clicks the popup menu Restore command
     Me.WindowState = vbNormal
     Result = SetForegroundWindow(Me.hwnd)
     Me.Show
End Sub



