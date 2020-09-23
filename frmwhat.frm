VERSION 5.00
Begin VB.Form frmwhat 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "What This is Used For? "
   ClientHeight    =   3465
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6090
   Icon            =   "frmwhat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Web 
      Caption         =   "Visit Our Website"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   $"frmwhat.frx":0442
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
      Height          =   2415
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "So What Is This Used For?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmwhat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

End Sub

Private Sub OKButton_Click()
frmwhat.Hide
End Sub

Private Sub Web_Click()
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate "www.skitzoduck.4dw.com"
frmBrowser.cboAddress = "www.skitzoduck.4dw.com"
End Sub
