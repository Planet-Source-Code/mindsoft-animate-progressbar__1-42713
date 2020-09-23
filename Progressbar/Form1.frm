VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   4905
   ClientTop       =   4080
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Install"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   960
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Activate Progressbar using timer
'Level Begginer
'Please Dont Forget to vote and add your comments
'-------------------------------------------------------------------


Private Sub Command1_Click()
'Enable Timer
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
End 'Close App
End Sub

Private Sub Command3_Click()
'Disable Timer
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Timer1.Enabled = False 'Disable the timer on app start up
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1 'Move Progressbar up 1 value
If ProgressBar1.Value = 100 Then
Timer1.Enabled = False
ProgressBar1.Value = 0 'Reset Progressbar
MsgBox "Finished Installing", vbInformation, "Done"
End If
End Sub
