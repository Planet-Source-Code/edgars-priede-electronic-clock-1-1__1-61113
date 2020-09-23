VERSION 5.00
Begin VB.Form Reminder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reminder"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3090
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   750
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox ReTime 
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   360
      Width           =   3645
   End
   Begin VB.TextBox Text 
      Height          =   1005
      Left            =   750
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1050
      Width           =   3675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5010
      Y1              =   2145
      Y2              =   2145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5010
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label2 
      Caption         =   "Remind me at this time (hh:mm:ss):"
      Height          =   225
      Left            =   750
      TabIndex        =   3
      Top             =   120
      Width           =   2745
   End
   Begin VB.Label Label1 
      Caption         =   "Enter reminder text:"
      Height          =   225
      Left            =   750
      TabIndex        =   1
      Top             =   810
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Reminder.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Reminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
FormOnTop Me.hWnd, True
ReTime.Text = Main.txtReminderTime
Text.Text = Main.txtReminderText
End Sub

Private Sub OK_Click()
If ReTime.Text = "" Then Main.R.Visible = False
Main.txtReminderTime.Text = ReTime.Text
Main.txtReminderText = Text.Text
Main.R.Visible = True
Unload Me
End Sub
