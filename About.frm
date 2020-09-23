VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Electronic Clock"
   ClientHeight    =   2400
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5055
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1656.523
   ScaleMode       =   0  'User
   ScaleWidth      =   4746.906
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "About.frx":0CCA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3570
      TabIndex        =   0
      Top             =   1995
      Width           =   1260
   End
   Begin VB.Label Label3 
      Caption         =   "E-Mail: edgars53@inbox.lv"
      Height          =   195
      Left            =   1425
      TabIndex        =   6
      Top             =   1590
      Width           =   2250
   End
   Begin VB.Label Label2 
      Caption         =   "Graphics: Edgars Priede"
      Height          =   255
      Left            =   1410
      TabIndex        =   5
      Top             =   1290
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Programmer: Edgars Priede"
      Height          =   225
      Left            =   1410
      TabIndex        =   4
      Top             =   990
      Width           =   2745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4620.134
      Y1              =   1314.865
      Y2              =   1314.865
   End
   Begin VB.Label lblTitle 
      Caption         =   "Electronic Clock"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   126.772
      X2              =   4620.134
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version: 1.1"
      Height          =   225
      Left            =   1410
      TabIndex        =   3
      Top             =   690
      Width           =   3885
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
FormOnTop Me.hWnd, True
End Sub

