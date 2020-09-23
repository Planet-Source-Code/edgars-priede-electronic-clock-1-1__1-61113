VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Run 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Run some program at designed time..."
   ClientHeight    =   2175
   ClientLeft      =   6495
   ClientTop       =   4320
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4500
   Begin VB.CommandButton Close 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3690
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTime 
      Height          =   315
      Left            =   750
      MaxLength       =   8
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton Browse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   3540
      TabIndex        =   2
      Top             =   1110
      Width           =   795
   End
   Begin VB.TextBox txtProgram 
      Height          =   315
      Left            =   750
      TabIndex        =   1
      Top             =   1110
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Run program on time (hh:mm:ss):"
      Height          =   255
      Left            =   750
      TabIndex        =   3
      Top             =   120
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Program to run:"
      Height          =   225
      Left            =   750
      TabIndex        =   0
      Top             =   870
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Run.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   4350
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4350
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "Run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Browse_Click()
    With CD
        .Filter = "Executable Files (*.exe)|*.exe|MS-DOS Command Files (*.com)|*.com|Batch Files (*.bat)|*.bat|All Files (*.*)|*.*"
        .ShowOpen
        txtProgram = .FileName
    End With
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Form_Load()
FormOnTop Me.hWnd, True
txtTime.Text = Main.txtRunTime
txtProgram.Text = Main.txtProgram
End Sub

Private Sub OK_Click()
If txtTime.Text = "" Then Main.P.Visible = False
Main.txtRunTime = txtTime
Main.txtProgram = txtProgram
Main.P.Visible = True
Unload Me
End Sub
