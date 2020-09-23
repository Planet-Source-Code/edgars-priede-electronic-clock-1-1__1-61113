VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4545
   ClientLeft      =   6330
   ClientTop       =   3060
   ClientWidth     =   9780
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   652
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOnTop 
      Caption         =   "On Top"
      Height          =   255
      Left            =   8130
      TabIndex        =   14
      Top             =   780
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.TextBox txtReminderText 
      Height          =   435
      Left            =   6420
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtReminderTime 
      Height          =   435
      Left            =   6420
      TabIndex        =   12
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtProgram 
      Height          =   345
      Left            =   330
      TabIndex        =   9
      Top             =   2100
      Width           =   1665
   End
   Begin VB.TextBox txtRunTime 
      Height          =   345
      Left            =   330
      TabIndex        =   8
      Top             =   1560
      Width           =   1665
   End
   Begin VB.Timer tmrMore2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   2520
   End
   Begin VB.Timer tmrMore 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3420
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   735
      Left            =   8070
      TabIndex        =   5
      Top             =   30
      Width           =   960
   End
   Begin VB.Timer tmrTime 
      Interval        =   990
      Left            =   3810
      Top             =   1950
   End
   Begin VB.OptionButton optLCD 
      Caption         =   "LCD"
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   870
      Width           =   1245
   End
   Begin VB.OptionButton optBlue 
      Caption         =   "Blue LED's"
      Height          =   195
      Left            =   4800
      TabIndex        =   3
      Top             =   660
      Width           =   1275
   End
   Begin VB.OptionButton optGreen 
      Caption         =   "Green LED's"
      Height          =   195
      Left            =   4800
      TabIndex        =   2
      Top             =   450
      Width           =   1275
   End
   Begin VB.OptionButton optYellow 
      Caption         =   "Yellow LED's"
      Height          =   195
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   1275
   End
   Begin VB.OptionButton optRed 
      Caption         =   "Red LED's"
      Height          =   195
      Left            =   4800
      TabIndex        =   0
      Top             =   30
      Width           =   1275
   End
   Begin VB.Timer tmrPoints2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3030
      Top             =   1950
   End
   Begin VB.Timer tmrPoints1 
      Interval        =   500
      Left            =   2610
      Top             =   1950
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Run some program at designed time"
      Height          =   1035
      Left            =   7110
      TabIndex        =   7
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reminder"
      Height          =   1035
      Left            =   6150
      TabIndex        =   6
      Top             =   30
      Width           =   975
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      ToolTipText     =   "Run some program at designed time is on..."
      Top             =   750
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label R 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      ToolTipText     =   "Reminder is on..."
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgActive2 
      Height          =   1065
      Left            =   5280
      Picture         =   "Main.frx":0CCA
      Top             =   2010
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgInactive2 
      DataSource      =   "imgInactive2"
      Height          =   1065
      Left            =   4995
      Picture         =   "Main.frx":1A5C
      Top             =   2010
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgActive 
      Height          =   1065
      Left            =   4710
      Picture         =   "Main.frx":27EE
      Top             =   2010
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgInactive 
      Height          =   1065
      Left            =   4440
      Picture         =   "Main.frx":3580
      Top             =   2010
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Line Line1 
      X1              =   405
      X2              =   405
      Y1              =   2
      Y2              =   72
   End
   Begin VB.Shape Point1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   3
      Left            =   2820
      Top             =   645
      Width           =   75
   End
   Begin VB.Shape Point1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   2
      Left            =   2820
      Top             =   375
      Width           =   75
   End
   Begin VB.Shape Point1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   1
      Left            =   1410
      Top             =   645
      Width           =   75
   End
   Begin VB.Image Image6 
      Height          =   885
      Left            =   3600
      Top             =   105
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   885
      Left            =   3000
      Top             =   105
      Width           =   495
   End
   Begin VB.Shape Point1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   0
      Left            =   1410
      Top             =   375
      Width           =   75
   End
   Begin VB.Image Image4 
      Height          =   885
      Left            =   2190
      Top             =   105
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   1590
      Top             =   105
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   780
      Top             =   105
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   180
      Top             =   105
      Width           =   495
   End
   Begin VB.Shape Back 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   1065
      Left            =   15
      Top             =   15
      Width           =   4455
   End
   Begin VB.Image imgMore 
      Height          =   1065
      Left            =   4485
      Picture         =   "Main.frx":4312
      ToolTipText     =   "Options..."
      Top             =   15
      Width           =   225
   End
   Begin VB.Menu m_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu m_Show 
         Caption         =   "Show"
      End
      Begin VB.Menu separator3 
         Caption         =   "-"
      End
      Begin VB.Menu m_Reminder 
         Caption         =   "Reminder..."
      End
      Begin VB.Menu m_Run 
         Caption         =   "Run some program..."
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu m_About 
         Caption         =   "About"
      End
      Begin VB.Menu separator2 
         Caption         =   "-"
      End
      Begin VB.Menu m_Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SegColor As String
Dim optDefault As String
Dim CurrentTime As String
Dim h1, h2, m1, m2, s1, s2 As String


Private Sub chkOnTop_Click()
FormOnTop Me.hWnd, chkOnTop.Value
End Sub

Private Sub Command1_Click()
About.Show
End Sub

Private Sub Command2_Click()
Reminder.Show
End Sub

Private Sub Command3_Click()
Run.Show
End Sub

Private Sub Form_Load()
FormOnTop Me.hWnd, chkOnTop.Value
    Main.Width = 4815
    Main.Height = 1455
    GetCurrentTime
    
    ''SegColor = INIGetSettingString("Segments", "Color", App.Path & "\cache.dat")
    ''Back.BackColor = INIGetSettingString("Back", "Color", App.Path & "\cache.dat")
    optDefault = INIGetSettingString("opt", "Default", App.Path & "\cache.dat")
    txtReminderTime.Text = INIGetSettingString("Reminder", "Time", App.Path & "\cache.dat")
    txtReminderText.Text = INIGetSettingString("Reminder", "Text", App.Path & "\cache.dat")
    txtRunTime.Text = INIGetSettingString("Run", "Time", App.Path & "\cache.dat")
    txtProgram.Text = INIGetSettingString("Run", "Path", App.Path & "\cache.dat")
    R.Visible = INIGetSettingString("Reminder", "R", App.Path & "\cache.dat")
    P.Visible = INIGetSettingString("Run", "P", App.Path & "\cache.dat")


        Select Case optDefault
            Case "optRed"
                optRed = True
            Case "optYellow"
                optYellow = True
            Case "optGreen"
                optGreen = True
            Case "optBlue"
                optBlue = True
            Case "optLCD"
                optLCD = True
        End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ''INIWriteSetting "Segments", "Color", SegColor, App.Path & "\cache.dat"
    ''INIWriteSetting "Back", "Color", Back.BackColor, App.Path & "\cache.dat"
    INIWriteSetting "opt", "Default", optDefault, App.Path & "\cache.dat"
    INIWriteSetting "Reminder", "Time", txtReminderTime.Text, App.Path & "\cache.dat"
    INIWriteSetting "Reminder", "Text", txtReminderText.Text, App.Path & "\cache.dat"
    INIWriteSetting "Run", "Time", txtRunTime.Text, App.Path & "\cache.dat"
    INIWriteSetting "Run", "Path", txtProgram.Text, App.Path & "\cache.dat"
    INIWriteSetting "Reminder", "R", R.Visible, App.Path & "\cache.dat"
    INIWriteSetting "Run", "P", P.Visible, App.Path & "\cache.dat"
End Sub

Private Sub imgMore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Main.Width > 9150 Then
        imgMore.Picture = imgActive2.Picture
        tmrMore2.Enabled = True
    End If
    If Main.Width < 9150 Then
        imgMore.Picture = imgActive.Picture
        tmrMore.Enabled = True
    End If
End Sub

Private Sub imgMore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMore.Picture = imgInactive.Picture
End Sub

Private Sub optBlue_Click()
Dim i As Integer

    If optBlue = True Then
    For i = 0 To 3
        Point1(i).BorderColor = &HFFFF00
        Point1(i).BackColor = &HFFFF00
    Next i
    
    P.ForeColor = &HFFFF00
    R.ForeColor = &HFFFF00

        SegColor = "B"
        Back.BackColor = &H0&
        optDefault = "optBlue"
        DrawSegments
    End If
End Sub

Private Sub optGreen_Click()
Dim i As Integer

    For i = 0 To 3
        Point1(i).BorderColor = &HFF00&
        Point1(i).BackColor = &HFF00&
    Next i
    
    P.ForeColor = &HFF00&
    R.ForeColor = &HFF00&

    If optGreen = True Then
        SegColor = "G"
        Back.BackColor = &H0&
        optDefault = "optGreen"
        DrawSegments
    End If
End Sub

Private Sub optLCD_Click()
Dim i As Integer

    For i = 0 To 3
        Point1(i).BorderColor = &H0&
        Point1(i).BackColor = &H0&
    Next i
    
    P.ForeColor = &H0&
    R.ForeColor = &H0&

    If optLCD = True Then
        SegColor = "L"
        Back.BackColor = &HC0C0C0
        optDefault = "optLCD"
        DrawSegments
    End If
End Sub

Private Sub optRed_Click()
Dim i As Integer

    For i = 0 To 3
        Point1(i).BorderColor = &HFF&
        Point1(i).BackColor = &HFF&
    Next i
    
    P.ForeColor = &HFF&
    R.ForeColor = &HFF&

    If optRed = True Then
        SegColor = "R"
        Back.BackColor = &H0&
        optDefault = "optRed"
        DrawSegments
    End If
End Sub

Private Sub optYellow_Click()
Dim i As Integer

    For i = 0 To 3
        Point1(i).BorderColor = &HFFFF&
        Point1(i).BackColor = &HFFFF&
    Next i

    P.ForeColor = &HFFFF&
    R.ForeColor = &HFFFF&

    If optYellow = True Then
        SegColor = "Y"
        Back.BackColor = &H0&
        optDefault = "optYellow"
        DrawSegments
    End If
End Sub

Private Sub tmrMore_Timer()
    Main.Width = Main.Width + 55
    If Main.Width > 9150 Then
    tmrMore.Enabled = False
    imgMore.Picture = imgInactive2.Picture
    End If
End Sub

Private Sub tmrMore2_Timer()
    Main.Width = Main.Width - 55
    If Main.Width < 4816 Then
    tmrMore2.Enabled = False
    imgMore.Picture = imgInactive.Picture
    End If
End Sub

Private Sub tmrPoints1_Timer()
Dim i As Integer

    For i = 0 To 3
        Point1(i).Visible = False
    Next i
    
    tmrPoints2.Enabled = True
    tmrPoints1.Enabled = False
End Sub

Private Sub tmrPoints2_Timer()
Dim i As Integer

    For i = 0 To 3
        Point1(i).Visible = True
    Next i
    
    tmrPoints1.Enabled = True
    tmrPoints2.Enabled = False
End Sub

Private Sub DrawSegments()
    Image1.Picture = LoadResPicture(h1 & SegColor, bitmap)
    Image2.Picture = LoadResPicture(h2 & SegColor, bitmap)
    Image3.Picture = LoadResPicture(m1 & SegColor, bitmap)
    Image4.Picture = LoadResPicture(m2 & SegColor, bitmap)
    Image5.Picture = LoadResPicture(s1 & SegColor, bitmap)
    Image6.Picture = LoadResPicture(s2 & SegColor, bitmap)
End Sub

Private Sub GetCurrentTime()
    CurrentTime = Time
    If Len(CurrentTime) = 7 Then CurrentTime = "N" & CurrentTime
    h1 = Mid(CurrentTime, 1, 1)
    h2 = Mid(CurrentTime, 2, 1)
    m1 = Mid(CurrentTime, 4, 1)
    m2 = Mid(CurrentTime, 5, 1)
    s1 = Mid(CurrentTime, 7, 1)
    s2 = Mid(CurrentTime, 8, 1)
End Sub

Private Sub tmrTime_Timer()
    GetCurrentTime
    DrawSegments
    RunProgram
    Remind
End Sub

Private Sub RunProgram()
    If txtRunTime.Text = Time Then
        Shell txtProgram.Text, vbNormalFocus
        MsgBox txtRunTime.Text & " - Program is running!", vbInformation, "Electronic Clock"
        P.Visible = False
        txtProgram.Text = ""
        txtRunTime.Text = ""
    End If
End Sub

Private Sub Remind()
    If txtReminderTime.Text = Time Then
        MsgBox txtReminderText.Text, vbInformation, "Reminder"
        R.Visible = False
        txtReminderTime.Text = ""
        txtReminderText.Text = ""
    End If
End Sub
