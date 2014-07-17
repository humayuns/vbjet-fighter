VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   738
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox en 
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   0
      Left            =   11700
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   660
      ScaleWidth      =   2250
      TabIndex        =   16
      Top             =   3300
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Timer Timer3 
      Left            =   8610
      Top             =   3990
   End
   Begin VB.PictureBox shooter 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   1440
      Picture         =   "Form1.frx":0E31
      ScaleHeight     =   660
      ScaleWidth      =   2250
      TabIndex        =   13
      Top             =   5280
      Width           =   2250
   End
   Begin VB.PictureBox Picture1 
      Height          =   270
      Left            =   7140
      ScaleHeight     =   210
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   7890
      Width           =   285
   End
   Begin VB.PictureBox fire 
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   1680
      Picture         =   "Form1.frx":182B
      ScaleHeight     =   90
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   8520
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8580
      Top             =   840
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   8505
      Picture         =   "Form1.frx":1887
      ScaleHeight     =   660
      ScaleWidth      =   2250
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   660
      Left            =   5655
      Picture         =   "Form1.frx":2281
      ScaleHeight     =   660
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   5205
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stage="
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   7920
      Width           =   615
   End
   Begin VB.Label lblStage 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   7920
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   2610
      Picture         =   "Form1.frx":2D26
      Top             =   7560
      Width           =   4500
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000008&
      Caption         =   "Label5"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9720
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Caption         =   "Label6"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9720
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000008&
      Caption         =   "Y = 0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "X = 0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape en1 
      BorderColor     =   &H80000009&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   4  'Upward Diagonal
      Height          =   495
      Left            =   11880
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Speed ="
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12720
      TabIndex        =   6
      Top             =   10680
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13320
      TabIndex        =   5
      Top             =   10680
      Width           =   615
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score="
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Player1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim u, d, l, r  As Boolean ' Basic true/false variables


Dim VBJet As New JETPlane
Dim VBJetMissile As New Missile

Dim EnemyJet() As New EnemyJetPlane
Dim EnemyJetCount As Integer
Dim EnemyJetMVar As Integer

Dim score As Long
Dim stage As Long
Dim level_score As Integer




Private Sub Form_Load()
    

    
    ' Startup and basic settings
    
    lblScore.Caption = "0"
    
    VBJet.x = 0
    VBJet.y = 0
    
    EnemyJetCount = 2
    EnemyJetMVar = 10
    
    ReDim EnemyJet(EnemyJetCount)
    
    Dim i As Integer
    For i = 1 To EnemyJetCount
        'EnemyJet(i).X = Rnd(Me.ScaleWidth / 2) + Me.ScaleWidth / 2
        'EnemyJet(i).Y = Rnd(Me.ScaleHeight)
        EnemyJet(i).speed = 5
        'On Error Resume Next
        Load en(i)
        en(i).Visible = True
        SetEn i
        
    Next i
    
    VBJet.Power = 1
    
    stage = 1

End Sub



Private Sub Form_Paint()
    shooter.SetFocus
End Sub

Private Sub shooter_KeyDown(KeyCode As Integer, Shift As Integer)

    '
    If KeyCode = 49 Then speed = speed - 1
    If speed <= 0 Then speed = 0
    If speed > 30 Then speed = 30
    
    If KeyCode = 50 Then speed = speed + 1
    If KeyCode = vbKeyLeft Then l = True
    If KeyCode = vbKeyRight Then r = True
    If KeyCode = vbKeyUp Then u = True
    If KeyCode = vbKeyDown Then d = True
    
    If KeyCode = vbKeySpace Then
        If Not VBJetMissile.Visible Then
            FireIt
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me: Form2.Show
    

End Sub



Private Sub shooter_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Check keyup event
    ' If keyup call then disable movement check
    If KeyCode = vbKeyLeft Then l = False
    If KeyCode = vbKeyRight Then r = False
    If KeyCode = vbKeyUp Then u = False
    If KeyCode = vbKeyDown Then d = False
    
End Sub

Private Sub Timer1_Timer()

    Static ch As Boolean
    
    ch = Not ch
    
    ' Animation steps of Jet Fighter
    
    If ch Then
        shooter.Picture = Picture2.Picture
    Else
        shooter.Picture = Picture3.Picture
    End If

End Sub

Private Sub Timer2_Timer()


    If l Then
        VBJet.x = VBJet.x - speed
        If VBJet.x < 0 Then VBJet.x = 0
    End If
    
    If r Then
        VBJet.x = VBJet.x + speed
        If VBJet.x >= Me.ScaleWidth - 100 Then VBJet.x = Me.ScaleWidth - 100
   End If
    
    If u Then
        VBJet.y = VBJet.y - speed
        If VBJet.y < 0 Then VBJet.y = 0
    End If
    
    If d Then
        VBJet.y = VBJet.y + speed
        If VBJet.y >= Me.ScaleHeight - 100 Then VBJet.y = Me.ScaleHeight - 100
    End If
    
    Label5.Caption = "X = " & VBJet.x
    Label6.Caption = "Y = " & VBJet.y
    
    shooter.Left = VBJet.x
    shooter.Top = VBJet.y
    
    Label3.Caption = CStr(speed)
    
    
    Dim i As Integer
    For i = 1 To EnemyJetCount
    
    If VBJetMissile.Visible Then
        
        VBJetMissile.x = VBJetMissile.x + 20
        If VBJetMissile.x > Me.ScaleWidth Then
            VBJetMissile.Visible = False
            fire.Visible = False
        
        End If
        
        fire.Left = VBJetMissile.x
        fire.Top = VBJetMissile.y
        
        
        If level_score = 9 Then
'            MsgBox "Good work Level completed!!!" & vbCrLf & "Now starting Level " & stage + 1, vbInformation, "Congratulations!"
            stage = stage + 1
            level_score = 0
            EnemyJet(i).speed = EnemyJet(i).speed + 5 ' Increase enemny speed

            en(i).Picture = LoadPicture((App.Path & "\data\op2.gif"))
        End If
        
        
        If (VBJetMissile.y > EnemyJet(i).y And VBJetMissile.y < EnemyJet(i).y + 60) And (VBJetMissile.x > EnemyJet(i).x) Then
            score = score + 10
            level_score = level_score + 1
            VBJetMissile.Visible = False
            SetEn i
        End If
        

    
    Else
        fire.Visible = False
    End If
    
    EnemyJet(i).x = EnemyJet(i).x - EnemyJet(i).speed
    If EnemyJet(i).y > VBJet.y Then
        EnemyJet(i).y = EnemyJet(i).y - (Rnd(EnemyJetMVar))
    Else
        EnemyJet(i).y = EnemyJet(i).y + (Rnd(EnemyJetMVar))
    End If
    
    en(i).Left = EnemyJet(i).x
    en(i).Top = EnemyJet(i).y
    
    If EnemyJet(i).x < -200 Then
        SetEn i
        en(i).Top = EnemyJet(i).y
    End If
    lblScore = CStr(score)
    lblStage = CStr(stage)
    
    Label8.Caption = "EX = " & EnemyJet(i).x
    Label7.Caption = "EY = " & EnemyJet(i).y
    
    
        
    If (VBJet.y > EnemyJet(i).y - 40 And VBJet.y < EnemyJet(i).y + 30) And (VBJet.x > EnemyJet(i).x And VBJet.x < EnemyJet(i).x + 100) Then
        
      VBJet.Power = VBJet.Power + 1
      'If VBJet.power > 1 Then MediaPlayer2.Play
      Picture1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
      
      SetEn i

      
      Select Case VBJet.Power
      
      Case 2
      Image1.Picture = LoadPicture(App.Path & "\data\fuel50.gif")
    
      Case 3
      Image1.Picture = LoadPicture(App.Path & "\data\fuel20.gif")
      
      Case 4
      Image1.Picture = LoadPicture(App.Path & "\data\game-over.gif")
      
      End Select
      
      If VBJet.Power = 4 Then
          
          MsgBox "Game Over", vbCritical, "Shooter"
          Unload Me
          Form2.Show
      
      End If
    
    End If
    
    
Next i
End Sub

' Fires missle
Private Sub FireIt()
    'MediaPlayer1.Play
    VBJetMissile.Visible = True
    
    VBJetMissile.x = shooter.Left + 100
    VBJetMissile.y = shooter.Top + 50
    fire.Visible = True
End Sub



' Set enemy postion
Public Sub SetEn(index As Integer)
    
    
    EnemyJet(index).y = Int(Rnd * Me.ScaleHeight) - 100
    EnemyJet(index).x = Int(Rnd * Me.ScaleHeight) + Me.ScaleWidth
    en(index).Left = EnemyJet(index).x
    en(index).Top = EnemyJet(index).y
    
    

End Sub

Private Sub Timer3_Timer()
    Dim i As Integer
    For i = 1 To EnemyJetCount
        EnemyJet(i).speed = EnemyJet(i).speed + 5
        EnemyJetMVar = EnemyJetMVar + 5
    Next i
End Sub
