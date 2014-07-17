VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options - Shooter"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Min 0 "
      Height          =   495
      Left            =   2220
      TabIndex        =   5
      Top             =   780
      Width           =   3195
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Speed"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim s As Integer

Private Sub Command1_Click()
    

    s = Val(Text1.Text)
    If s < 0 Or s > 30 Then
    Call Er
    Exit Sub
    End If
    
    
    speed = Text1.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Text1.Text = CStr(Module1.speed)
    Label3.Caption = "Minimum Speed is 0." & vbCrLf & "Maximum Speed is 30."
    
End Sub




Private Sub Er()
MsgBox "Enter any number between 0 and 30."
s = 10
End Sub
