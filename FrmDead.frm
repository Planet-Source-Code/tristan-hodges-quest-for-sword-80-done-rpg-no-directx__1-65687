VERSION 5.00
Begin VB.Form FrmDead 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Death"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "You are no longer... a Livin' Man... But you can be!"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   3720
      Picture         =   "FrmDead.frx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   1920
      Picture         =   "FrmDead.frx":2262
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click to Ressurect"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   1215
      Left            =   360
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "You got too reckless, you got run through, you lie dead on the ground,  to be eaten by a hungry Grue."
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "FrmDead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
HarmonyStopMusic
pSound "livinman"
WaitGO = False
End Sub

Private Sub Label2_Click()
pSound "click"
Image1.Picture = Image2.Picture
Timer1.Enabled = True
'HarmonyPlayMusic mMusic
End Sub

Private Sub Timer1_Timer()
Blood = 100
St = StM
MsgBox "You have been ressurected, but your body was weakened by death.  -3 to Str/End/Agi"
Stre = Stre - 3
Endu = Endu - 3
Agil = Agil - 3
If Stre < 1 Then
    MsgBox "You're so weak that gravity itself crushes you.  Game over."
    End
End If
If Endu < 1 Then
    MsgBox "You're so exhausted from lifting your hand that you die.  Game over."
    End
End If
If Agil < 1 Then
    MsgBox "You're so clumbsy and far from nimble on your feet that you can't even walk correctly.  Game Over."
    End
End If
Unload Me
FrmMain.Enabled = True
FrmMain.SetFocus
FrmMain.Timer1.Enabled = True
FrmMain.Timer2.Enabled = True
FrmMain.Timer3.Enabled = True
End Sub
