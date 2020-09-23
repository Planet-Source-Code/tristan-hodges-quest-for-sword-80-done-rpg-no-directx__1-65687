VERSION 5.00
Begin VB.Form FrmChar 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Character"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Begin"
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   3255
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF8080&
         Caption         =   "You excelled at weapon sharpening"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   3255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF8080&
         Caption         =   "You saved up a large sum of Gold"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF8080&
         Caption         =   "You could run faster than everyone else"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF8080&
         Caption         =   "You had a sword which made you feel more energetic every time you held it"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   3255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF8080&
         Caption         =   "You always salvaged the most treasure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   3255
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FF8080&
         Caption         =   "You always had a stock of potions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1560
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF8080&
         Caption         =   "Strength"
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Agility"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Endurance"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   1
      Text            =   "Rav"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   240
      X2              =   3480
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "From Childhood:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   240
      X2              =   3480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Growing up, you always favored:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   240
      X2              =   3480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "FrmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
pSound "select"
If Text1.Text = "" Then
    MsgBox "The nameless hero, I doubt it."
    Exit Sub
End If
If Text1.Text = "Nya" Then
    Stre = 30
    Agil = 30
    Endu = 30
    Gold = 5000
    TrFinding = 20
    Sharpening = 50
End If
chrName = Text1.Text
Randomize
If Option1.Value = True Then
    Endu = Endu + 3 + Int(Rnd * 3) + 1
    Stre = Stre + Int(Rnd * 3)
End If
If Option2.Value = True Then
    Agil = Agil + 3 + Int(Rnd * 3) + 1
    Stre = Stre + Int(Rnd * 3)
End If
If Option3.Value = True Then
    Stre = Stre + 5 + Int(Rnd * 3) + 1
End If
If Option6.Value = True Then
    Sharpening = Sharpening + 10
End If
If Option5.Value = True Then
    Gold = Gold + 250
End If
If Option4.Value = True Then
    mSpd = 1.6
End If
If Option7.Value = True Then
    PT "Old Shortsword of Energy"
End If
If Option8.Value = True Then
    TrFinding = TrFinding + 5
End If
If Option9.Value = True Then
    PT "Potion of Energy"
    PT "Potion of Energy"
    PT "Potion of Energy"
    PT "Large Bandage"
    PT "Large Bandage"
    PT "Weak Potion of Energy"
End If
Blood = 100
cLevel = 1
If chrName = "Dev" Then
    cLevel = 4
    cExp = 25
    Gold = 210
    PT "Chain-Reinforced Leather"
    PT "Longsword of Energy"
    PT "Windwalker Boots"
    PT "Potion of Energy"
    PT "Large Bandage"
    PT "Fine Sharpening Stone"
    PT "Weak Potion of Energy"
    Stre = 21
    Endu = 20
    Agil = 15
End If
cExp = 0
Unload Me
FrmMain.Enabled = True
FrmMain.Timer1.Enabled = True
FrmMain.Timer2.Enabled = True
FrmMain.SetFocus
FrmMain.txtSel.SetFocus
St = Endu * StMult
End Sub

Private Sub Form_Load()
MsgBox "Welcome to the game.  This is just a little quick tutorial to get you started, as a lot of the aspects of this game you are probably not familiar with.  But first, a little fundamentals.  Strength determines how often you'll knock an enemy stupid (dazed and confused) and how much stamina you'll take from them when they block or dode your attack.  More about the battle system later.  Agility defends you against a dazed and confused status, and it controls how fast you attack and how much you dodge.  Endurance directly translates into your stamina.  Blocking is also controlled by your endurance.  Choose which type of character you want to be wisely."
MsgBox "You get to choose a childhood bonus in this next character creation screen.  There's a couple interesting things there.  First, weapon sharpening.  It's a skill which allows you to sharpen your own weapons.  Bluntly (no phun intended), it determines how bad you cut and enemy--how bad they bleed.  More on that later.  Also interesting, run speed.  Move speed is increased by choosing that bonus, and by equipping high quality boots.  The best boots in the game, Boots of Blinking, makes you move almost like teleportation.  A thing about Swords.   Rarer swords have the suffix ''of Energy'' or the prefix ''Genki'' (Japanese for energetic).  These weapons restore almost double your normal Stamina regen per 5 seconds ingame.  Very handy.  You may also notice, treasure salvaging.  It's a skill which determines how often you find treasure on enemies.  More treasure, the richer you are (no monster in the game drops gold directly).  You should know about the game a bit more, so here we go!"
Stre = 8
Endu = 8
Agil = 8
Gold = 160
TrFinding = 12
Sharpening = 16
Specials = 5
Blood = 100
End Sub

Private Sub Option1_Click()
pSound "click"
End Sub

Private Sub Option2_Click()
pSound "click"
End Sub

Private Sub Option3_Click()
pSound "click"
End Sub

Private Sub Option4_Click()
pSound "click"
End Sub

Private Sub Option5_Click()
pSound "click"
End Sub

Private Sub Option6_Click()
pSound "click"
End Sub

Private Sub Option7_Click()
pSound "click"
End Sub

Private Sub Option8_Click()
pSound "click"
End Sub

Private Sub Option9_Click()
pSound "click"
End Sub
