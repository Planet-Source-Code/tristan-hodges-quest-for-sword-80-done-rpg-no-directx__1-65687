VERSION 5.00
Begin VB.Form FrmBattle 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest for Sword - Battle"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
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
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer6 
      Interval        =   50
      Left            =   720
      Top             =   840
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1740
      Begin VB.Image Dirs 
         Height          =   255
         Index           =   6
         Left            =   1440
         Picture         =   "FrmBattle.frx":0000
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Dirs 
         Height          =   255
         Index           =   5
         Left            =   1200
         Picture         =   "FrmBattle.frx":03B6
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Dirs 
         Height          =   255
         Index           =   4
         Left            =   960
         Picture         =   "FrmBattle.frx":076C
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Dirs 
         Height          =   255
         Index           =   3
         Left            =   720
         Picture         =   "FrmBattle.frx":0B22
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Dirs 
         Height          =   255
         Index           =   2
         Left            =   480
         Picture         =   "FrmBattle.frx":0ED8
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Dirs 
         Height          =   255
         Index           =   1
         Left            =   240
         Picture         =   "FrmBattle.frx":128E
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Dirs 
         Height          =   255
         Index           =   0
         Left            =   0
         Picture         =   "FrmBattle.frx":1644
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6600
      TabIndex        =   1
      Top             =   3960
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   15000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Image ChrConfu 
      Height          =   840
      Left            =   8280
      Picture         =   "FrmBattle.frx":19FA
      Top             =   3240
      Width           =   780
   End
   Begin VB.Image sdDy 
      Height          =   255
      Left            =   3120
      Picture         =   "FrmBattle.frx":3C5C
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image sdRy 
      Height          =   255
      Left            =   2760
      Picture         =   "FrmBattle.frx":4012
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image sdUy 
      Height          =   255
      Left            =   2400
      Picture         =   "FrmBattle.frx":43C8
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image sdLy 
      Height          =   255
      Left            =   2040
      Picture         =   "FrmBattle.frx":477E
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image sdD 
      Height          =   255
      Left            =   1440
      Picture         =   "FrmBattle.frx":4B34
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image sdR 
      Height          =   255
      Left            =   1080
      Picture         =   "FrmBattle.frx":4EEA
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image sdU 
      Height          =   255
      Left            =   720
      Picture         =   "FrmBattle.frx":52A0
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image sdL 
      Height          =   255
      Left            =   360
      Picture         =   "FrmBattle.frx":5656
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image ChrVictor 
      Height          =   840
      Left            =   7560
      Picture         =   "FrmBattle.frx":5A0C
      Top             =   3240
      Width           =   780
   End
   Begin VB.Image ChrBlock 
      Height          =   840
      Left            =   8280
      Picture         =   "FrmBattle.frx":7C6E
      Top             =   2280
      Width           =   780
   End
   Begin VB.Image ChrDodge 
      Height          =   840
      Left            =   8280
      Picture         =   "FrmBattle.frx":9ED0
      Top             =   1440
      Width           =   780
   End
   Begin VB.Image ChrSlashed 
      Height          =   840
      Left            =   8280
      Picture         =   "FrmBattle.frx":C132
      Top             =   480
      Width           =   780
   End
   Begin VB.Image ChrAtk2 
      Height          =   840
      Left            =   7440
      Picture         =   "FrmBattle.frx":E394
      Top             =   2280
      Width           =   780
   End
   Begin VB.Image ChrAtk1 
      Height          =   840
      Left            =   7440
      Picture         =   "FrmBattle.frx":105F6
      Top             =   1440
      Width           =   780
   End
   Begin VB.Image ChrIdle 
      Height          =   840
      Left            =   7440
      Picture         =   "FrmBattle.frx":12858
      Top             =   480
      Width           =   780
   End
   Begin VB.Image Chr 
      Height          =   840
      Left            =   4320
      Picture         =   "FrmBattle.frx":14ABA
      Top             =   1200
      Width           =   780
   End
   Begin VB.Image Enemy 
      Height          =   1920
      Left            =   1440
      Picture         =   "FrmBattle.frx":16D1C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1470
   End
   Begin VB.Shape eeSt 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   1500
      Left            =   150
      Top             =   645
      Width           =   165
   End
   Begin VB.Shape eeBlood 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   1500
      Left            =   390
      Top             =   645
      Width           =   165
   End
   Begin VB.Shape chrSt 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   1500
      Left            =   6030
      Top             =   645
      Width           =   165
   End
   Begin VB.Shape chrBlood 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   1500
      Left            =   6270
      Top             =   645
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   1005
      Left            =   105
      Top             =   2265
      Width           =   6405
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      Height          =   1560
      Left            =   6240
      Top             =   600
      Width           =   225
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      Height          =   1560
      Left            =   6000
      Top             =   600
      Width           =   225
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      Height          =   1560
      Left            =   120
      Top             =   600
      Width           =   225
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      Height          =   1560
      Left            =   360
      Top             =   600
      Width           =   225
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   240
      Width           =   6735
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "FrmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chr_Click()
HarmonyRelease
HarmonyTermMidi
End
End Sub

Private Sub Form_Load()
sSound
HarmonyStopMusic
If bEnemy = "Toggin" Then
    HarmonyPlayMusic App.Path & "\Sound\Toggin.mid"
Else
    HarmonyPlayMusic App.Path & "\Sound\Battle.mid"
End If
ProcEnemy bEnemy
Dim tH As Integer
Dim tW As Integer
Me.Width = 6735
Me.Height = 3810
FrmMain.Timer1.Enabled = False
FrmMain.Timer2.Enabled = False
FrmMain.Timer3.Enabled = False
Enemy.Picture = LoadPicture(App.Path & "\Img\MON\Lich.gif")
tH = Enemy.Height
tW = Enemy.Width
Enemy.Stretch = False
Enemy.Picture = LoadPicture(App.Path & "\Img\MON\" & bEnemy & ".gif")
Enemy.Stretch = True
Enemy.Top = 8 + (tH - (Enemy.Height * 2))
Enemy.Width = Enemy.Width * 2
Enemy.Height = Enemy.Height * 2
Chr.Picture = ChrIdle.Picture

Me.Caption = "Quest for Sword - Battle: " & bEnemy

CanAttack = False
acAmt = 4
acInputted = -1
CreateDirs
If ((Agil + bAgi + bShAgi + bBootsAgi) * 100) > 6000 Then
    Timer3.Interval = 1
Else
    Timer3.Interval = 6000 - (Agil + bAgi + bShAgi + bBootsAgi) * 100
End If
Timer3.Enabled = True

ItemWait = False

Timer4.Interval = 5000 - eAgi * 35
Timer4.Enabled = True

Me.Show
Text2.SetFocus

If FirstBattle = 0 Then
    FirstBattle = 1
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    MsgBox "I see this is your first battle.  There's a few things you should know before you jump right in.  The fighting system is based around a Stamina/Blood level relationship.  Depending on your character, you use stamina to dodge or block attacks, and of course to attack.  Your Endurance (End) directly translates into how much Stamina you have.  If you begin to fatigue, you'll notice that you'll less and less be able to block attacks.  The same goes for the Enemy.  You'll notice right off the bat that you'll graze and be grazed on occasion from attacks.  There is nothing avoiding being grazed in a fight, but how bad it is brings us to our next topic; Sharpness.  Simply put, the sharper your sword is (you can find that out in the Equip screen, by pressing Item Stats), the worse of cut you make to the enemy will be."
    MsgBox "The goal of the battle on your part is, to get the enemy's Stamina down, while keeping yours as high as possible, and then dealing the enemy a fatal blow--or Mortal Wound as it is called in the game.  If the enemy is to get your stamina down low, you may receive mortal wound aswell.  What to do when you get cut?  In battle there are two ways of dealing with this.  The cheapest and best way is to buy a Battle-Ready Bandage.  It'll stop most cuts, save fatal blows.  Another way (more expensive) is using a Packet of Blood.  This infusion is an instant boost to your blood level which can often save you."
    MsgBox "Finally, to attack.  You'll notice in a couple of seconds after you click okay to this, directional keys will appear over your character's head.  These must be pressed on your directional key pad (not numpad), and immediately when followed through with, you will attack.  The faster you are, the faster you attack.  If you happen to accidentaly mess up, you'll notice that another directional key is placed up there for you to input.  This stays throughout the whole battle.  You start out with 4 directional keys to hit per attack, but can get as many as 7.  Watch out for Dazed and Confused, it might be tricky to deal with while inputting attack keys!  With that said, battle away."
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer6.Enabled = True
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
'item screen key goes here!!
'remember that
If KeyCode = vbKeySpace Then
    If ItemWait = False Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        FrmBattle.Enabled = False
        GotoBattle = True
        FrmItem.Show
        FrmItem.SetFocus
    End If
    Exit Sub
End If
If CanAttack = False Then Exit Sub
Dim curkey As String
Dim tmpInp As Integer
tmpInp = acInputted
If KeyCode = vbKeyLeft Then
    If acDir(acInputted + 1) = "l" Then
        acInputted = acInputted + 1
    End If
End If
If KeyCode = vbKeyUp Then
    If acDir(acInputted + 1) = "u" Then
        acInputted = acInputted + 1
    End If
End If
If KeyCode = vbKeyRight Then
    If acDir(acInputted + 1) = "r" Then
        acInputted = acInputted + 1
    End If
End If
If KeyCode = vbKeyDown Then
    If acDir(acInputted + 1) = "d" Then
        acInputted = acInputted + 1
    End If
End If
If tmpInp = acInputted Then
    acAmt = acAmt + 1
    If acAmt = 8 Then acAmt = 7
    Attacking = False
    acInputted = -1
    delayC = 0
    CreateDirs
    Exit Sub
End If
If acInputted = acAmt - 1 Then
    Attacking = True
    CanAttack = False
    acInputted = -1
    delayC = 0
    For I = 0 To 6
        Dirs(I).Visible = False
    Next I
    CreateDirs
End If
End Sub

Private Sub Timer1_Timer()
If Blood < 0 Then
    Blood = 0
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Dazed = False
    DazedTime = 0
    Bleeding = False
    BleedTime = 0
    BleedInt = 0
    For I = 0 To 7
        acDir(I) = ""
    Next I
    CanAttack = False
    Me.Enabled = False
    Unload Me
    FrmDead.Show
    FrmDead.SetFocus
    Exit Sub
End If
If eBlood < 0 Then eBlood = 0
If St < 0 Then St = 0
chrSt.Height = (St / StM) * 100
chrSt.Top = 43 + (100 - ((St / StM) * 100))
chrBlood.Height = Blood
chrBlood.Top = 43 + (100 - Blood)
eeSt.Height = (eSt / eStM) * 100
eeSt.Top = 43 + (100 - ((eSt / eStM) * 100))
eeBlood.Height = eBlood
eeBlood.Top = 43 + (100 - eBlood)
'bleeding
If Bleeding = True Then
    If Int(Rnd * (12 - BleedInt)) = 1 Then
        If Int(Rnd * 3) = 0 Then
            Blood = Blood - 2
        Else
            Blood = Blood - 1
        End If
    End If
    BleedTime = BleedTime - 1
    If BleedTime = 0 Then
        Bleeding = False
        abt ".: You stop bleeding"
    End If
End If
If eBleeding = True Then
    If Int(Rnd * (12 - eBleedInt)) = 1 Then
        If Int(Rnd * 3) = 0 Then
            eBlood = eBlood - 2
        Else
            eBlood = eBlood - 1
        End If
    End If
    eBleedTime = eBleedTime - 1
    If eBlood < 1 Then
        eBleedTime = 0
        eBleeding = False
        FrmMain.Timer1.Enabled = True
        FrmMain.Timer2.Enabled = True
        FrmMain.Timer3.Enabled = True
        FrmMain.Enabled = True
        FrmMain.SetFocus
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        Timer5.Enabled = False
        Dazed = False
        DazedTime = 0
        If bEnemy = "Toggin" Then
            FrmMain.iNPC(0).Left = 1000
            FrmMain.Col(NPCIndex(0)).Left = 1000
            Toggin = True
        End If
        If bEnemy = "Tiamaunt" Then
            Tiamaunt = True
        End If
        AfterBattleWait = 35
        For I = 0 To 7
            acDir(I) = ""
        Next I
        DropsAndExp
        HarmonyStopMusic
        HarmonyPlayMusic App.Path & "\Sound\" & mMusic & ".mid"
        Unload Me
        Exit Sub
    End If
    If eBleedTime = 0 Then
        eBleeding = False
        abt ">> The Enemy's wound stops bleeding"
    End If
End If
'/bleeding

If CanAttack = True Then
    Frame1.Visible = True
    For I = 0 To 6
        Dirs(I).Visible = False
    Next I
    For I = 0 To acAmt - 1
        Dirs(I).Visible = True
        If acInputted <> I Then
            ProcDirPic
        Else
            ProcDirPic
        End If
    Next I
    If acAmt = 4 And Frame1.Left <> 280 Then
        Frame1.Left = 280
    ElseIf acAmt = 5 And Frame1.Left <> 272 Then
        Frame1.Left = 272
    ElseIf acAmt = 6 And Frame1.Left <> 264 Then
        Frame1.Left = 264
    ElseIf acAmt = 7 And Frame1.Left <> 256 Then
        Frame1.Left = 256
    End If
End If

If Attacking = True Then
    delayC = delayC + 1
    If delayC = 1 Then
        Chr.Left = 259
        Chr.Picture = ChrAtk1.Picture
    ElseIf delayC = 2 Then
        Chr.Left = 250
        Chr.Picture = ChrAtk2.Picture
    ElseIf delayC = 4 Then
        Chr.Left = 288
        Chr.Picture = ChrIdle.Picture
        Attacking = False
        Timer3.Enabled = True
        'you attack!
        Attack
    End If
End If
If eAttacking = True Then
    eDelay = eDelay + 1
    If eDelay = 1 Then
        Enemy.Left = 112
    ElseIf eDelay = 2 Then
        Enemy.Left = 128
    ElseIf eDelay = 3 Then
        Enemy.Left = 96
        EnemyAttack
    End If
End If
End Sub

Private Sub Timer2_Timer()
If Right(eWeap, 6) = "Energy" Or Len(eWeap) = 19 Then
    St = St + 1
End If
St = St + 1
If St > StM Then St = StM
End Sub

Private Sub Timer3_Timer()
CanAttack = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
eAttacking = True
eDelay = 0
ItemWait = False
If Timer4.Interval = 10000 Then
    abt ">> The enemy looks less dazed and confused"
End If
Timer4.Interval = 5000 - eAgi * 35
End Sub

Private Sub Timer5_Timer()
If Attacking = False Then
    If Chr.Picture = ChrSlashed.Picture Then Chr.Picture = ChrIdle.Picture
    If Chr.Picture = ChrDodge.Picture Then Chr.Picture = ChrIdle.Picture
    If Chr.Picture = ChrBlock.Picture Then Chr.Picture = ChrIdle.Picture
End If
End Sub

Private Sub Timer6_Timer()
If Dazed = True Then
    DazedTime = DazedTime - 1
    If Attacking = False Then
    If Chr.Picture <> ChrConfu.Picture Then
        Chr.Picture = ChrConfu.Picture
    End If
    End If
    If DazedTime < 0 Then
        If Attacking = False Then
            Chr.Picture = ChrIdle.Picture
        End If
        Dazed = False
        abt ".:: You feel less dazed and confused."
        If acAmt = 4 And Frame1.Left <> 280 Then
            Frame1.Left = 280
        ElseIf acAmt = 5 And Frame1.Left <> 272 Then
            Frame1.Left = 272
        ElseIf acAmt = 6 And Frame1.Left <> 264 Then
            Frame1.Left = 264
        ElseIf acAmt = 7 And Frame1.Left <> 256 Then
            Frame1.Left = 256
        End If
        Frame1.Top = 56
    End If
    If Dazed = True Then
        Select Case Int(Rnd * 3)
        Case 0
            If acAmt = 4 And Frame1.Left <> 280 Then
                Frame1.Left = 280
            ElseIf acAmt = 5 And Frame1.Left <> 272 Then
                Frame1.Left = 272
            ElseIf acAmt = 6 And Frame1.Left <> 264 Then
                Frame1.Left = 264
            ElseIf acAmt = 7 And Frame1.Left <> 256 Then
                Frame1.Left = 256
            End If
            Frame1.Top = 56
            Frame1.Top = Frame1.Top - Int(Rnd * 30)
            Frame1.Left = Frame1.Left + (Int(Rnd * 30) - 18)
        Case 1
            If acAmt = 4 And Frame1.Left <> 280 Then
                Frame1.Left = 280
            ElseIf acAmt = 5 And Frame1.Left <> 272 Then
                Frame1.Left = 272
            ElseIf acAmt = 6 And Frame1.Left <> 264 Then
                Frame1.Left = 264
            ElseIf acAmt = 7 And Frame1.Left <> 256 Then
                Frame1.Left = 256
            End If
            Frame1.Top = 56
            Frame1.Top = Frame1.Top + Int(Rnd * 30)
        Case 2
            If acAmt = 4 And Frame1.Left <> 280 Then
                Frame1.Left = 280
            ElseIf acAmt = 5 And Frame1.Left <> 272 Then
                Frame1.Left = 272
            ElseIf acAmt = 6 And Frame1.Left <> 264 Then
                Frame1.Left = 264
            ElseIf acAmt = 7 And Frame1.Left <> 256 Then
                Frame1.Left = 256
            End If
            Frame1.Top = 56
            Frame1.Top = Frame1.Top - Int(Rnd * 30)
            Frame1.Left = Frame1.Left + (Int(Rnd * 30) - 18)
        End Select
    End If
End If
End Sub
