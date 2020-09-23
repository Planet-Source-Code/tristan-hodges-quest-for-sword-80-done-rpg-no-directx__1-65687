VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest for Sword"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":0000
   ScaleHeight     =   658
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   964
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   1680
      Top             =   360
   End
   Begin VB.TextBox txtSel 
      Height          =   315
      Left            =   11520
      TabIndex        =   9
      Top             =   6840
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   8640
      Width           =   10575
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   4080
   End
   Begin VB.Shape CurStB 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   1500
      Left            =   10020
      Top             =   6975
      Width           =   165
   End
   Begin VB.Label lGold 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   8100
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gold:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   8100
      Width           =   615
   End
   Begin VB.Label lBlood 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   10200
      TabIndex        =   5
      Top             =   6720
      Width           =   375
   End
   Begin VB.Shape sBlood 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   1500
      Left            =   10260
      Top             =   6975
      Width           =   165
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      Height          =   1560
      Left            =   10230
      Top             =   6930
      Width           =   225
   End
   Begin VB.Label lName 
      BackStyle       =   0  'Transparent
      Caption         =   "Lv. 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   8100
      Width           =   855
   End
   Begin VB.Label lStats 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99/99/99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Str/End/Agi:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lStamina 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stamina:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   8400
      Width           =   855
   End
   Begin VB.Shape Col 
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   11400
      Top             =   7200
      Width           =   615
   End
   Begin VB.Image Chr 
      Height          =   480
      Left            =   4680
      Picture         =   "FrmMain.frx":129042
      Top             =   6480
      Width           =   480
   End
   Begin VB.Image iNPC 
      Height          =   480
      Index           =   0
      Left            =   10920
      Picture         =   "FrmMain.frx":129138
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Treas 
      Height          =   480
      Index           =   0
      Left            =   11160
      Picture         =   "FrmMain.frx":129253
      Top             =   5640
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Left            =   -30
      Top             =   8340
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Left            =   -30
      Top             =   8040
      Width           =   2025
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      Height          =   1560
      Left            =   9990
      Top             =   6930
      Width           =   225
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chr_Click()
Bleed 7
a ".: A pointer came out of nowhere and stabbed you!"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
txtSel_KeyUp KeyCode, Shift
End Sub

Private Sub Form_Load()
Me.Width = 10650
If CheckFile(App.Path & "\Harmony.dll") = False Then
    MsgBox "Harmony.dll is required!  It is in the app's location labeled 'harmony.renameme'.  You must rename it to harmony.dll for this game to work!"
    End
End If
mSpd = 1.2
CanMove = True
AnimDir = "d"
Me.Show
Me.Enabled = False
FrmChar.Show
a ".: Quest for Sword, v1.0 - Tristan Hodges, 2006"
a ".: Esc key brings up the menu.  There's a help section there, along with more detailed character statistics and information."
For I = 1 To 150
    Load Col(I)
    Col(I).Visible = True
    Col(I).BackStyle = 1
Next I
For I = 1 To 150
    Load iNPC(I)
    iNPC(I).Visible = True
Next I
For I = 1 To 150
    Load Treas(I)
    Treas(I).Visible = True
Next I
HarmonyCreate
HarmonyInitMidi
Initialize Me.hwnd
NewGame
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 1 Then
    Me.Caption = X & "/" & Y
End If
End Sub

Private Sub Form_Terminate()
HarmonyRelease
HarmonyTermMidi
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
HarmonyRelease
HarmonyTermMidi
End
End Sub

Private Sub lGold_Click()
'PT InputBox("atomic materializer, item name please (yes they existed back then :o")
End Sub

Private Sub lName_Click()
'LoadMap InputBox("hai~~! Teleportation service, hai~!")
End Sub

Private Sub lStamina_Click()
'Stre = InputBox("LOL!")
'Endu = InputBox("LOL.")
'Agil = InputBox("LOL?")
End Sub

Private Sub lStats_Click()
'AfterBattleWait = 500
End Sub

Private Sub Text1_GotFocus()
If Me.Enabled = True Then
    txtSel.SetFocus
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
txtSel_KeyUp KeyCode, Shift
End Sub

Private Sub Timer1_Timer()
If CanMove = True Then
    Character
End If
End Sub

Private Sub Timer2_Timer()
If CanMove = True And IsMoving = True Then
    AnimPos = AnimPos + 1
    If AnimPos = 4 Then AnimPos = 0
    RendChar
    AfterBattleWait = AfterBattleWait - 1
End If
If IsMoving = False Then
    AnimPos = 1
    RendChar
End If
If Blood > -1 Then
    lBlood.Caption = Blood
    sBlood.Top = 465 + (100 - Blood)
    sBlood.Height = Blood
End If
If Blood < 1 Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Bleeding = False
    BleedInt = 0
    BleedTime = 0
    FrmMain.Enabled = False
    FrmDead.Show
    FrmDead.SetFocus
End If
StM = (Endu + (bShEnd / 2) + (bArmorEnd / 2) + (bEnd / 2)) * StMult
lName = "Lv. " & cLevel & " " & chrName
lStamina = St & "/" & StM
lStats = Stre + bStr + bArmorStr & "/" & Endu + bEnd + bArmorEnd + bShEnd & "/" & Agil + bAgi + bArmorAgi + bBootsAgi + bShAgi
lGold = Gold
If Me.Enabled = True Then
    'txtSel.SetFocus
End If

If St < 0 Then St = 0
If St > StM Then St = StM
CurStB.Height = (St / StM) * 100
CurStB.Top = 465 + (100 - ((St / StM) * 100))
MapRun

'bleeding
'bleeding
If Bleeding = True Then
    If Int(Rnd * (12 - BleedInt)) = 1 Then
        Blood = Blood - 1
    End If
    BleedTime = BleedTime - 1
    If BleedTime < 1 Then
        Bleeding = False
        a ">> You stop bleeding"
    End If
End If
'/bleeding
'/bleeding
End Sub

Private Sub Timer3_Timer()
If Right(eWeap, 6) = "Energy" Or Len(eWeap) = 19 Then
    St = St + (3 + (cLevel / 2))
End If
St = St + (5 + (cLevel / 1.5))
If St > StM Then St = StM
End Sub

Private Sub txtSel_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    NPCSelCheck
    If tIndex > -1 Then
        If tState(tIndex) = 0 Then
            GetTreasure
        ElseIf tState(tIndex) = 1 Then
            a ">> This Treasure Box is empty!"
        End If
    End If
End If
If KeyCode = vbKeyEscape Then
    FrmControl.Show
    FrmMain.Enabled = False
End If
If Me.Enabled = True Then
    If KeyCode = vbKeyI Then
        FrmItem.Show
        Me.Enabled = False
    End If
End If
If KeyCode = vbKeyE Then
    FrmEquip.Show
    Me.Enabled = False
End If
End Sub

