VERSION 5.00
Begin VB.Form FrmControl 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
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
   ScaleHeight     =   750
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Help System"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "More Stats"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FrmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
pSound "error"
FrmMain.Enabled = True
FrmMain.SetFocus
Unload Me
End Sub

Private Sub Label1_Click()
pSound "click"
MsgBox chrName & vbCrLf & vbCrLf & " Weapon Sharpening: " & Sharpening & vbCrLf & " Treasure Salvaging: " & TrFinding
End Sub

Private Sub Label2_Click()
pSound "click"
If MsgBox("Do you want to learn about The Basics?", vbYesNo) = vbYes Then
    MsgBox "Moving your character is done by pressing the directional keys.  Depending on which starting bonus you chose, you might run quickily, or you might run at normal speeds.  The world is 6 by 5, meaning 30 maps total, not including dungeon maps.  It is important to find towns along your journey and use their expertise to make your character more powerful.  This includes crafting weapons, buying supplies like sharpening stones and stamina potions, buyin armor, resting, etc.  The bottom left of your screen has the basic stats, and on the right side of the screen there is a red bar.  This is your blood, a rating between 0 and 100 (unless an enchantment deems otherwise).  Townsfolk will point you in the direction of cool locations if you're having trouble finding them."
End If
If MsgBox("Do you want to learn about Towns and Services?", vbYesNo) = vbYes Then
    MsgBox "Towns provide: Inns for replenishing ALL of your stats, Armor shops, Weapon shops which allow you to craft weapons, buy them (though it costs a lot more than crafting does, but crafting requires you find the materials).  There are item shops which provide potions that replenish stamina, or even boost your stats for a period of time, even perminately.  Item shops also have sharpening stones which range in quality, which allow you to sharpen your weapons."
End If
If MsgBox("Do you want to learn about Weapon Sharpness/Dullness?", vbYesNo) = vbYes Then
    MsgBox "The game has 8 ratings for weapons.  Blunt, Very Dull, Dull, Medium, Mildly Sharp, Sharp, Very Sharp, Razor Sharp.  Obviously, the sharper the weapon is, the more damage it does.  Item shops sell various sharpening stones that can be used a random amount of time (it breaks 1 out of 3 times you use it to sharpen it, in which you have to buy a new one).  In other words, if you were lucky you could sharpen your weapon with the same sharpening stone 100 times, but you could also break it after 1 sharpening.  There are 5 different sharpening stones.  Pocket Sharpening Stone, Poor Sharpening Stone, Fine Sharpening Stone, Quality Sharpening Stone, Exquisite Sharpening Stone.  The better the quality, the more sharp you can get your sword.  Only with a high sharpening skill and an exquisite sharpening stone can you achieve Razor Sharp blade status."
End If
If MsgBox("Do you want to learn about Treasure Salvaging stats?", vbYesNo) = vbYes Then
    MsgBox "Collecting materials is important in this world.  No monster in the game except the obvious (humans, goblins) give gold as a result of winning a battle.  However, they often have things to salvage from them.  For instance, a zombie would have zombie fangs, a crab would have crab pincers, an imp might have imp skin, which is a common ingrediant in sword handles.  The higher this skill is, the more often you'll find these materials.  The higher this skill is, the more money you'll have, too.  Simple."
End If
If MsgBox("Do you want to learn about the Battle System?", vbYesNo) = vbYes Then
    MsgBox "The game was designed from day one to be a challange.  It isn't suprising if you had a hard time with your first three or four battles.  Timing is important.  YOUR skill, not your stats, will win you battles.  Okay, so the battle starts and you're greeted by a Wormling or Imp, or whatever it is.  It's attacking pretty fast, and you don't know quite what to do.  If you notice, above your character's head are directional characters.  There should be 4 of them at a time.  Enter them correctly as fast as you can and your character will attack just as you press the last directional key.  If you mess up, a penalty is issued and you have to press one more directional key to complete another attack.  For example, if you mess up 2 times, to complete attacks for the whole rest of the battle, you must input 6 directional keys instead of the beginning 4.  Choose a comfortable speed so you don't mess up, some can be tricky!"
    MsgBox "The item menu is accessed by pressing the spacebar.  The battle is paused by doing this, and you're able to use any of the potions and drinks you have.  The more your skill in Specials is, the more you will randomly perform succession attacks which do extra damage.  Think of them as critical hits."
    MsgBox "The Battle mechanics don't use hit points, or HP.  Instead, incorporated into the very foundation of the game is a Stamina/Blood level relationship.  Let us create an example.  A roaming Ronin decides to kill you.  The battle is evenly matched.  He attacks you and you block the attack.  Your stamina goes down a little bit.  He follows up the attack and you manage to block it again, your stamina goes down and is now dropping below half.  You slash at him and he evades your attack, his stamina will go down as it takes him energy to narrowly evade your sword.  He counters and it grazes your body.  You begin to bleed, very slightly.  You mess up your attack and he gets to try and slash again.  This time, he cuts you.  Not lethally, but it causes you to bleed more.  Now your Blood meter goes down each second, and won't stop until you're out of battle and you dress the wound or it stops on its own.  In short, the lower your stamina is, the less you can block or evade attacks."
    MsgBox "When your stamina gets below 25%, you're risking the enemy's mortal blow which will kill you instantly.  This is the same for the enemy.  The goal is to take away the enemy's stamina and deal it a mortal blow, or give it enough cuts to where it bleeds to death.  Agility controls evading and parrying, Endurance controls blocking and stamina, and Strength determines how much stamina you take away from the enemy when they block or dodge, and how bad the cut you make when you connect with them is."
End If
End Sub

Private Sub Label3_Click()
pSound "click"
MsgBox "Quest for Sword (2006)" & vbCrLf & vbCrLf & "  Programming: Tristan Hodges" & vbCrLf & "  Graphics: Final Fantasy 1 (a little by me)" & vbCrLf & "  Music: Tristan Hodges (Battle theme: Nobuo Uematsu)"
End Sub
