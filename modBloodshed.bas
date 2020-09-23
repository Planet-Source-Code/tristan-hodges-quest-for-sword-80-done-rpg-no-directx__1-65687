Attribute VB_Name = "modBloodshed"
Public Bleeding As Boolean
Public BleedInt As Integer
Public BleedTime As Integer
Public eBleeding As Boolean
Public eBleedInt As Integer
Public eBleedTime As Integer
'enemy
Public bEnemy As String
Public eStr As Integer
Public eEnd As Integer
Public eAgi As Integer
Public eSt As Integer
Public eStM As Integer
Public eDrop As String
Public eDropChance As Integer
Public eExpGive As Integer
Public eBlood As Integer
Public eAttacking As Boolean
Public eDelay As Integer
' system
Public CanAttack As Boolean
Public Attacking As Boolean
Public delayC As Integer
Public GotoBattle As Boolean
Public ItemWait As Boolean
Public Dazed As Boolean
Public DazedTime As Integer
' dir-keys
Public acAmt As Integer
Public acInputted As Integer
Public acDir(7) As String

Public Sub Bleed(Intensity As Integer)
Bleeding = True
BleedInt = Intensity
BleedTime = Intensity * 12 + Int(Rnd * (Intensity * 5))
End Sub
Public Sub ProcEnemy(sEnemy As String)
Select Case sEnemy
Case "Imp"
    SetE 6, 4, 8, 10, "Lace Wrappings", 80
Case "Wolf"
    SetE 6, 6, 5, 18, "Steel Chunk", 70
Case "Mutant Crab"
    SetE 10, 8, 4, 22, "Crab Shelling", 90
Case "Lesser Eye"
    SetE 10, 6, 10, 25, "Potion of Energy", 60
Case "Asp"
    SetE 10, 7, 20, 20, "Hardened Leather", 70
Case "Cave Worm"
    SetE 9, 13, 5, 20, "Ruby", 60
Case "Badman"
    SetE 14, 13, 10, 30, "Old Chainmail", 40
Case "Greater Imp"
    SetE 12, 12, 17, 30, "Ruby", 50
Case "Red Amaga"
    SetE 14, 13, 10, 30, "Packet of Blood", 30
Case "Sahag"
    SetE 12, 12, 18, 25, "Wind Cloth", 65
Case "Geist"
    SetE 10, 16, 15, 25, "Ruby", 65
Case "Ogre"
    SetE 16, 18, 5, 40, "Steel Chunk", 90
Case "Evilman"
    SetE 16, 13, 10, 35, "Broadsword", 20
Case "Image"
    SetE 10, 15, 16, 30, "Diamond", 35
Case "Pirate"
    SetE 16, 14, 13, 35, "Pirate's Buckler", 25
Case "Ankylo"
    SetE 15, 20, 5, 35, "Crab Shelling", 100
Case "Desert Bull"
    SetE 19, 14, 9, 40, "Damascus Steel Chunk", 35
Case "Earth Image"
    SetE 25, 5, 15, 40, "Wind Cloth", 60
Case "Toggin"
    SetE 27, 35, 15, 100, "Murasame", 40
Case "Follower of Toggin"
    SetE 11, 11, 10, 35, "Damascus Steel Chunk", 25
Case "Pede"
    SetE 19, 16, 20, 40, "Charcoal", 90
Case "Cobra"
    SetE 22, 10, 25, 45, "Leather", 90
Case "Sand Hydra"
    SetE 20, 20, 17, 40, "Diamond", 30
Case "Eye"
    SetE 25, 18, 25, 45, "Quality Potion of Energy", 35
Case "Tiamaunt"
    SetE 20, 40, 10, 75, "Tiamaunt Fang", 50
End Select
End Sub
Public Sub SetE(seStr As Integer, seEnd As Integer, seAgi As Integer, seExpGive As Integer, seDrop As String, Optional seDropChance As Integer = 25)
eStr = seStr
eEnd = seEnd
eAgi = seAgi
eSt = eEnd * StMult
eStM = eEnd * StMult
eExpGive = seExpGive
eDrop = seDrop
eDropChance = seDropChance
eBlood = 100
End Sub
Public Sub CreateDirs()
Randomize
For I = 0 To acAmt - 1
    Select Case Int(Rnd * 4)
    Case 0
        acDir(I) = "l"
    Case 1
        acDir(I) = "u"
    Case 2
        acDir(I) = "r"
    Case 3
        acDir(I) = "d"
    End Select
Next I
End Sub
Public Sub ProcDirPic()
With FrmBattle
    For I = 0 To acAmt - 1
        Select Case acDir(I)
        Case "l"
            If I - 1 < acInputted Then
                .Dirs(I).Picture = .sdLy.Picture
            Else
                .Dirs(I).Picture = .sdL.Picture
            End If
        Case "u"
            If I - 1 < acInputted Then
                .Dirs(I).Picture = .sdUy.Picture
            Else
                .Dirs(I).Picture = .sdU.Picture
            End If
        Case "r"
            If I - 1 < acInputted Then
                .Dirs(I).Picture = .sdRy.Picture
            Else
                .Dirs(I).Picture = .sdR.Picture
            End If
        Case "d"
            If I - 1 < acInputted Then
                .Dirs(I).Picture = .sdDy.Picture
            Else
                .Dirs(I).Picture = .sdD.Picture
            End If
        End Select
    Next I
End With
End Sub
Public Sub Attack()
Dim eStP As Integer
Dim ColStr As Integer
Dim ColEnd As Integer
Dim ColAgi As Integer
Dim tmplol As Integer
ColStr = Stre + bStr + bArmorStr
ColEnd = Endu + bEnd + bArmorEnd + bShEnd
ColAgi = Agil + bAgi + bBootsAgi + bShAgi

St = St - 1
If St < 0 Then St = 0
eStP = (eSt / eStM) * 100
If Int(Rnd * 4) = 0 Then
    eDull = eDull - 1
    For I = 0 To 250
        If Items(I) = eWeap Then
            ItemDull(I) = ItemDull(I) - 1
        End If
    Next I
End If
'first judge if the enemy blocks or dodges, or gets cut
If eEnd > eAgi Then
    If eStP > 49 Then
        If Int(Rnd * (eEnd / 2)) <> 1 Then
            abt ">> The Enemy blocked your attack!"
            eSt = eSt - (Int(Rnd * (ColStr / 2.5)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
        If Int(Rnd * 4) = 0 Then
            abt ">> The enemy narrowly blocks your sword"
            eSt = eSt - (Int(Rnd * (ColStr / 2.5)) + (ColStr / 3))
            If eSt < 0 Then eSt = 0
            Exit Sub
        Else
            'enemy bleed
            EnemyBleed Int(Rnd * 2) + 5
            eBlood = eBlood - 5
            abt ">> You grazed the Enemy with your attack!"
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
    End If
    If eStP > 24 And eStP < 50 Then
        If Int(Rnd * (eEnd / 3)) <> 1 Then
            abt ">> The Enemy blocked your attack, but barely!"
            eSt = eSt - (Int(Rnd * (ColStr / 2.5)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
        If Int(Rnd * 4) = 0 Then
            abt ">> The enemy barely evades your sword"
            eSt = eSt - (Int(Rnd * (ColStr / 2.5)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        Else
            'enemy bleed
            EnemyBleed Int(Rnd * 3) + 5
            eBlood = eBlood - 5
            abt ">> You grazed the Enemy with your attack!"
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
    End If
    Dim rofl As Integer
    rofl = Int(Rnd * 5)
    If eStP < 25 Then
        If rofl = 0 Then
            abt ">> The enemy barely blocks your sword"
            eSt = eSt - (Int(Rnd * (ColStr / 2.5)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        ElseIf rofl = 1 Then
            EnemyBleed 10
            eBlood = eBlood - 25
            abt ">> You gave the Enemy a Mortal Wound !"
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
        Else
            EnemyBleed Int(Rnd * 5) + 6
            eBlood = eBlood - 15
            abt ">> You cut the Enemy deeply !"
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
    End If
ElseIf eAgi > eEnd Then
    If eStP > 49 Then
        If Int(Rnd * (eAgi / 2)) <> 1 Then
            abt ">> The Enemy dodges your attack!"
            eSt = eSt - (Int(Rnd * (ColStr / 3)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
        If Int(Rnd * 3) = 0 Then
            abt ">> The enemy narrowly evades your sword"
            eSt = eSt - (Int(Rnd * (ColStr / 3)) + 2)
            If eSt < 0 Then eSt = 0
            Exit Sub
        Else
            'enemy bleed
            EnemyBleed Int(Rnd * 2) + 5
            abt ">> You grazed the Enemy with your attack!"
            eBlood = eBlood - 5
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
    End If
    If eStP > 24 And eStP < 50 Then
        If Int(Rnd * (eAgi / 3)) <> 1 Then
            abt ">> The Enemy dodges your attack!"
            eSt = eSt - (Int(Rnd * (ColStr / 3)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
        If Int(Rnd * 3) = 0 Then
            abt ">> The enemy narrowly evades your sword"
            eSt = eSt - (Int(Rnd * (ColStr / 3)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        Else
            'enemy bleed
            EnemyBleed Int(Rnd * 2) + 5
            abt ">> You grazed the Enemy with your attack!"
            eBlood = eBlood - 5
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
    End If
    tmplol = Int(Rnd * 4)
    If eStP < 25 Then
        If tmplol = 0 Then
            abt ">> The enemy barely dodges your sword"
            eSt = eSt - (Int(Rnd * (ColStr / 3)) + (ColStr / 4))
            If eSt < 0 Then eSt = 0
            Exit Sub
        ElseIf tmplol = 1 Then
            EnemyBleed 10
            eBlood = eBlood - 25
            abt ">> You gave the Enemy a Mortal Wound !"
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
        Else
            EnemyBleed Int(Rnd * 5) + 6
            abt ">> You cut the Enemy deeply !"
            eBlood = eBlood - 15
            eSt = eSt - Int(Rnd * (ColStr / 2.5))
            If eSt < 0 Then eSt = 0
            Exit Sub
        End If
    End If
End If

End Sub
Public Sub abt(sText As String)
If FrmBattle.Text1.Text = "" Then
    FrmBattle.Text1.Text = sText
Else
    FrmBattle.Text1.Text = FrmBattle.Text1.Text & vbCrLf & sText
End If
FrmBattle.Text1.SelStart = Len(FrmBattle.Text1.Text)
End Sub
Public Sub EnemyBleed(Intensity As Integer)
eBleeding = True
If Int(Rnd * eAgi) = 0 Then
    abt ">> The Enemy looks dazed and confused"
    'If eAttacking = False Then
        FrmBattle.Timer4.Enabled = False
        FrmBattle.Timer4.Interval = 10000
        FrmBattle.Timer4.Enabled = True
    'End If
End If
If eDull < 20 Then
    eBleedInt = Intensity - 1
    eBleedTime = Intensity * 8 + Int(Rnd * (Intensity * 5))
ElseIf eDull > 19 And eDull < 30 Then
    eBleedInt = Intensity
    eBleedTime = Intensity * 9 + Int(Rnd * (Intensity * 5))
ElseIf eDull > 29 And eDull < 40 Then
    eBleedInt = Intensity
    eBleedTime = Intensity * 11 + Int(Rnd * (Intensity * 5))
ElseIf eDull > 39 And eDull < 50 Then
    eBleedInt = Intensity
    eBleedTime = Intensity * 13 + Int(Rnd * (Intensity * 6))
ElseIf eDull > 49 And eDull < 70 Then
    eBleedInt = Intensity + 1
    eBleedTime = Intensity * 14 + Int(Rnd * (Intensity * 6))
ElseIf eDull > 69 And eDull < 90 Then
    eBleedInt = Intensity + 2
    eBleedTime = Intensity * 16 + Int(Rnd * (Intensity * 7))
ElseIf eDull > 89 Then
    eBleedInt = Intensity + 3
    eBleedTime = Intensity * 18 + Int(Rnd * (Intensity * 10))
End If
If eBleedInt > 10 Then eBleedInt = 10
End Sub
Public Sub EnemyAttack()
Dim StP As Integer
Dim ColStr As Integer
Dim ColEnd As Integer
Dim ColAgi As Integer
Dim tmplol As Integer
ColStr = Stre + bStr + bArmorStr
ColEnd = Endu + bEnd + bArmorEnd + bShEnd
ColAgi = Agil + bAgi + bBootsAgi + bShAgi

StP = (St / StM) * 100
Dim homega As Integer
If Int(Rnd * 2) = 0 Then
If StP > 49 Then
    If Int(Rnd * (ColAgi / 1.5)) < (eStr / 4) Then
        Bleed Int(Rnd * 2) + 4
        abt ".: You were grazed by the Enemy's attack!  You begin to bleed!"
        RollDazed
        FrmBattle.Chr.Picture = FrmBattle.ChrSlashed.Picture
        St = St - Int(Rnd * (eStr / 2.5))
        If St < 0 Then St = 0
        Exit Sub
    Else
        If ColEnd > ColAgi Then
            abt ".: You block the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 3)) + (eStr / 3))
        Else
            abt ".: You evade the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 4)) + (eStr / 4))
        End If
        Exit Sub
    End If
ElseIf StP < 50 And StP > 24 Then
    If Int(Rnd * (ColAgi / 2)) < (eStr / 5) Then
        Bleed Int(Rnd * 3) + 4
        abt ".: You were grazed by the Enemy!"
        RollDazed
        FrmBattle.Chr.Picture = FrmBattle.ChrSlashed.Picture
        St = St - Int(Rnd * (eStr / 2.5))
        If St < 0 Then St = 0
        Exit Sub
    Else
        If ColEnd > ColAgi Then
            abt ".: You barely block the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 3)) + (eStr / 4))
            FrmBattle.Chr.Picture = FrmBattle.ChrBlock.Picture
        Else
            abt ".: You narrowly dodge the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 4)) + (eStr / 4))
            FrmBattle.Chr.Picture = FrmBattle.ChrDodge.Picture
        End If
        Exit Sub
    End If
ElseIf StP < 25 Then
    homega = Int(Rnd * 5)
    If homega < 3 Then
        If ColEnd > ColAgi Then
            abt ".: You narrowly block the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 3)) + (eStr / 4))
            FrmBattle.Chr.Picture = FrmBattle.ChrBlock.Picture
        Else
            abt ".: You barely evade the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 4)) + (eStr / 4))
            FrmBattle.Chr.Picture = FrmBattle.ChrDodge.Picture
        End If
    ElseIf homega = 3 Then
        abt ".:: The Enemy gave you a Mortal Wound!"
        Bleed 10
        Blood = Blood - 25
        RollDazed
        FrmBattle.Chr.Picture = FrmBattle.ChrSlashed.Picture
        St = St - Int(Rnd * (eStr / 2.5))
        If St < 0 Then St = 0
    ElseIf homega = 4 Then
        abt ".: The Enemy wounds you deeply!"
        RollDazed
        Bleed Int(Rnd * 5) + 6
        Blood = Blood - 15
        FrmBattle.Chr.Picture = FrmBattle.ChrSlashed.Picture
        St = St - Int(Rnd * (eStr / 2.5))
        If St < 0 Then St = 0
    End If
End If
Else
If StP > 49 Then
    If Int(Rnd * (ColEnd / 1.5)) < (eStr / 4) Then
        Bleed Int(Rnd * 2) + 5
        abt ".: You were grazed by the Enemy's attack!  You begin to bleed!"
        St = St - Int(Rnd * (eStr / 2.5))
        RollDazed
        If St < 0 Then St = 0
        Exit Sub
    Else
        If ColEnd > ColAgi Then
            abt ".: You block the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 3)) + (eStr / 3))
            FrmBattle.Chr.Picture = FrmBattle.ChrBlock.Picture
        Else
            abt ".: You evade the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 4)) + (eStr / 4))
            FrmBattle.Chr.Picture = FrmBattle.ChrDodge.Picture
        End If
        Exit Sub
    End If
ElseIf StP < 50 And StP > 24 Then
    If Int(Rnd * (ColEnd / 2)) < (eStr / 5) Then
        Bleed Int(Rnd * 3) + 5
        abt ".: You were grazed lightly by the Enemy!"
        RollDazed
        St = St - Int(Rnd * (eStr / 2.5))
        If St < 0 Then St = 0
        Exit Sub
    Else
        If ColEnd > ColAgi Then
            abt ".: You barely block the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 3)) + (eStr / 4))
            FrmBattle.Chr.Picture = FrmBattle.ChrBlock.Picture
        Else
            abt ".: You narrowly dodge the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 4)) + (eStr / 4))
            FrmBattle.Chr.Picture = FrmBattle.ChrDodge.Picture
        End If
        Exit Sub
    End If
ElseIf StP < 25 Then
    homega = Int(Rnd * 5)
    If homega < 3 Then
        If ColEnd > ColAgi Then
            abt ".: You narrowly block the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 3)) + (eStr / 4))
        Else
            abt ".: You barely evade the enemy's attack!"
            St = St - (Int(Rnd * (eStr / 4)) + (eStr / 4))
        End If
    ElseIf homega = 3 Then
        abt ".:: The Enemy gave you a Mortal Wound!"
        Bleed 10
        Blood = Blood - 25
        RollDazed
        St = St - Int(Rnd * (eStr / 2.5))
        If St < 0 Then St = 0
    ElseIf homega = 4 Then
        abt ".: The Enemy wounds you deeply!"
        Bleed Int(Rnd * 5) + 6
        Blood = Blood - 10
        RollDazed
        St = St - Int(Rnd * (eStr / 2.5))
        If St < 0 Then St = 0
    End If
End If
End If
End Sub
Public Sub DropsAndExp()
Randomize
If Int(Rnd * (120 - TrFinding)) + 1 < eDropChance Then
    ParseTreasure eDrop
    a ".: On the Enemy you found: " & eDrop
    pSound "itemgain"
    HarmonyStopMusic
    Do: DoEvents
    Loop Until isPlaying = False
    HarmonyPlayMusic mMusic
Else
    a ".: You didn't find anything of value on the Enemy."
End If
a ".: You gain " & eExpGive & " experience."
cExp = cExp + eExpGive
ProcessExp
End Sub
Public Sub RollDazed()
Dim ColAgi As Integer
ColAgi = Agil + bAgi + bBootsAgi + bShAgi
If Int(Rnd * ColAgi) < (eAgi / 1.5) Then
    DazedTime = Int(Rnd * ((eAgi * 3) - ColAgi)) + (eStr * 3)
    If DazedTime < 20 Then DazedTime = 20
    Dazed = True
    If bEnemy = "Asp" Or bEnemy = "Cobra" Then
        abt ".:: The venomous poison makes you feel ill"
    ElseIf bEnemy = "Ogre" Or bEnemy = "Toggin" Then
        abt ".:: The enemy's blunt weapon bops you into a dazed and confused state."
    Else
        abt ".:: The enemy's attack made you feel dazed and confused!"
    End If
End If
End Sub
