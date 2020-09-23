Attribute VB_Name = "ModMaps"

Public Sub MapGlobals(m As String)
FrmMain.Caption = "Quest for Sword"
Select Case m
Case "Direshore"
    aTr 0, "Battle-Ready Bandage", 63, 449, True
    FrmMain.Caption = "Quest for Sword - Direshore"
Case "1"
    aTr 1, "Exquisite Sharpening Stone", 130, 385, True
    aTr 2, "Quality Potion of Energy", 352, 119, True
    If Int(Rnd * 5) = 0 Then
        aTr 3, "Longsword of Energy", 640, 32, False
    Else
        aTr 3, "Half-Tang Wakizashi", 640, 32, False
    End If
    FrmMain.Caption = "Quest for Sword - Direshore Beach"
Case "6"
    FrmMain.Caption = "Quest for Sword - Direshore Area"
Case "6-cave1"
    FrmMain.Caption = "Quest for Sword - Direshore Caves"
    aTr 4, "Fine Sharpening Stone", 383, 413, True
Case "6-cave2"
    FrmMain.Caption = "Quest for Sword - Direshore Caves"
    aTr 5, "Ruby", 484, 355, True
Case "6-cave3"
    FrmMain.Caption = "Quest for Sword - Direshore Caves"
    aTr 6, "Windwalker Boots", 508, 299
Case "11-Inn"
    aTr 7, "Packet of Blood", 256, 160
    aTr 8, "Diamond", 382, 284
Case "Amdopple"
    FrmMain.Caption = "Quest for Sword - Amdopple"
Case "17"
    FrmMain.Caption = "Quest for Sword - Amdopple Outskirts"
Case "Dead Castle"
    FrmMain.Caption = "Quest for Sword - Dead Castle"
Case "Dead Castle-2"
    FrmMain.Caption = "Quest for Sword - Dead Castle"
    aTr 9, "Battle-Ready Bandage", 290, 116
    aTr 10, "Hellviel Chunk", 126, 160
Case "Pirate Lair-3"
    aTr 11, "Falchion", 225, 200
    aTr 12, "Diamond", 425, 200
Case "28"
    FrmMain.Caption = "Quest for Sword - Toggin's Pass"
    If Toggin = True Then
        FrmMain.iNPC(0).Left = 1000
        FrmMain.Col(NPCIndex(0)).Left = 1000
    End If
Case "27"
    aTr 13, "Exquisite Potion of Energy", 128, 320, True
Case "Oasis"
    aTr 14, "Steve", 316, 30, True
    aTr 15, "Mimic", 61, 63, True
Case "23"
    aTr 16, "Exquisite Sharpening Stone", 300, 243
End Select
End Sub

Public Sub ParseTreasure(sTr As String)
Select Case sTr
' potions
Case "Weak Potion of Energy"
    ai "Weak Potion of Energy", "Potion", 10
Case "Potion of Energy"
    ai "Potion of Energy", "Potion", 30
Case "Quality Potion of Energy"
    ai "Quality Potion of Energy", "Potion", 80
Case "Exquisite Potion of Energy"
    ai "Exquisite Potion of Energy", "Potion", 200
Case "Packet of Blood"
    ai "Packet of Blood", "Potion", 250
'bandages
Case "Small Bandage"
    ai "Small Bandage", "Bandage", 10
Case "Large Bandage"
    ai "Large Bandage", "Bandage", 30
Case "Battle-Ready Bandage"
    ai "Battle-Ready Bandage", "Bandage", 50
'sharpening stones
Case "Pocket Sharpening Stone"
    ai "Pocket Sharpening Stone", "Stone", 5
Case "Poor Sharpening Stone"
    ai "Poor Sharpening Stone", "Stone", 10
Case "Fine Sharpening Stone"
    ai "Fine Sharpening Stone", "Stone", 20
Case "Quality Sharpening Stone"
    ai "Quality Sharpening Stone", "Stone", 40
Case "Exquisite Sharpening Stone"
    ai "Exquisite Sharpening Stone", "Stone", 60
' weapons
Case "Old Shortsword of Energy"
    ai "Old Shortsword of Energy", "Sword", 200, 5, 1, 1, 40
Case "Old Shortsword"
    ai "Old Shortsword", "Sword", 25, 2, , , 30
Case "Shortsword"
    ai "Shortsword", "Sword", 50, 4, 0, -1, 40
Case "Half-Tang Wakizashi"
    ai "Half-Tang Wakizashi", "Sword", 100, 6, 0, 0, 50
Case "Longsword"
    ai "Longsword", "Sword", 125, 8, 0, -1, 40
Case "Longsword of Energy"
    ai "Longsword of Energy", "Sword", 250, 8, 0, -1, 40
Case "Full-Tang Wakizashi"
    ai "Full-Tang Wakizashi", "Sword", 160, 8, , 1, 60
Case "Kodachi of Energy"
    ai "Kodachi of Energy", "Sword", 220, 9, , 3, 55
Case "Broadsword"
    ai "Broadsword", "Sword", 220, 11, 4, -3, 50
Case "Katana"
    ai "Katana", "Sword", 300, 11, 1, 3, 70
Case "Katana of Energy"
    ai "Katana of Energy", "Sword", 500, 11, 1, 3, 70
Case "Falchion"
    ai "Falchion", "Sword", 300, 12, 2, -1, 50
Case "Murasame"
    ai "Murasame", "Sword", 500, 12, 6, -2, 70
Case "Damascus Katana"
    ai "Damascus Katana", "Sword", 600, 14, 1, 5, 70
Case "Tiamaunt Fang"
    ai "Tiamaunt Fang", "Sword", 300, 13, 2, 2, 80
Case "Ken no Tenshi"
    ai "Ken no Tenshi", "Sword", 1500, 16, 2, 4, 80
Case "Genki Ken no Tenshi"
    ai "Genki Ken no Tenshi", "Sword", 2000, 16, 2, 4, 80
Case "Masamune"
    ai "Masamune", "Sword", 2000, 17, 0, 10, 90
Case "Venfer"
    ai "Venfer", "Sword", 2000, 22, 10, -5, 80
Case "Steve"
    ai "Steve", "Sword", 300, 10, 10, 10, 10
'armor
Case "Cotton Clothing"
    ai "Cotton Clothing", "Armor", 5, 0, 1, 1
Case "Leather Armor"
    ai "Leather Armor", "Armor", 15, 0, 2, 1
Case "Hardened Leather Armor"
    ai "Hardened Leather Armor", "Armor", 30, 0, 4, 1
Case "Chain-Reinforced Leather"
    ai "Chain-Reinforced Leather", "Armor", 60, 0, 6, 1
Case "Old Chainmail"
    ai "Old Chainmail", "Armor", 120, 0, 8, 0
Case "Rusted Steel Armor"
    ai "Rusted Steel Armor", "Armor", 160, 0, 10, -1
Case "Chainmail"
    ai "Chainmail", "Armor", 220, 0, 11, 0
Case "Steel Armor"
    ai "Steel Armor", "Armor", 280, 0, 13, -3
Case "Plated Leather"
    ai "Plated Leather", "Armor", 350, 0, 12, 3
Case "Damascus Plated Leather"
    ai "Damascus Plated Leather", "Armor", 460, 0, 14, 5
Case "Platemail"
    ai "Platemail", "Armor", 1000, 0, 21, -5
Case "Helltempered Plated Leather"
    ai "Helltempered Plated Leather", "Armor", 1500, 0, 18, 5
Case "Venfer's Platemail"
    ai "Venfer's Platemail", "Armor", 1600, 0, 24, -5
'shield
Case "Wooden Shield"
    ai "Wooden Shield", "Shield", 10, 0, 2, -1
Case "Iron-Plated Shield"
    ai "Iron-Plated Shield", "Shield", 30, , 4, -2
Case "Steel Shield"
    ai "Steel Shield", "Shield", 60, , 6, -4
Case "Pirate's Buckler"
    ai "Pirate's Buckler", "Shield", 100, 0, 5, -1
Case "Damascus Shield"
    ai "Damascus Shield", "Shield", 300, , 10, -6
Case "Void Shield"
    ai "Void Shield", "Shield", 500, 0, 12, -7
Case "Elemental Shield"
    ai "Elemental Shield", "Shield", 1000, , 7, 0
Case "Venfer's Shield"
    ai "Venfer's Shield", "Shield", 1500, , 15, -8
'boots
Case "Cloth Boots"
    ai "Cloth Boots", "Boots", 10, 0.1, , 4
Case "Leather Boots"
    ai "Leather Boots", "Boots", 30, 0.2, , 6
Case "Windwalker Boots"
    ai "Windwalker Boots", "Boots", 80, 0.2, , 8
Case "Windstalker Boots"
    ai "Windstalker Boots", "Boots", 170, 0.3, , 14
Case "Boots of Stride"
    ai "Boots of Stride", "Boots", 600, 0.3, , 18
Case "Hellspeed Boots"
    ai "Hellspeed Boots", "Boots", 900, 0.4, , 22
Case "Boots of Blinking"
    ai "Boots of Blinking", "Boots", 1200, 2, , 25
Case "Air"
    ai "Air", "Boots", 2, 1
'crafts
Case "Diamond"
    ai "Diamond", "Craft", 110
Case "Ruby"
    ai "Ruby", "Craft", 80
Case "Steel Chunk"
    ai "Steel Chunk", "Craft", 45
Case "Lace Wrappings"
    ai "Lace Wrappings", "Craft", 25
Case "Brass Chunk"
    ai "Brass Chunk", "Craft", 25
Case "Charcoal"
    ai "Charcoal", "Craft", 15
Case "Damascus Steel Chunk"
    ai "Damascus Steel Chunk", "Craft", 70
Case "Hellviel Chunk"
    ai "Hellviel Chunk", "Craft", 80
Case "Leather"
    ai "Leather", "Craft", 20
Case "Hardened Leather"
    ai "Hardened Leather", "Craft", 40
Case "Crab Shelling"
    ai "Crab Shelling", "Craft", 20
Case "Wind Cloth"
    ai "Wind Cloth", "Craft", 40
End Select
End Sub
Public Sub ParseUse(sI As String)
Select Case sI
Case "Weak Potion of Energy"
    ReplStamina 10
Case "Potion of Energy"
    ReplStamina 18
Case "Quality Potion of Energy"
    ReplStamina 30
Case "Exquisite Potion of Energy"
    ReplStamina 60
Case "Packet of Blood"
    Blood = Blood + 20
    If Blood > 100 Then Blood = 100
Case "Small Bandage"
    If Bleeding = False Then
        If GotoBattle = True Then
            abt ".: You're not bleeding!"
            Exit Sub
        Else
            a ".: You're not bleeding!"
            Exit Sub
        End If
    End If
    If GotoBattle = True Then
        abt ".: You cannot dress a wound in battle!"
    Else
        If BleedInt > 7 Then
            a ".: The bandage does no good; the wound is too large!"
        Else
            BleedInt = 0
            BleedTime = 0
            Bleeding = False
            a ".: The bandage fit nicely around the wound and shortly after, it stopped bleeding."
        End If
    End If
Case "Large Bandage"
    If Bleeding = False Then
        If GotoBattle = True Then
            abt ".: You're not bleeding!"
            Exit Sub
        Else
            a ".: You're not bleeding!"
            Exit Sub
        End If
    End If
    If GotoBattle = True Then
        abt ".: You cannot dress a wound in battle!"
        Exit Sub
    Else
        BleedInt = 0
        BleedTime = 0
        Bleeding = False
        a ".: The bandage fit nicely around the wound and shortly after, it stopped bleeding."
    End If
Case "Battle-Ready Bandage"
    If Bleeding = False Then
        If GotoBattle = True Then
            abt ".: You're not bleeding!"
            Exit Sub
        Else
            a ".: You're not bleeding!"
            Exit Sub
        End If
    End If
    If GotoBattle = True Then
        If BleedInt < 10 Then
            abt ".: The Patch fit on nicely and quickily, you stop bleeding."
            BleedInt = 0
            BleedTime = 0
            Bleeding = False
            Exit Sub
        Else
            abt ".: The Patch was too small and didn't stop all the bleeding!"
            BleedInt = 8
            BleedTime = BleedTime - 15
            If BleedTime < 0 Then BleedTime = 0
            Exit Sub
        End If
    Else
        a ".: Don't you think you should save those for battles?"
        If BleedInt < 10 Then
            BleedInt = 0
            BleedTime = 0
            Bleeding = False
            a ".: You stop bleeding"
            Exit Sub
        Else
            BleedInt = 7
            BleedTime = BleedTime - 10
            If BleedTime < 0 Then BleedTime = 0
            a ".: The patch isn't big enough, it only lessens the bleeding."
            Exit Sub
        End If
    End If
End Select
If Right(sI, 5) = "Stone" Then
    If GotoBattle = True Then
        MsgBox "You cannot sharpen weapons in battle, it's far too much of a delicate process!"
        Exit Sub
    End If
    UsingStone = sI
    FrmSharpen.Show
    FrmItem.Enabled = False
End If
End Sub
Public Sub ProcessWShop()
With FrmWShop
Select Case sMap
Case "Direshore Weapon Shop"
    .List1.AddItem "Old Shortsword"
    .List1.AddItem "Shortsword"
    .List1.AddItem "Half-Tang Wakizashi"
    .List1.AddItem "Longsword"
Case "Amdopple Weapon Shop"
    .List1.AddItem "Longsword"
    .List1.AddItem "Broadsword"
    .List1.AddItem "Katana"
Case "Sandstone Weapon Shop"
    .List1.AddItem "Falchion"
    .List1.AddItem "Katana of Energy"
    .List1.AddItem "Murasame"
End Select
End With
End Sub
Public Sub ProcessAShop()
With FrmAShop
Select Case sMap
Case "Direshore Armor Shop"
    .List1.AddItem "Leather Armor"
    .List1.AddItem "Hardened Leather Armor"
    .List1.AddItem "Chain-Reinforced Leather"
    .List1.AddItem "Old Chainmail"
    .List1.AddItem "Wooden Shield"
    .List1.AddItem "Iron-Plated Shield"
    .List1.AddItem "Cloth Boots"
    .List1.AddItem "Leather Boots"
Case "Amdopple Armor Shop"
    .List1.AddItem "Rusted Steel Armor"
    .List1.AddItem "Chainmail"
    .List1.AddItem "Steel Armor"
    .List1.AddItem "Steel Shield"
    .List1.AddItem "Windwalker Boots"
    .List1.AddItem "Windstalker Boots"
Case "Sandstone Armor Shop"
    .List1.AddItem "Plated Leather"
    .List1.AddItem "Damascus Plated Leather"
    .List1.AddItem "Damascus Shield"
    .List1.AddItem "Boots of Stride"
End Select
End With
End Sub
Public Sub EnterZoneGlobals(mFrom As String, mTo As String)
Select Case mFrom
Case "16"
    If mTo = "17" Then
        a ".: You see a town a ways north of here."
    End If
Case "Dead Castle"
    If mTo = "17" Then
        a ".: You see a town north of here."
    End If
Case "17"
    If mTo = "Dead Castle" And cLevel < 5 Then
        a ".: You get a feeling you're not strong enough to be in such a dangerous place..."
    End If
End Select
End Sub

