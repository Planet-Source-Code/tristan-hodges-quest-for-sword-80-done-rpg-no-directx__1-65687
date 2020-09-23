Attribute VB_Name = "ModLife"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'map
Public ColBodies As Integer
Public Collision As Boolean
Public lX As Integer
Public lY As Integer
Public cIndex As Integer
Public mMusic As String
Public mBeforeNC As String
Public AfterBattleWait As Integer
'main
Public CanMove As Boolean
Public Const StMult As Single = 2.5
'chr
Public mSpd As Single
Public AnimPos As Integer
Public AnimDir As String
Public IsMoving As Boolean
Public cLevel As Integer
Public cExp As Integer
'chr stats
Public chrName As String
Public St As Integer
Public StM As Integer
Public Blood As Integer
Public Stre As Integer
Public Endu As Integer
Public Agil As Integer
Public Gold As Integer
Public TrFinding As Single
Public Sharpening As Single
Public Specials As Single
' location
Public sMap As String
Public MonOcur As Integer
Public MonName As String
Public MonName2 As String
Public Cols As Integer

Public Warps As Integer
Public WarpIndex(20) As Integer
Public WarpMap(20) As String
Public WarpX(20) As Integer
Public WarpY(20) As Integer

Public NPCs As Integer
Public NPCIndex(150) As Integer
Public NPCName(150) As String
Public NPCsay(150) As String
Public NPCimg(150) As Integer
Public NPCNear As Boolean
Public NPCNearI As Integer
' treasure
Public t As Integer
Public tState(150) As Integer
Public tItem(150) As String
Public tIndex As Integer
' inventory
Public Items(250) As String
Public ItemType(250) As String
Public ItemStr(250) As Single
Public ItemEnd(250) As Integer
Public ItemAgi(250) As Integer
Public ItemWorth(250) As Integer
Public ItemDull(250) As Integer
' equip
Public bStr As Integer
Public bAgi As Integer
Public bEnd As Integer
Public eWeap As String
Public eArmor As String
Public bArmorStr As Integer
Public bArmorAgi As Integer
Public bArmorEnd As Integer
Public eShield As String
Public bShEnd As Integer
Public bShAgi As Integer
Public eBoots As String
Public bBootsMove As Single
Public bBootsAgi As Integer
Public eDull As Integer
'sharpening
Public UsingStone As String
Public ShPlus As Integer
'shop
Public CurSB As Integer
Public SelCost As Integer
'crafting
Public Crafting As String
'gameover
Public WaitGO As Boolean
'game events
Public Toggin As Boolean
Public Tiamaunt As Boolean
'help system
Public FirstBattle As Integer
Public FirstItemShop As Integer

Public Function GetKey(lngKey As Long) As Boolean
If GetKeyState(lngKey&) < 0 Then
    GetKey = True
Else
    GetKey = False
End If
End Function

Public Sub Character()
Collision = False
IsMoving = False
If GetKey(vbKeyLeft) = True Then
    FrmMain.Chr.Left = FrmMain.Chr.Left - (mSpd + bBootsMove)
    AnimDir = "l"
    IsMoving = True
End If
If GetKey(vbKeyRight) = True Then
    FrmMain.Chr.Left = FrmMain.Chr.Left + (mSpd + bBootsMove)
    AnimDir = "r"
    IsMoving = True
End If
If GetKey(vbKeyUp) = True Then
    FrmMain.Chr.Top = FrmMain.Chr.Top - (mSpd + bBootsMove)
    AnimDir = "u"
    IsMoving = True
End If
If GetKey(vbKeyDown) = True Then
    FrmMain.Chr.Top = FrmMain.Chr.Top + (mSpd + bBootsMove)
    AnimDir = "d"
    IsMoving = True
End If
With FrmMain
For I = 0 To .Col.UBound
    If .Chr.Left + 7 < .Col(I).Left + .Col(I).Width And .Col(I).Left < .Chr.Left + .Chr.Width - 7 Then
        If .Chr.Top + 15 < .Col(I).Top + .Col(I).Height And .Col(I).Top < .Chr.Top + .Chr.Height Then
            Collision = True
            cIndex = I
            WarpCheck
        End If
    End If
Next I
tIndex = -1
If t > -1 Then
For I = 0 To 150
    If .Chr.Left - 10 < .Treas(I).Left + .Treas(I).Width And .Treas(I).Left < .Chr.Left + .Chr.Width + 10 Then
        If .Chr.Top - 10 < .Treas(I).Top + .Treas(I).Height And .Treas(I).Top < .Chr.Top + .Chr.Height + 10 Then
            tIndex = I
        End If
    End If
    If .Chr.Left + 5 < .Treas(I).Left + .Treas(I).Width And .Treas(I).Left < .Chr.Left + .Chr.Width - 5 Then
        If .Chr.Top + 12 < .Treas(I).Top + .Treas(I).Height And .Treas(I).Top < .Chr.Top + .Chr.Height Then
            Collision = True
        End If
    End If
Next I
End If
NPCCheck
If Collision = True Then
If GetKey(vbKeyLeft) = True Then
    .Chr.Left = .Chr.Left + (mSpd + bBootsMove)
End If
If GetKey(vbKeyRight) = True Then
    .Chr.Left = .Chr.Left - (mSpd + bBootsMove)
End If
If GetKey(vbKeyUp) = True Then
    .Chr.Top = .Chr.Top + (mSpd + bBootsMove)
End If
If GetKey(vbKeyDown) = True Then
    .Chr.Top = .Chr.Top - (mSpd + bBootsMove)
End If
End If
End With
End Sub
Public Sub RendChar()
With FrmMain
    .Chr.Picture = LoadPicture(App.Path & "/Img" & "/" & AnimDir & AnimPos & ".gif")
End With
End Sub
Public Sub a(sText As String)
If FrmMain.Text1.Text = "" Then
    FrmMain.Text1 = sText
    Exit Sub
End If
FrmMain.Text1.Text = FrmMain.Text1.Text & vbCrLf & sText
FrmMain.Text1.SelStart = Len(FrmMain.Text1.Text)
End Sub
Public Sub NewGame()
t = -1
If chrName = "Dev" Then

Else
    'LoadMap "Direshore" '
End If
LoadMap "Direshore"
If FrmChar.Visible = True Then
    HarmonyStopMusic
    HarmonyPlayMusic App.Path & "\Sound\Adventure3.mid"
End If
PT "Small Bandage"
PT "Potion of Energy"
PT "Large Bandage"
PT "Packet of Blood"
End Sub
Public Sub PT(treasu As String)
ParseTreasure treasu
End Sub
Public Sub LoadMap(sName As String)
Dim sLine As String
sMap = sName
FrmMain.Picture = LoadPicture(App.Path & "\Maps\" & sName & ".bmp")
'HarmonyStopMusic
For I = 0 To 150
    FrmMain.iNPC(I).Left = 1000
    FrmMain.Treas(I).Left = 1400
Next I
For I = 1 To 150
    Unload FrmMain.Col(I)
    Load FrmMain.Col(I)
    FrmMain.Col(I).Visible = True
    FrmMain.Col(0).Left = 1000
Next I
mBeforeNC = ""
Open App.Path & "\Maps\" & sName & ".map" For Input As #1
    Input #1, MonOcur
    Input #1, MonName
    Input #1, MonName2
    Dim tmpbuf As String
    Input #1, tmpbuf
    If tmpbuf = mMusic Then
        mBeforeNC = "Don't"
    Else
        mMusic = tmpbuf
    End If
    Input #1, Cols
    For I = 0 To Cols
        Input #1, sLine
        FrmMain.Col(I).Left = CInt(sLine)
        Input #1, sLine
        FrmMain.Col(I).Top = CInt(sLine)
        Input #1, sLine
        FrmMain.Col(I).Width = CInt(sLine)
        Input #1, sLine
        FrmMain.Col(I).Height = CInt(sLine)
    Next I
    Input #1, Warps
    If Warps > -1 Then
        For I = 0 To Warps
            Input #1, WarpIndex(I)
            Input #1, WarpMap(I)
            Input #1, WarpX(I)
            Input #1, WarpY(I)
        Next I
    End If
    Input #1, NPCs
    If NPCs > -1 Then
        For I = 0 To NPCs
            Input #1, NPCIndex(I)
            Input #1, NPCName(I)
            Input #1, NPCsay(I)
            Input #1, NPCimg(I)
            FrmMain.iNPC(I).Picture = LoadPicture(App.Path & "\Img\NPC\" & NPCimg(I) & ".gif")
            FrmMain.iNPC(I).Left = FrmMain.Col(NPCIndex(I)).Left
            FrmMain.iNPC(I).Top = FrmMain.Col(NPCIndex(I)).Top
        Next I
    End If
Close #1
MapGlobals sMap
If mBeforeNC <> "Don't" Then
    HarmonyPlayMusic App.Path & "\Sound\" & mMusic & ".mid"
    sSound
End If
End Sub
Public Sub WarpCheck()
If Warps > -1 Then
For an = 0 To Warps
    If WarpIndex(an) = cIndex Then
        If CheckFile(App.Path & "\Maps\" & WarpMap(an) & ".map") = False Then
            MsgBox "This map does not exist.  There is always another route to getting to where you want to go.  Since, of course, the game is not finished."
            WarpIndex(an) = 140
            Exit Sub
        End If
        FrmMain.Chr.Left = WarpX(an)
        FrmMain.Chr.Top = WarpY(an)
        EnterZoneGlobals sMap, WarpMap(an)
        LoadMap WarpMap(an)
    End If
Next an
End If
End Sub
Public Sub NPCCheck()
NPCNear = False
NPCNearI = -1
With FrmMain
For I = 0 To .Col.UBound
    If .Chr.Left - 10 < .Col(I).Left + .Col(I).Width And .Col(I).Left < .Chr.Left + .Chr.Width + 10 Then
        If .Chr.Top - 10 < .Col(I).Top + .Col(I).Height And .Col(I).Top < .Chr.Top + .Chr.Height + 10 Then
            NPCNear = True
            NPCNearI = I
        End If
    End If
Next I
End With
End Sub
Public Sub NPCSelCheck()
'If GetKey(vbKeySpace) = True Then
    If NPCs > -1 Then
        For I = 0 To NPCs
            If NPCNearI = NPCIndex(I) And NPCNear = True Then
                If InStr(1, NPCsay(I), "%") > 0 Then
                    Dim Buff As String
                    Dim lPiece As String
                    Dim rPiece As String
                    For gog = 1 To Len(NPCsay(I)) - 1
                        Buff = Mid(NPCsay(I), gog, 1)
                        lPiece = Left(NPCsay(I), gog - 1)
                        rPiece = Right(NPCsay(I), Len(NPCsay(I)) - gog)
                        If Buff = "%" Then
                            NPCsay(I) = lPiece & "," & rPiece
                        End If
                    Next gog
                End If
                a ".:: " & NPCName(I) & ":  " & NPCsay(I)
                If NPCName(I) = "Item Shopkeeper" Then
                    FrmMain.Enabled = False
                    FrmIShop.Show
                    FrmIShop.SetFocus
                End If
                If NPCName(I) = "Weapon Shopkeeper" Then
                    FrmMain.Enabled = False
                    FrmWShop.Show
                    FrmWShop.SetFocus
                End If
                If NPCName(I) = "Armor Shopkeeper" Then
                    FrmMain.Enabled = False
                    FrmAShop.Show
                    FrmAShop.SetFocus
                End If
                If NPCName(I) = "Weapon Crafter" Then
                    FrmMain.Enabled = False
                    FrmCraft.Show
                    FrmCraft.SetFocus
                End If
                If NPCName(I) = "Innkeeper" Then
                    ProcInn
                End If
                If NPCName(I) = "Toggin" Then
                    'game event
                    'Toggin = True
                    bEnemy = "Toggin"
                    FrmMain.Enabled = False
                    FrmBattle.Show
                    IsMoving = False
                    IsBattle = 0
                    
                    '/game event
                End If
            End If
nya:
        Next I
    End If
'End If
End Sub
Public Sub MapRun()
Dim IsBattle As Integer
Randomize
If MonOcur > 0 Then
    IsBattle = Int(Rnd * (70 - MonOcur))
    If IsMoving = True And IsBattle = 0 And AfterBattleWait < 1 Then
        If Int(Rnd * 2) = 0 Then
            bEnemy = MonName
        Else
            bEnemy = MonName2
        End If
        FrmMain.Enabled = False
        FrmBattle.Show
        IsMoving = False
        IsBattle = 0
    End If
End If
End Sub
Public Sub aTr(TrIndex As Integer, stItem As String, stX As Integer, stY As Integer, Optional Invis As Boolean)
t = t + 1
FrmMain.Treas(TrIndex).Left = stX
FrmMain.Treas(TrIndex).Top = stY
tItem(TrIndex) = stItem
If Invis = True Then FrmMain.Treas(TrIndex).Picture = Nothing
End Sub
Public Sub GetTreasure()
tState(tIndex) = 1
ParseTreasure tItem(tIndex)
a ".: In the Treasure Box you found: " & tItem(tIndex)
If tItem(tIndex) = "Mimic" Then
    If sMap = "Oasis" Then
        bEnemy = "Tiamaunt"
        FrmMain.Enabled = False
        FrmBattle.Show
        IsMoving = False
        IsBattle = 0
    End If
End If
End Sub
Public Function aiEmpty() As Integer
For I = 0 To 250
    If Items(I) = "" Then
        aiEmpty = I
        Exit Function
    End If
Next I
End Function
Public Sub ai(siName As String, siType As String, Optional siWorth As Integer, Optional iiStr As Single, Optional iiEnd As Integer, Optional iiAgi As Integer, Optional iiDull As Integer)
' REMEMBER, for Boots, use its Str for move speed bonus!
Dim touse As Integer
touse = aiEmpty
Items(touse) = siName
ItemType(touse) = siType
ItemStr(touse) = iiStr
ItemEnd(touse) = iiEnd
ItemAgi(touse) = iiAgi
ItemWorth(touse) = siWorth
ItemDull(touse) = iiDull
End Sub
Public Sub aiRemove(siName As String)
For I = 0 To 250
    If Items(I) = siName Then
        Items(I) = ""
        ItemType(I) = ""
        ItemStr(I) = 0
        ItemEnd(I) = 0
        ItemAgi(I) = 0
        ItemWorth(I) = 0
        Exit Sub
    End If
Next I
End Sub
Public Sub ReplStamina(stamAmnt As Integer)
St = St + stamAmnt
If St > StM Then St = StM
End Sub
Public Sub ReplBlood(bloodAmnt As Integer)
Blood = Blood + bloodAmnt
If Blood > 100 Then Blood = 100
End Sub
Public Function CalcSharpness(sharpness As Integer)
If sharpness < 10 Then
    CalcSharpness = "Blunt"
ElseIf sharpness > 9 And sharpness < 20 Then
    CalcSharpness = "Very Dull"
ElseIf sharpness > 19 And sharpness < 30 Then
    CalcSharpness = "Dull"
ElseIf sharpness > 29 And sharpness < 60 Then
    CalcSharpness = "Medium"
ElseIf sharpness > 59 And sharpness < 70 Then
    CalcSharpness = "Mildly Sharp"
ElseIf sharpness > 69 And sharpness < 80 Then
    CalcSharpness = "Sharp"
ElseIf sharpness > 79 And sharpness < 90 Then
    CalcSharpness = "Very Sharp"
ElseIf sharpness > 89 Then
    CalcSharpness = "Razor Sharp"
End If
End Function
Public Function fCraft(sCName As String) As Integer
Dim SoFar As Integer
For I = 0 To 250
    If ItemType(I) = "Craft" Then
        If sCName = Items(I) Then SoFar = SoFar + 1
    End If
Next I
fCraft = SoFar
End Function
Public Sub ProcInn()
Select Case sMap
Case "Direshore Inn"
    SleepAtInn 10
Case "Amdopple Inn"
    SleepAtInn 20
Case "11-Inn"
    SleepAtInn 15
Case "27-Inn"
    SleepAtInn 30
Case "Sandstone Inn"
    SleepAtInn 30
End Select
End Sub
Public Sub SleepAtInn(icost As Integer)
If MsgBox("Do you wish to stay the night at the Inn and replenish your Stamina and Blood?  It will cost you " & icost & " gold.", vbYesNo) = vbYes Then
    If Gold < icost Then
        MsgBox "You don't have enough Gold."
        Exit Sub
    End If
    St = StM
    Blood = 100
    a ".:: You rest up for the night and awake the next day full of energy!"
    Gold = Gold - icost
    Bleeding = False
    BleedTime = 0
    BleedInt = 0
    Exit Sub
End If
End Sub
Public Sub ProcessExp()
Select Case cLevel
Case 1
    If cExp > 40 Then
        LevelUp
    End If
Case 2
    If cExp > 60 Then
        LevelUp
    End If
Case 3
    If cExp > 100 Then
        LevelUp
    End If
Case 4
    If cExp > 150 Then
        LevelUp
    End If
Case 5
    If cExp > 210 Then
        LevelUp
    End If
Case 6
    If cExp > 270 Then
        LevelUp
    End If
Case 7
    If cExp > 330 Then
        LevelUp
    End If
Case 8
    If cExp > 410 Then
        LevelUp
    End If
Case 9
    If cExp > 480 Then
        LevelUp
    End If
Case 10
    If cExp > 500 Then
        LevelUp
    End If
Case 11
    If cExp > 600 Then
        LevelUp
    End If
Case 12
    If cExp > 700 Then
        LevelUp
    End If
Case 13
    If cExp > 850 Then
        LevelUp
    End If
Case 14
    If cExp > 1000 Then
        LevelUp
    End If
End Select
End Sub
Public Sub LevelUp()
Dim tmp As Integer
cLevel = cLevel + 1
a ".: You have honed your skills to Level " & cLevel
tmp = Int(Rnd * 1) + 2
a ".: Your Strength has gone up by " & tmp
Stre = Stre + tmp
tmp = Int(Rnd * 1) + 2
a ".: Your Endurance has gone up by " & tmp
Endu = Endu + tmp
tmp = Int(Rnd * 1) + 2
a ".: Your Agility has gone up by " & tmp
Agil = Agil + tmp
TrFinding = TrFinding + 1.5
Sharpening = Sharpening + 2
If TrFinding > 50 Then TrFinding = 50
If Sharpening > 50 Then Sharpening = 50
cExp = 0
Blood = Blood + 50
If Blood > 100 Then Blood = 100
St = StM
HarmonyStopMusic
pSound "levelup"
Do: DoEvents
Loop Until isPlaying = False
HarmonyPlayMusic mMusic
End Sub
Public Function CheckFile(InFileName As String) As Boolean

On Error GoTo ErrHandler
CheckFile = False
'Check whether the file or folder exist '
If Dir(InFileName) <> "" Then
'Check is it a directory(folder)
If (GetAttr(InFileName) And vbDirectory) = 0 Then
CheckFile = True
Else
MsgBox "File doesn't exist!", vbCritical
Exit Function
End If

Else
MsgBox "File doesn't exist!", vbCritical
Exit Function
End If

ErrHandler:
End Function
