VERSION 5.00
Begin VB.Form FrmIShop 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Shop"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
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
   ScaleHeight     =   2670
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   360
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   2130
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lCurG 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Gold:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Cost:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buy Item"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   3000
      X2              =   4680
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sell Your Items"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buy Shop's Items"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   2415
      Left            =   120
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "FrmIShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label1_Click
If FirstItemShop = 0 Then
    FirstItemShop = 1
    MsgBox "The Item Shop is the best place to sell the stuff you get from monsters.  But that's self explanitory.  I'm here to explain the wonderous uses of Sharpening Stones and Bandages, which can appear misleading at first.  I'll be brief.  Sharpening stones can sharpen your weapons, and the quality of the sharpening stone makes a sharper sword.  There are 3 types of bandages.  Small Bandages (which stops grazes and light cuts), Large Bandages (which stop large cuts, even mortal wounds), and Battle-Ready Bandage.  The Battle-Ready Bandage is the only bandage you can use in battle.  It stops most bleeding, save a mortal wound.  It will, however, stunt a mortal wound and make it less lethal.  Use these wisely."
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
pSound "error"
FrmMain.Enabled = True
FrmMain.SetFocus
Unload Me
End Sub

Private Sub Label1_Click()
pSound "click"
CurSB = 0
SelCost = 0
Label3 = "Buy Item"
List1.Clear
List1.AddItem "Weak Potion of Energy"
List1.AddItem "Potion of Energy"
List1.AddItem "Quality Potion of Energy"
List1.AddItem "Exquisite Potion of Energy"
List1.AddItem "Packet of Blood"
List1.AddItem "Small Bandage"
List1.AddItem "Large Bandage"
List1.AddItem "Battle-Ready Bandage"
List1.AddItem "Pocket Sharpening Stone"
List1.AddItem "Poor Sharpening Stone"
List1.AddItem "Fine Sharpening Stone"
List1.AddItem "Quality Sharpening Stone"
List1.AddItem "Exquisite Sharpening Stone"
List1.AddItem "Imp Skin"
List1.AddItem "Diamond"
List1.AddItem "Ruby"
List1.AddItem "Steel Chunk"
List1.AddItem "Lace Wrappings"
List1.AddItem "Brass Chunk"
List1.AddItem "Charcoal"
If sMap = "Sandstone Item Shop" Then
     List1.AddItem "Hellviel Chunk"
     List1.AddItem "Wind Cloth"
     List1.AddItem "Leather"
End If
End Sub

Private Sub Label2_Click()
pSound "click"
CurSB = 1
SelCost = 0
List1.Clear
For I = 0 To 250
    If Items(I) <> "" Then
        If ItemType(I) = "Sword" Or ItemType(I) = "Armor" Or ItemType(I) = "Boots" Or ItemType(I) = "Shield" Or ItemType(I) = "Potion" Or ItemType(I) = "Stone" Or ItemType(I) = "Bandage" Or ItemType(I) = "Craft" Then
            List1.AddItem Items(I)
        End If
    End If
Next I
Label3 = "Sell Item"
End Sub

Private Sub Label3_Click()
If List1.ListIndex = -1 Then Exit Sub
If Label3 = "Buy Item" Then
    If Gold < lCost Then
        MsgBox "Not enough gold!"
        Exit Sub
    End If
    Gold = Gold - lCost
    ParseTreasure List1.List(List1.ListIndex)
    pSound "coins"
Else
    If eWeap = List1.List(List1.ListIndex) Then
        MsgBox "Unequip this item before selling it."
        Exit Sub
    End If
    If eArmor = List1.List(List1.ListIndex) Then
        MsgBox "Unequip this item before selling it."
        Exit Sub
    End If
    If eShield = List1.List(List1.ListIndex) Then
        MsgBox "Unequip this item before selling it."
        Exit Sub
    End If
    If eBoots = List1.List(List1.ListIndex) Then
        MsgBox "Unequip this item before selling it."
        Exit Sub
    End If
    Gold = Gold + lCost
    aiRemove List1.List(List1.ListIndex)
    List1.RemoveItem List1.ListIndex
    pSound "coins"
End If
End Sub

Private Sub List1_Click()
pSound "select"
If List1.ListIndex = -1 Then Exit Sub
'If CurSB = 0 Then
    ParseTreasure List1.List(List1.ListIndex)
    For I = 0 To 250
        If Items(I) = List1.List(List1.ListIndex) Then
            SelCost = ItemWorth(I)
            aiRemove List1.List(List1.ListIndex)
            Exit Sub
        End If
    Next I
'End If
End Sub

Private Sub Timer1_Timer()
If CurSB = 1 Then
    Label1.BackColor = &HFF8080
    Label2.BackColor = &HC00000
    If Label4.Caption <> "Sell For:" Then Label4.Caption = "Sell For:"
Else
    Label1.BackColor = &HC00000
    Label2.BackColor = &HFF8080
    If Label4.Caption <> "Item Cost:" Then Label4.Caption = "Item Cost:"
End If
lCurG = Gold
If CurSB = 0 Then
    lCost = SelCost
Else
    lCost = CInt(SelCost / 1.5)
End If
End Sub
