VERSION 5.00
Begin VB.Form FrmEquip 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipment"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
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
   ScaleHeight     =   2850
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Current Equipment"
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   3615
      Begin VB.Label lStats 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999/999/999"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Str/End/Agi:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lCurB 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Old Shortsword of Energy"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Boots:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lCurS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Old Shortsword of Energy"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Shield:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lcurA 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Old Shortsword of Energy"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Armor:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lCurW 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Old Shortsword of Energy"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Weapon:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   2340
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drop"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Stats"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unequip"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Equip"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   2600
      Left            =   120
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "FrmEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
For I = 0 To 250
    If ItemType(I) = "Sword" Or ItemType(I) = "Armor" Or ItemType(I) = "Boots" Or ItemType(I) = "Shield" Then
        List1.AddItem Items(I)
    End If
Next I
If eDull < 10 And eWeap <> "" Then
    MsgBox "Your weapon is extremely dull.  Buy a sharpening stone and sharpen it, you'll really notice a difference!"
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
If List1.ListIndex = -1 Then
    MsgBox "The air is already equipped!"
    Exit Sub
End If
Dim thei As Integer
For I = 0 To 250
    If Items(I) = List1.List(List1.ListIndex) Then
        thei = I
        Exit For
    End If
Next I
If thei = -1 Then
    MsgBox "Critical Error while finding item.  What the hell are you doing?"
    Exit Sub
End If
Select Case ItemType(thei)
Case "Sword"
    eWeap = Items(thei)
    bStr = ItemStr(thei)
    bAgi = ItemAgi(thei)
    bEnd = ItemEnd(thei)
    eDull = ItemDull(thei)
Case "Armor"
    eArmor = Items(thei)
    bArmorStr = ItemStr(thei)
    bArmorAgi = ItemAgi(thei)
    bArmorEnd = ItemEnd(thei)
Case "Shield"
    eShield = Items(thei)
    bShEnd = ItemEnd(thei)
    bShAgi = ItemAgi(thei)
Case "Boots"
    eBoots = Items(thei)
    bBootsMove = ItemStr(thei)
    bBootsAgi = ItemAgi(thei)
End Select
End Sub

Private Sub Label2_Click()
pSound "click"
If List1.ListIndex = -1 Then
    MsgBox "The air cannot be unequipped."
    Exit Sub
End If
Dim thei As Integer
For I = 0 To 250
    If Items(I) = List1.List(List1.ListIndex) Then
        thei = I
        Exit For
    End If
Next I
Select Case ItemType(thei)
Case "Sword"
    bStr = 0
    bAgi = 0
    bEnd = 0
    eWeap = ""
Case "Armor"
    bArmorStr = 0
    bArmorAgi = 0
    bArmorEnd = 0
    eArmor = ""
Case "Shield"
    bShEnd = 0
    bShAgi = 0
    eShield = ""
Case "Boots"
    bBootsAgi = 0
    bBootsMove = 0
    eBoots = ""
End Select
End Sub

Private Sub Label3_Click()
pSound "click"
If List1.ListIndex = -1 Then
    MsgBox "It would take a god to drop the air."
    Exit Sub
End If
If MsgBox("Really drop item?  It's gone for good if you do.", vbYesNo) = vbYes Then
Dim thei As Integer
For I = 0 To 250
    If Items(I) = List1.List(List1.ListIndex) Then
        thei = I
        Exit For
    End If
Next I
Label2_Click
aiRemove Items(thei)
List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub Label4_Click()
pSound "click"
If List1.ListIndex = -1 Then
    MsgBox "Air:" & vbCrLf & "Sharpness: Dull"
    Exit Sub
End If
Dim thei As Integer
For I = 0 To 250
    If Items(I) = List1.List(List1.ListIndex) Then
        thei = I
        Exit For
    End If
Next I
If ItemType(thei) = "Sword" Then
    MsgBox Items(thei) & " (" & ItemType(thei) & ")" & vbCrLf & vbCrLf & "  Attack: " & ItemStr(thei) & vbCrLf & "  Endurance Mod: " & ItemEnd(thei) & vbCrLf & "  Agility Mod: " & ItemAgi(thei) & vbCrLf & vbCrLf & "  Sharpness: " & ItemDull(thei) & " (" & CalcSharpness(ItemDull(thei)) & ")"
ElseIf ItemType(thei) = "Armor" Then
    MsgBox Items(thei) & " (" & ItemType(thei) & ")" & vbCrLf & vbCrLf & "  Strength Mod: " & ItemStr(thei) & vbCrLf & "  Defense: " & ItemEnd(thei) & vbCrLf & "  Agility Mod: " & ItemAgi(thei)
ElseIf ItemType(thei) = "Shield" Then
    MsgBox Items(thei) & " (" & ItemType(thei) & ")" & vbCrLf & vbCrLf & "  Strength Mod: " & ItemStr(thei) & vbCrLf & "  Defense: " & ItemEnd(thei) & vbCrLf & "  Agility Mod: " & ItemAgi(thei)
Else
    MsgBox Items(thei) & " (Boots)" & vbCrLf & vbCrLf & "  Move Speed Bonus: " & ItemStr(I) & vbCrLf & "  Agility Bonus: " & ItemAgi(thei)
End If
End Sub

Private Sub List1_Click()
pSound "select"
End Sub

Private Sub Timer1_Timer()
lCurW = eWeap
lcurA = eArmor
lCurS = eShield
lCurB = eBoots
lStats = Stre + bStr + bArmorStr & "/" & Endu + bEnd + bArmorEnd + bShEnd & "/" & Agil + bAgi + bArmorAgi + bBootsAgi + bShAgi
If lCurW = "" Then lCurW = "None Equipped"
If lcurA = "" Then lcurA = "None Equipped"
If lCurS = "" Then lCurS = "None Equipped"
If lCurB = "" Then lCurB = "None Equipped"
End Sub
