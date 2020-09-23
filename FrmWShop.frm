VERSION 5.00
Begin VB.Form FrmWShop 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weapon Shop"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
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
   ScaleHeight     =   3075
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   360
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   2550
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lGold 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lAtk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lMods 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lSharp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99(Razor Sharp)"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Sharp:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "End/Agi Mod:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   3240
      X2              =   5400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Gold:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buy Item"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   3240
      X2              =   5400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buy Shop's Items"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sell Your Items"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   2805
      Left            =   120
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "FrmWShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
CurSB = 0
Label1_Click
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
List1.Clear
Label3 = "Buy Item"
ProcessWShop
End Sub

Private Sub Label2_Click()
pSound "click"
CurSB = 1
SelCost = 0
List1.Clear
For I = 0 To 250
    If Items(I) <> "" Then
        If ItemType(I) = "Sword" Or ItemType(I) = "Armor" Or ItemType(I) = "Shield" Or ItemType(I) = "Boots" Or ItemType(I) = "Craft" Or ItemType(I) = "Potion" Or ItemType(I) = "Bandage" Then
            List1.AddItem Items(I)
        End If
    End If
Next I
Label3 = "Sell Item"
End Sub

Private Sub Label3_Click()
If List1.ListIndex = -1 Then Exit Sub
Select Case Label3
Case "Buy Item"
    If Gold < lCost Then
        MsgBox "I'm afraid you don't have enough gold to purchase this item.  So sorry."
        Exit Sub
    End If
    Gold = Gold - SelCost
    ParseTreasure List1.List(List1.ListIndex)
    pSound "coins"
Case "Sell Item"
    If eWeap = List1.List(List1.ListIndex) Then
        eWeap = ""
        bStr = 0
        bAgi = 0
        bEnd = 0
    End If
        If eArmor = List1.List(List1.ListIndex) Then
        eArmor = ""
        bArmorStr = 0
        bArmorAgi = 0
        bArmorEnd = 0
    End If
    If eShield = List1.List(List1.ListIndex) Then
        eShield = ""
        bShEnd = 0
        bShAgi = 0
    End If
    If eBoots = List1.List(List1.ListIndex) Then
        eBoots = ""
        bBootsAgi = 0
        bBootsMove = 0
    End If
    Gold = Gold + (SelCost / 1.3)
    aiRemove List1.List(List1.ListIndex)
    List1.RemoveItem List1.ListIndex
    lSharp = ""
    lAtk = "0"
    lmod = "0"
    SelCost = 0
    lCost = 0
    pSound "coins"
End Select
End Sub

Private Sub List1_Click()
pSound "select"
If List1.ListIndex = -1 Then Exit Sub
    ParseTreasure List1.List(List1.ListIndex)
    For I = 0 To 250
        If Items(I) = List1.List(List1.ListIndex) Then
            SelCost = ItemWorth(I)
            lAtk = ItemStr(I)
            lMods = ItemEnd(I) & "/" & ItemAgi(I)
            lSharp = ItemDull(I) & "(" & CalcSharpness(ItemDull(I)) & ")"
            aiRemove List1.List(List1.ListIndex)
            Exit Sub
        End If
    Next I
End Sub

Private Sub Timer1_Timer()
If CurSB = 1 Then
    Label1.BackColor = &HFF8080
    Label2.BackColor = &HC00000
    If Label3.Caption <> "Sell For:" Then Label4.Caption = "Sell For:"
    lCost = CInt(SelCost / 1.5)
Else
    Label1.BackColor = &HC00000
    Label2.BackColor = &HFF8080
    If Label3.Caption <> "Item Cost:" Then Label4.Caption = "Item Cost:"
    lCost = SelCost
End If
lGold = Gold
End Sub
