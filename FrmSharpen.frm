VERSION 5.00
Begin VB.Form FrmSharpen 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weapon Sharpening"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
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
   ScaleHeight     =   1815
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   360
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   1290
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lskill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sharpening Skill:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lsharp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sharpness:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sharpen"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FrmSharpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For I = 0 To 250
    If ItemType(I) = "Sword" Then
        List1.AddItem Items(I)
    End If
Next I
Me.Caption = "Sharpening using: " & UsingStone
End Sub

Private Sub Form_Unload(Cancel As Integer)
pSound "error"
Unload FrmItem
Unload Me
End Sub

Private Sub Label1_Click()
Randomize
pSound "click"
If List1.ListIndex = -1 Then
    MsgBox "You think to yourself, 'The air is already sharp enough.'  You shiver for a moment."
    Exit Sub
End If
If UsingStone = "Broke" Then
    MsgBox "Your Sharpening Stone has broke.  Go back to the Item Menu and Use another to sharpen more, or go to town and buy another stone in the Item shop if you don't have any more."
    Exit Sub
End If
Dim tmpS As Integer
For I = 0 To 250
    If Items(I) = List1.List(List1.ListIndex) Then
        tmpS = ItemDull(I)
    End If
Next I
Select Case UsingStone
Case "Pocket Sharpening Stone"
    ShPlus = 2
Case "Poor Sharpening Stone"
    ShPlus = 4
Case "Fine Sharpening Stone"
    ShPlus = 6
Case "Quality Sharpening Stone"
    ShPlus = 8
Case "Exquisite Sharpening Stone"
    ShPlus = 10
End Select
If Sharpening * 2 < tmpS Then
    If MsgBox("You realize that this blade is sharper than you are at sharpening!  There's a very small chance you'd be able to make this blade any more sharp than it already is.  Infact, almost surely you'll end up breaking your stone.  Try to sharpen it anyway?", vbYesNo) = vbNo Then Exit Sub
    If Int(Rnd * 20) > 0 Then
        GoTo break:
    Else
        Sharpening = Sharpening + 0.6
    End If
End If
tmpS = tmpS + Int(Rnd * (Sharpening / 3)) + ShPlus
For I = 0 To 250
    If Items(I) = List1.List(List1.ListIndex) Then
        ItemDull(I) = tmpS
        Select Case UsingStone
        Case "Pocket Sharpening Stone"
            Sharpening = Sharpening + 0.1
        Case "Poor Sharpening Stone"
            Sharpening = Sharpening + 0.1
        Case "Fine Sharpening Stone"
            Sharpening = Sharpening + 0.2
        Case "Quality Sharpening Stone"
            Sharpening = Sharpening + 0.2
        Case "Exquisite Sharpening Stone"
            Sharpening = Sharpening + 0.25
        End Select
    End If
Next I
If eWeap = List1.List(List1.ListIndex) Then
    eDull = tmpS
End If
If Int(Rnd * 3) = 1 Then
break:
    MsgBox "Your sharpening stone broke!"
    aiRemove UsingStone
    UsingStone = "Broke"
End If
End Sub

Private Sub Timer1_Timer()
lskill = Sharpening
If List1.ListIndex = -1 Then
    lsharp = "N/A"
    Exit Sub
End If
For I = 0 To 250
    If Items(I) = List1.List(List1.ListIndex) Then
        lsharp = ItemDull(I) & "(" & CalcSharpness(ItemDull(I)) & ")"
    End If
Next I
End Sub
