VERSION 5.00
Begin VB.Form FrmItem 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
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
   ScaleHeight     =   2940
   ScaleWidth      =   3390
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   2130
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drop"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Use"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   2655
      Left            =   120
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FrmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
For I = 0 To 250
    If Items(I) <> "" Then
        If GotoBattle = False Then
            If ItemType(I) = "Potion" Or ItemType(I) = "Stone" Or ItemType(I) = "Bandage" Or ItemType(I) = "Craft" Then
                List1.AddItem Items(I)
            End If
        Else
            If ItemType(I) = "Potion" Or ItemType(I) = "Bandage" Then
                List1.AddItem Items(I)
            End If
        End If
    End If
Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
pSound "error"
If GotoBattle = False Then
    FrmMain.Enabled = True
    FrmMain.SetFocus
End If
If GotoBattle = True Then
    FrmBattle.Enabled = True
    FrmBattle.SetFocus
    FrmBattle.Timer1.Enabled = True
    FrmBattle.Timer2.Enabled = True
    FrmBattle.Timer3.Enabled = True
    FrmBattle.Timer4.Enabled = True
    Unload Me
    GotoBattle = False
    Exit Sub
End If
Unload Me
End Sub

Private Sub Label1_Click()
pSound "click"
If List1.ListIndex = -1 Then
    MsgBox "Hmm?  Oh, do you want something?"
    Exit Sub
End If
ParseUse List1.List(List1.ListIndex)
If Right(List1.List(List1.ListIndex), 5) <> "Stone" Then
    aiRemove List1.List(List1.ListIndex)
    List1.RemoveItem List1.ListIndex
End If
If GotoBattle = True Then
    ItemWait = True
    Form_Unload 0
End If
End Sub

Private Sub Label2_Click()
pSound "click"
If List1.ListIndex = -1 Then
    MsgBox "You cannot drop air."
    Exit Sub
End If
aiRemove List1.List(List1.ListIndex)
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Label2_DblClick()
Label1_Click
End Sub

Private Sub List1_Click()
pSound "select"
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbkeyenter Or KeyCode = vbKeySpace Then
    Label1_Click
End If
End Sub
