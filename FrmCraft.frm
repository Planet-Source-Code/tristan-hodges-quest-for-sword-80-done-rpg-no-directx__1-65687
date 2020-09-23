VERSION 5.00
Begin VB.Form FrmCraft 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Here's what I can craft for you."
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   840
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Crab Shelling:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Hardened Leather:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Brass Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Hellviel Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Damascus Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Charcoal:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Steel Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lace Wrappings:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label nWindCloth 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   52
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label nCrabShelling 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   51
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label nHardenedLeather 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   50
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label nLeather 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   49
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label nHellvielChunk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label nDamascusChunk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   47
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label nCharcoal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   46
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label nBrassChunk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   45
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label nLaceWrappings 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   44
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label nSteelChunk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   43
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label nRuby 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   42
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label nDiamond 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   41
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label yWindCloth 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   40
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label yCrabShelling 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   39
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label yHardenedLeather 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   38
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label yLeather 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   37
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label yHellvielChunk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   36
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label yDamascusChunk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   35
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label yCharcoal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label yBrassChunk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   33
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label yLaceWrappings 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   32
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label ySteelChunk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   31
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label yRuby 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   30
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label yDiamond 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   29
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Wind Cloth:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Crab Shelling:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Hardened Leather:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Leather:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Hellviel Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Damascus Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Charcoal:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Brass Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Lace Wrappings:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Materials you need:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Steel Chunk:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Ruby:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Diamond:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Craft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Wind Cloth:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Leather:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Materials you have:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ruby:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Diamond:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   3855
      Left            =   120
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I would like to craft this..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmCraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Height = 2925
Select Case sMap
Case "Direshore Weapon Shop"
    List1.AddItem "Longsword"
    List1.AddItem "Full-Tang Wakizashi"
Case "Sandstone Weapon Shop"
    List1.AddItem "Damascus Katana"
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
pSound "error"
FrmMain.Enabled = True
FrmMain.SetFocus
Unload Me
End Sub

Private Sub Label1_Click()
pSound "click"
If List1.ListIndex = -1 And fCraft("Wind Cloth") = 1 Then
    MsgBox "You have crafted Air !"
    ParseTreasure "Air"
    Exit Sub
End If
Crafting = List1.List(List1.ListIndex)
List1.Enabled = False
Me.Top = Me.Top - 2000
Me.Height = 6885
yDiamond = fCraft("Diamond")
yRuby = fCraft("Ruby")
ySteelChunk = fCraft("Steel Chunk")
yLaceWrappings = fCraft("Lace Wrappings")
yBrassChunk = fCraft("Brass Chunk")
yCharcoal = fCraft("Charcoal")
yDamascusChunk = fCraft("Damascus Steel Chunk")
yhellviel = fCraft("Hellviel Chunk")
yLeather = fCraft("Leather")
yHardenedLeather = fCraft("Hardened Leather")
yCrabShelling = fCraft("Crab Shelling")
yWindCloth = fCraft("Wind Cloth")
End Sub

Private Sub Label15_Click()
pSound "click"
If yDiamond < nDiamond Then GoTo Nooo:
If yRuby < nRuby Then GoTo Nooo:
If ySteelChunk < nSteelChunk Then GoTo Nooo:
If yLaceWrappings < nLaceWrappings Then GoTo Nooo:
If yBrassChunk < nBrassChunk Then GoTo Nooo:
If yCharcoal < nCharcoal Then GoTo Nooo:
If yDamascusChunk < nDamascusChunk Then GoTo Nooo:
If yHellvielChunk < nHellvielChunk Then GoTo Nooo:
If yLeather < nLeather Then GoTo Nooo:
If yHardenedLeather < nHardenedLeather Then GoTo Nooo:
If yCrabShelling < nCrabShelling Then GoTo Nooo:
If yWindCloth < nWindCloth Then GoTo Nooo:
ParseTreasure Crafting
MsgBox "You have sucessfully crafted: " & Crafting
For I = 1 To nDiamond
    aiRemove "Diamond"
Next I
For I = 1 To nRuby
    aiRemove "Ruby"
Next I
For I = 1 To nSteelChunk
    aiRemove "Steel Chunk"
Next I
For I = 1 To nLaceWrappings
    aiRemove "Lace Wrappings"
Next I
For I = 1 To nBrassChunk
    aiRemove "Brass Chunk"
Next I
For I = 1 To nCharcoal
    aiRemove "Charcoal"
Next I
For I = 1 To nDamascusChunk
    aiRemove "Damascus Steel Chunk"
Next I
For I = 1 To nHellvielChunk
    aiRemove "Hellviel Chunk"
Next I
For I = 1 To nLeather
    aiRemove "Leather"
Next I
For I = 1 To nHardenedLeather
    aiRemove "Hardened Leather"
Next I
For I = 1 To nCrabShelling
    aiRemove "Crab Shelling"
Next I
For I = 1 To nWindCloth
    aiRemove "Wind Cloth"
Next I
a ">> Crafted Item: " & Crafting & " !"
Crafting = ""
FrmMain.Enabled = True
FrmMain.SetFocus
Unload Me
Exit Sub
Nooo:
MsgBox "You don't have enough materials to craft that of which you wanted to craft!"
End Sub

Private Sub List1_Click()
pSound "select"
If List1.ListIndex = -1 Then Exit Sub
Dim om As String
om = List1.List(List1.ListIndex)
Select Case om
Case "Longsword"
    nSteelChunk = 1
    nCharcoal = 2
    nCrabShelling = 1
Case "Full-Tang Wakizashi"
    nSteelChunk = 3
    nLaceWrappings = 1
    nCharcoal = 2
Case "Damascus Katana"
    nDamascusChunk = 4
    nSteelChunk = 2
    nRuby = 1
    nLaceWrappings = 2
End Select
End Sub

Private Sub Timer1_Timer()
yDiamond.FontBold = True
nDiamond.FontBold = True
yRuby.FontBold = True
nRuby.FontBold = True
ySteelChunk.FontBold = True
nSteelChunk.FontBold = True
yLaceWrappings.FontBold = True
nLaceWrappings.FontBold = True
yBrassChunk.FontBold = True
nBrassChunk.FontBold = True
yCharcoal.FontBold = True
nCharcoal.FontBold = True
yDamascusChunk.FontBold = True
nDamascusChunk.FontBold = True
yHellvielChunk.FontBold = True
nHellvielChunk.FontBold = True
yLeather.FontBold = True
nLeather.FontBold = True
yHardenedLeather.FontBold = True
nHardenedLeather.FontBold = True
yCrabShelling.FontBold = True
nCrabShelling.FontBold = True
yWindCloth.FontBold = True
nWindCloth.FontBold = True
If yDiamond = 0 Then yDiamond.FontBold = False
If nDiamond = 0 Then nDiamond.FontBold = False
If yRuby = 0 Then yRuby.FontBold = False
If nRuby = 0 Then nRuby.FontBold = False
If ySteelChunk = 0 Then ySteelChunk.FontBold = False
If nSteelChunk = 0 Then nSteelChunk.FontBold = False
If yLaceWrappings = 0 Then yLaceWrappings.FontBold = False
If nLaceWrappings = 0 Then nLaceWrappings.FontBold = False
If yBrassChunk = 0 Then yBrassChunk.FontBold = False
If nBrassChunk = 0 Then nBrassChunk.FontBold = False
If yCharcoal = 0 Then yCharcoal.FontBold = False
If nCharcoal = 0 Then nCharcoal.FontBold = False
If yDamascusChunk = 0 Then yDamascusChunk.FontBold = False
If nDamascusChunk = 0 Then nDamascusChunk.FontBold = False
If yHellvielChunk = 0 Then yHellvielChunk.FontBold = False
If nHellvielChunk = 0 Then nHellvielChunk.FontBold = False
If yLeather = 0 Then yLeather.FontBold = False
If nLeather = 0 Then nLeather.FontBold = False
If yHardenedLeather = 0 Then yHardenedLeather.FontBold = False
If nHardenedLeather = 0 Then nHardenedLeather.FontBold = False
If yCrabShelling = 0 Then yCrabShelling.FontBold = False
If nCrabShelling = 0 Then nCrabShelling.FontBold = False
If yWindCloth = 0 Then yWindCloth.FontBold = False
If nWindCloth = 0 Then nWindCloth.FontBold = False
End Sub

