VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Maker"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
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
   Picture         =   "FrmMain.frx":0000
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   752
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "m"
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "i"
      Height          =   255
      Left            =   10560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Width           =   255
   End
   Begin VB.ListBox List3 
      Height          =   480
      Left            =   11040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5640
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   690
      Left            =   11160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6360
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   270
      Left            =   11040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4680
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   840
   End
   Begin VB.CommandButton Command6 
      Caption         =   "s"
      Height          =   255
      Left            =   10560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "l"
      Height          =   255
      Left            =   10560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "n"
      Height          =   255
      Left            =   10560
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "e"
      Height          =   255
      Left            =   10560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Col 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FrmMain.frx":129042
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   10920
      TabIndex        =   4
      Top             =   3480
      Width           =   480
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Col_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Dim ToMName As String
    Warps = Warps + 1
    List1.AddItem Index
    ToMName = InputBox("To Map")
    List1.AddItem ToMName
    n2cgo = True
    For i = 0 To Col.UBound
        Col(i).Left = Col(i).Left + 1000
    Next i
    Me.Picture = LoadPicture(App.Path & "\" & ToMName & ".bmp")
    Col(Index).ForeColor = vbRed
End If
If Button = 2 Then
    
    NPCs = NPCs + 1
    List2.AddItem Index
    List2.AddItem InputBox("Name")
    List2.AddItem InputBox("Say")
    List2.AddItem InputBox("Img #")
    
    Col(Index).ForeColor = vbBlue
End If
End Sub

Private Sub Command1_Click()
ms = 0
Cols = Cols + 1
Col(Cols).Left = 320
Col(Cols).Top = 256
End Sub

Private Sub Command2_Click()
MonOcur = InputBox("occurance of battles")
MonName = InputBox("Monster 1")
MonName2 = InputBox("Monster 2")
End Sub

Private Sub Command3_Click()
mMusic = InputBox("Music")
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
List1.Clear
List2.Clear
List3.Clear
Dim lName As String
Dim sLine As String
lName = InputBox("file")
lName = App.Path & "\" & lName
Me.Picture = LoadPicture(lName & ".bmp")
curMap = lName
Open lName & ".map" For Input As #1
    Input #1, MonOcur
    Input #1, MonName
    Input #1, MonName2
    Input #1, mMusic
    Input #1, Cols
    For i = 0 To Cols
        FrmMain.Col(i).ForeColor = vbBlack
        Input #1, sLine
        FrmMain.Col(i).Left = CInt(sLine)
        Input #1, sLine
        FrmMain.Col(i).Top = CInt(sLine)
        Input #1, sLine
        FrmMain.Col(i).Width = CInt(sLine)
        Input #1, sLine
        FrmMain.Col(i).Height = CInt(sLine)
    Next i
    
    If Warps > -1 Then
        Input #1, Warps
        For i = 0 To Warps
            Input #1, sLine
            FrmMain.Col(CInt(sLine)).ForeColor = vbRed
            FrmMain.List1.AddItem sLine
            Input #1, sLine
            FrmMain.List1.AddItem sLine
            Input #1, sLine
            FrmMain.List1.AddItem sLine
            Input #1, sLine
            FrmMain.List1.AddItem sLine
        Next i
    End If
    
    If NPCs > -1 Then
        Input #1, NPCs
        For i = 0 To NPCs
            Input #1, sLine
            FrmMain.Col(CInt(sLine)).ForeColor = vbBlue
            FrmMain.List2.AddItem sLine
            Input #1, sLine
            FrmMain.List2.AddItem sLine
            Input #1, sLine
            FrmMain.List2.AddItem sLine
            Input #1, sLine
            FrmMain.List2.AddItem sLine
        Next i
    End If
Close #1
End Sub

Private Sub Command6_Click()
If MonName = "" Then
    If MsgBox("No Monster data, continue anyway?", vbYesNo) = vbNo Then Exit Sub
End If
If mMusic = "" Then
    If MsgBox("No Music data, continue anyway?", vbYesNo) = vbNo Then Exit Sub
End If
SaveMap
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()
Dim thatfile As String
thatfile = InputBox("file")
Me.Picture = LoadPicture(App.Path & "\" & thatfile & ".bmp")
curMap = "\" & thatfile
End Sub

Private Sub Form_Load()
Me.Width = 10920
Me.Height = 9015
For i = 1 To 150
    Load Col(i)
    Col(i).Visible = True
Next i
Cols = -1
Warps = -1
NPCs = -1
Tr = -1
curMap = "\Test World"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = X & "/" & Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Shift = 1 And n2cgo = True Then
    n2c = n2c + 1
    If n2c = 1 Then List1.AddItem X
    If n2c = 2 Then
        List1.AddItem Y
        n2cgo = False
        n2c = 0
        Me.Picture = LoadPicture(App.Path & curMap & ".bmp")
        For i = 0 To Col.UBound
            Col(i).Left = Col(i).Left - 1000
        Next i
    End If
End If
End Sub

Private Sub Timer1_Timer()
If Cols > -1 And ms = 0 Then
    If GetKey(vbKeyLeft) = True Then
        Col(Cols).Left = Col(Cols).Left - 32
    End If
    If GetKey(vbKeyRight) = True Then
        Col(Cols).Left = Col(Cols).Left + 32
    End If
    If GetKey(vbKeyUp) = True Then
        Col(Cols).Top = Col(Cols).Top - 32
    End If
    If GetKey(vbKeyDown) = True Then
        Col(Cols).Top = Col(Cols).Top + 32
    End If
End If
If Cols > -1 And ms = 1 Then
    If GetKey(vbKeyLeft) = True And Col(Cols).Width > 30 Then
        Col(Cols).Width = Col(Cols).Width - 32
    End If
    If GetKey(vbKeyRight) = True Then
        Col(Cols).Width = Col(Cols).Width + 32
    End If
    If GetKey(vbKeyUp) = True And Col(Cols).Height > 30 Then
        Col(Cols).Height = Col(Cols).Height - 32
    End If
    If GetKey(vbKeyDown) = True Then
        Col(Cols).Height = Col(Cols).Height + 32
    End If
End If
'If ms = 0 Then Me.Caption = "Map Maker - Move"
'If ms = 1 Then Me.Caption = "Map Maker - Scale"
If GetKey(vbKeyQ) = True Then ms = 0
If GetKey(vbKeyW) = True Then ms = 1
End Sub
