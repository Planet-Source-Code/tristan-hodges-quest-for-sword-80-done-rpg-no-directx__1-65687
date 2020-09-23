Attribute VB_Name = "mMap"
Public ms As Integer
Public Cols As Integer
Public Warps As Integer
Public NPCs As Integer
Public Tr As Integer
Public MonOcur As Integer
Public MonName As String
Public MonName2 As String
Public n2c As Integer
Public n2cgo As Boolean
Public curMap As String
Public mMusic As String
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Function GetKey(lngKey As Long) As Boolean
If GetKeyState(lngKey&) < 0 Then
    GetKey = True
Else
    GetKey = False
End If
End Function
Public Sub SaveMap()
Dim mName As String
mName = InputBox("name")
Open App.Path & "\" & mName & ".map" For Output As #1
    Print #1, MonOcur
    Print #1, MonName
    Print #1, MonName2
    Print #1, mMusic
    Print #1, Cols
    For i = 0 To Cols
        Print #1, FrmMain.Col(i).Left
        Print #1, FrmMain.Col(i).Top
        Print #1, FrmMain.Col(i).Width
        Print #1, FrmMain.Col(i).Height
    Next i
    Print #1, Warps
    If Warps > -1 Then
        For i = 0 To FrmMain.List1.ListCount - 1
            Print #1, FrmMain.List1.List(i)
        Next i
    End If
    Print #1, NPCs
    If NPCs > -1 Then
        For i = 0 To FrmMain.List2.ListCount - 1
            Print #1, FrmMain.List2.List(i)
        Next i
    End If
Close #1
End Sub
