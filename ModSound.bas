Attribute VB_Name = "ModSound"
Public Declare Function HarmonyCreate Lib "harmony.dll" () As Long
Public Declare Function HarmonyFadeInMusic Lib "harmony.dll" (ByVal TimeFactor As Long) As Long
Public Declare Function HarmonyFadeOutMusic Lib "harmony.dll" (ByVal TimeFactor As Long) As Long
Public Declare Function HarmonyGetVersion Lib "harmony.dll" () As Long
Public Declare Function HarmonyInitMidi Lib "harmony.dll" () As Long
Public Declare Function HarmonyPlayMusic Lib "harmony.dll" (ByVal Filename As String) As Long
Public Declare Function HarmonyRelease Lib "harmony.dll" () As Long
Public Declare Function HarmonySetMusicPanpot Lib "harmony.dll" (ByVal Panpot As LoadPictureColorConstants) As Long
Public Declare Function HarmonySetMusicSpeed Lib "harmony.dll" (ByVal Speed As Long) As Long
Public Declare Function HarmonySetMusicVolume Lib "harmony.dll" (ByVal NewVolume As Long) As Long
Public Declare Function HarmonyStopMusic Lib "harmony.dll" () As Long
Public Declare Function HarmonyTermMidi Lib "harmony.dll" () As Long
Option Explicit

Public mFilename As String
Public mFlags As Long

Public Declare Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long

Public Property Get Filename() As String
Filename = mFilename
End Property

Public Property Let Filename(ByVal vNewValue As String)
mFilename = vNewValue
End Property

Public Property Get Flags() As Long
Flags = mFlags
End Property

Public Property Let Flags(ByVal vNewValue As Long)
mFlags = vNewValue
End Property

Public Sub Play(TheFile As String)
Dim rc As Long
rc = sndPlaySound(App.Path & "\Sound\" & TheFile & ".wav", 1)
End Sub

