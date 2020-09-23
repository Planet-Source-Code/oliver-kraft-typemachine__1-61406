Attribute VB_Name = "Module1"
'Declaraciones y funcion para reproducir un sonido

Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long
' Constantes para PlaySound
Const SND_SYNC = (0)
Const SND_ASYNC = (1)
Const SND_NODEFAULT = (2)
Const SND_MEMORY = (4)
Const SND_LOOP = (8)
Const SND_NOSTOP = (16)
Public Sub PlaySound()
Dim rc
rc = sndplaysound(App.Path & "/tic.wav", SND_NODEFAULT + SND_ASYNC)
End Sub

