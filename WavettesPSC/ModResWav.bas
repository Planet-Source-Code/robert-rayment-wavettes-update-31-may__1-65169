Attribute VB_Name = "ModResWav"
' ModResWav.bas   ~Wavettes~

Option Explicit

Public Declare Function PlayWAV Lib "winmm.dll" Alias "PlaySoundA" _
(lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long


Public Const SND_ASYNC = &H1         ' return to program immediately
Public Const SND_MEMORY = &H4        ' play the sound from memory
Public Const SND_NODEFAULT = &H2     ' don't play the default sound if not available
'Public Const SND_NOSTOP = &H10       ' don't stop a currently playing sound
'Public Const SND_NOWAIT = &H2000     ' return immediately if driver not available
Public Const SND_PURGE = &H40        ' purge non-static events for task
Public Const SND_LOOP = &H8          ' loop the sound until next sndPlaySound
'Public Const SND_RESOURCE = &H40004  ' play from resource

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function PlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
   (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private WAVData$, WAVData2$
Private Sound() As Byte


Public Sub LoadWav()
Dim U As Long
   Sound = LoadResData(101, "PLAYFIRST")
   U = UBound(Sound()) + 1
   WAVData$ = Space$(U)
   CopyMemory ByVal WAVData$, Sound(0), U
   
   Sound = LoadResData(102, "NONAME")
   U = UBound(Sound()) + 1
   WAVData2$ = Space$(U)
   CopyMemory ByVal WAVData2$, Sound(0), U
End Sub
Public Sub PlayFirst()
   ' Complete play first
   PlaySound ByVal WAVData$, SND_NODEFAULT Or SND_MEMORY 'Or SND_NOSTOP
End Sub

Public Sub PlayNoName()
   ' Complete play first
   PlaySound ByVal WAVData2$, SND_NODEFAULT Or SND_MEMORY 'Or SND_NOSTOP
End Sub

Public Sub StopPlay()
   'PlaySound 0, SND_PURGE
   PlayWAV vbNull, 0, SND_PURGE
End Sub
