Attribute VB_Name = "SoundMod"
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Sub Hold Sound Functions, Api's and Variables '
'____________________________________________________'
' api for midi
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
' api for wave
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
' wave consts
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
' loop a wave
Sub WAVLoop(File As String)
On Error Resume Next
   ThisDir
   wFlags% = SND_ASYNC Or SND_LOOP
   X = sndPlaySound(File, wFlags%)
End Sub
' play a wave once
Sub WAVPlay(File As String)
On Error Resume Next
    ThisDir
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(" ", wFlags%)
    X = sndPlaySound(File, wFlags%)
End Sub
' play midi, no loop yet
Sub MidiPlay(File As String)
 On Error Resume Next
 ChDrive App.Path
 ChDir App.Path
 MIDIPath$ = File
 X& = mciSendString("open " & MIDIPath & " Type sequencer Alias MFile", 0&, 0, 0)
 X& = mciSendString("play MFile", 0&, 0, 0)
End Sub
Public Sub StopSounds(Optional Wav As Boolean, Optional Midi As Boolean)
 On Error Resume Next
If Wav = True Then Call WAVPlay(" ")
If Midi = True Then
 X& = mciSendString("stop MFile", 0&, 0, 0)
 X& = mciSendString("close MFile", 0&, 0, 0)
End If
End Sub
