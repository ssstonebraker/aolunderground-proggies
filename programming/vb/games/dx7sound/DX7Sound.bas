Attribute VB_Name = "DX7Sound"
Option Explicit

'This module was created by D.R Hall
'For more Information and latest version
'E-mail me, derek.hall@virgin.net


Private m_dx As New DirectX7
Private m_dxs As DirectSound 'Then there is the sub object, DirectSound:

Type dxBuffers
  isLoaded As Boolean
  Buffer As DirectSoundBuffer
End Type
Private SoundFolder As String 'Holds Path to Sound folder
Private SB() As dxBuffers 'An Array of BUFFERS,
Private CurrentBuffer As Integer 'Holds last assign Random Buffer Number

Public Sub SoundDir(FolderPath As String)
  SoundFolder = FolderPath & "\"
End Sub

Public Sub CreateBuffers(AmountOfBuffer As Integer, DefaultFile As String)
  ReDim SB(AmountOfBuffer)
  For AmountOfBuffer = 0 To AmountOfBuffer
    DX7LoadSound AmountOfBuffer, DefaultFile 'must assign a defualt sound
    VolumeLevel AmountOfBuffer, 50 ' set volume to 50% for default
  Next AmountOfBuffer
End Sub

Public Sub SetupDX7Sound(CurrentForm As Form)
  Set m_dxs = m_dx.DirectSoundCreate("") 'create a DSound object
 'Next you check for any errors, if there are no errors the user has got DX7 and a functional sound card

  If Err.Number <> 0 Then
    MsgBox "Unable to start DirectSound. Check to see that your sound card is properly installed"
    End
  End If
  m_dxs.SetCooperativeLevel CurrentForm.hwnd, DSSCL_PRIORITY 'THIS MUST BE SET BEFORE WE CREATE ANY BUFFERS
  
  'associating our DS object with our window is important. This tells windows to stop
  'other sounds from interfering with ours, and ours not to interfere with other apps.
  'The sounds will only be played when the from has got focus.
  'DSSCL_PRIORITY=no cooperation, exclusive access to the sound card, Needed for games
  'DSSCL_NORMAL=cooperates with other apps, shares resources, Good for general windows multimedia apps.
  
End Sub

Public Sub DX7LoadSound(Buffer As Integer, sfile As String)
  Dim Filename As String
  Dim bufferDesc As DSBUFFERDESC  'a new object that when filled in is passed to the DS object to describe
  Dim waveFormat As WAVEFORMATEX 'what sort of buffer to create
  
  bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN _
  Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC 'These settings should do for almost any app....
  
  waveFormat.nFormatTag = WAVE_FORMAT_PCM
  waveFormat.nChannels = 2    '2 channels
  waveFormat.lSamplesPerSec = 22050
  waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
  waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
  waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign

  Filename = SoundFolder & sfile
  On Error GoTo Continue
  Set SB(Buffer).Buffer = m_dxs.CreateSoundBufferFromFile(Filename, bufferDesc, waveFormat)
  SB(Buffer).isLoaded = True
  Exit Sub
Continue:
  MsgBox "Error can't find file: " & Filename
End Sub

Public Function PlaySoundAnyBuffer(Filename As String, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
  
  Do While SB(CurrentBuffer).Buffer.GetStatus = DSBSTATUS_PLAYING 'Find an empty buffer
    CurrentBuffer = CurrentBuffer + 1
    If CurrentBuffer > UBound(SB) Then CurrentBuffer = 0
  Loop

  DX7LoadSound CurrentBuffer, Filename
  If PanValue <> 50 Then PanSound CurrentBuffer, PanValue
  If Volume < 100 Then VolumeLevel CurrentBuffer, Volume
  If SB(CurrentBuffer).isLoaded Then SB(CurrentBuffer).Buffer.Play LoopIt 'dsb_looping=1, dsb_default=0
End Function

Public Sub PlaySoundWithPan(Buffer As Integer, Filename As String, Optional Volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte)
  DX7LoadSound Buffer, Filename
  If PanValue <> 50 And PanValue < 100 Then PanSound Buffer, PanValue
  If Volume < 100 Then VolumeLevel Buffer, Volume
  If SB(Buffer).isLoaded Then SB(Buffer).Buffer.Play LoopIt 'dsb_looping=1, dsb_default=0
End Sub

Public Sub PanSound(Buffer As Integer, PanValue As Byte)
  Select Case PanValue
    Case 0
      SB(Buffer).Buffer.SetPan -10000
    Case 100
      SB(Buffer).Buffer.SetPan 10000
    Case Else
      SB(Buffer).Buffer.SetPan (100 * PanValue) - 5000
  End Select
End Sub

Public Sub VolumeLevel(Buffer As Integer, Volume As Byte)
  If Volume > 0 Then ' stop division by 0
    SB(Buffer).Buffer.SetVolume (60 * Volume) - 6000
  Else
    SB(Buffer).Buffer.SetVolume -6000
  End If
End Sub

Public Function IsPlaying(Buffer As Integer) As Long
  IsPlaying = SB(Buffer).Buffer.GetStatus
End Function
