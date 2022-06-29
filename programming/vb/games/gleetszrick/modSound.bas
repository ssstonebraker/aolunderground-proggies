Attribute VB_Name = "modSound"
Option Explicit
' ---------------------------------------------------------------------
' Global Sound constants, variables, and others used within the game.
'
' *********************************************
' | @ Written by Pranay Uppuluri. @           |
' | @ Copyright (c) 1997-98 Pranay Uppuluri @ |
' *********************************************
'
' VB game example Break-Thru! by Mark Pruett ported
' to Visual Basic DirectX.
'
' Thanks for Patrice Scribe's
' DirectX.TLB for DirectX 3.0 or Higher, his dixuSprite Class, and
' his dixuDirectX module, this game looks to be easy to code.
'
' You can visit Patrice's home page at:
'
'           http://www.chez.com/scribe/  *OR*
'           http://ourworld.compuserve.com/homepages/pscribe/
'
' If it wasn't for his effort, I would have had to do a lot
' more coding than this!
' ---------------------------------------------------------------------

' Main Sound object
Public dsZricks As DirectSound

' Sound Buffers that hold our sound data.
Public dsbPaddleHit As DirectSoundBuffer
Public dsbZrickHit As DirectSoundBuffer
Public dsbWallHit As DirectSoundBuffer
Public dsbMissed As DirectSoundBuffer
Public dsbSetup As DirectSoundBuffer
Public dsbNewLevel As DirectSoundBuffer
Public dsbMove As DirectSoundBuffer

' Sound Resource Index Constants
Public Const wavPaddleHit = 1
Public Const wavMissed = 2
Public Const wavMove = 3
Public Const wavNewLevel = 4
Public Const wavSetup = 5
Public Const wavWallHit = 6
Public Const wavZrickHit = 7

' Sound tmp files
Public Const tmpPaddleHit = "~TMP0001"
Public Const tmpMissed = "~TMP0002"
Public Const tmpMove = "~TMP0003"
Public Const tmpNewLevel = "~TMP0004"
Public Const tmpSetup = "~TMP0005"
Public Const tmpWallHit = "~TMP0006"
Public Const tmpZrickHit = "~TMP0007"

' To play sound or not to play sound... that is the question...
Public bSoundIn As Boolean  ' True indicates to play sound
Public strSound As String   ' Contains a Space character and a Speaker.

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
Declare Function lstrcpy Lib "kernel32" (ByVal lpszDestinationString1 As Any, ByVal lpszSourceString2 As Any) As Long

Public Sub InitSound()
' ---------------------------------------------------------
' Loads the Sound files from the Resource file and prints
' them into the respective .tmp files immediately.
' ---------------------------------------------------------
Dim SoundBuffer As String
    ' Default is to play sound.
    bSoundIn = True
    
    strSound = Chr(SC_SPACE) & Chr(SC_SPEAKER)
    
    ' Set up the DirectSound (notice the DSSCL_NORMAL const)
    DirectSoundCreate ByVal 0&, dsZricks, Nothing
    dsZricks.SetCooperativeLevel frmMain.hwnd, DSSCL_NORMAL
    
    ' Load the wave data from the .EXE(.RES when not in exe format)
    ' file to .tmp files
    Open tmpZrickHit For Output As #1
        SoundBuffer = StrConv(LoadResData(wavZrickHit, "ZRICKS_SOUND"), vbUnicode)
        Print #1, SoundBuffer
    Close #1
    
    SoundBuffer = ""
    
    Open tmpPaddleHit For Output As #1
        SoundBuffer = StrConv(LoadResData(wavPaddleHit, "ZRICKS_SOUND"), vbUnicode)
        Print #1, SoundBuffer
    Close #1
    
    SoundBuffer = ""
    
    Open tmpMissed For Output As #1
        SoundBuffer = StrConv(LoadResData(wavMissed, "ZRICKS_SOUND"), vbUnicode)
        Print #1, SoundBuffer
    Close #1
    
    SoundBuffer = ""
    
    Open tmpMove For Output As #1
        SoundBuffer = StrConv(LoadResData(wavMove, "ZRICKS_SOUND"), vbUnicode)
        Print #1, SoundBuffer
    Close #1
    
    SoundBuffer = ""
    
    Open tmpNewLevel For Output As #1
        SoundBuffer = StrConv(LoadResData(wavNewLevel, "ZRICKS_SOUND"), vbUnicode)
        Print #1, SoundBuffer
    Close #1
    
    SoundBuffer = ""
    
    Open tmpSetup For Output As #1
        SoundBuffer = StrConv(LoadResData(wavSetup, "ZRICKS_SOUND"), vbUnicode)
        Print #1, SoundBuffer
    Close #1
    
    SoundBuffer = ""
    
    Open tmpWallHit For Output As #1
        SoundBuffer = StrConv(LoadResData(wavWallHit, "ZRICKS_SOUND"), vbUnicode)
        Print #1, SoundBuffer
    Close #1
    
    SoundBuffer = ""
    
    Call NoiseGet(dsZricks, tmpZrickHit, dsbZrickHit)
    Call NoiseGet(dsZricks, tmpPaddleHit, dsbPaddleHit)
    Call NoiseGet(dsZricks, tmpMissed, dsbMissed)
    Call NoiseGet(dsZricks, tmpMove, dsbMove)
    Call NoiseGet(dsZricks, tmpNewLevel, dsbNewLevel)
    Call NoiseGet(dsZricks, tmpSetup, dsbSetup)
    Call NoiseGet(dsZricks, tmpWallHit, dsbWallHit)
End Sub

Public Sub NoiseGet(Lds As DirectSound, ByVal fName As String, Ldsb As DirectSoundBuffer)
' --------------------------------------------------------
' Loads a WAV file into a DirectSoundBuffer
' --------------------------------------------------------

Dim hWave As Long
Dim pcmwave As WAVEFORMATEX
Dim lngSize As Long
Dim lngPosition As Long
Dim ptr1 As Long, ptr2 As Long, lng1 As Long, lng2 As Long
Dim aByte() As Byte
    
    ReDim aByte(1 To FileLen(fName))
    
    hWave = FreeFile
    
    Open fName For Binary As hWave
        Get hWave, , aByte
    Close hWave
    
    lngPosition = 1
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) <> "fmt"
        lngPosition = lngPosition + 1
    Wend
    
    CopyMemory VarPtr(pcmwave), VarPtr(aByte(lngPosition + 8)), Len(pcmwave)
    
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) + Chr$(aByte(lngPosition + 3)) <> "data"
        lngPosition = lngPosition + 1
    Wend
    
    CopyMemory VarPtr(lngSize), VarPtr(aByte(lngPosition + 4)), Len(lngSize)
    
    Dim dsbd As DSBUFFERDESC
    
    With dsbd
        .dwSize = Len(dsbd)
        .dwFlags = DSBCAPS_CTRLDEFAULT
        .dwBufferBytes = lngSize
        .lpwfxFormat = VarPtr(pcmwave)
    End With
    
    Lds.CreateSoundBuffer dsbd, Ldsb, Nothing
    
    Ldsb.Lock 0&, lngSize, ptr1, lng1, ptr2, lng2, 0&
    
    CopyMemory ptr1, VarPtr(aByte(lngPosition + 4 + 4)), lng1
    
    If lng2 <> 0 Then
        CopyMemory ptr2, VarPtr(aByte(lngPosition + 4 + 4 + lng1)), lng2
    End If
End Sub

Public Function NoisePlay(Lds As DirectSoundBuffer, Optional PanValue As Long)
If bSoundIn Then  ' play sound if the player wants to hear it.
    Lds.Stop
    If Not IsMissing(PanValue) Then
        Call Lds.SetPan(PanValue)
        Call Lds.Play(0, 0, 0)
    Else
        Call Lds.Play(0, 0, 0)
    End If
End If
End Function

Public Sub DSoundDone()
' ---------------------------------------------------------
' Uninitializes the DirectSoundBuffers, and DirectSound
' object itself.
' ---------------------------------------------------------
Dim i As Long

    Set dsbSetup = Nothing
    Set dsbWallHit = Nothing
    Set dsbZrickHit = Nothing
    Set dsbNewLevel = Nothing
    Set dsbMissed = Nothing
    Set dsbMove = Nothing
    Set dsbPaddleHit = Nothing

    Set dsZricks = Nothing
    
    ' clean up the sound temp files...
    ChDir App.Path
    
    For i = 1 To 7
        Kill "~TMP000" & i
    Next i

End Sub

Public Function GetPanValue(ByVal X As Double, ByVal Width As Long) As Long
' -------------------------------------------------------
' Gets the Pan Value (this enables 3D sound) using the
' object's X position and the Screen's Width.
' -------------------------------------------------------
     GetPanValue = (20000 * (X + (Width / 2)) / 640) - 10000
End Function
