Option Explicit

'-----------------------------------------------------------------------------------------
'   Name    :   SOUNDER.BAS
'   Author  :   Peter Wright
'   Date    :   1 March 1994
'
'   Notice  :   This code is freely distributable.
'           :   Peter Wright, Psynet Ltd.
'           :   peter@gendev.demon.co.uk
'-----------------------------------------------------------------------------------------


' To play sound we need to use a feature of windows known as the MCI, the Multimedia Control
' interface. Unfortunately its only accessible through a DLL call. (We cover DLLS a little
' later in the book
Declare Function mciSendString Lib "MMSystem" (ByVal SendString As String, ByVal ReturnString As String, ByVal ReturnLength As Integer, ByVal Callback As Integer) As Long

Sub PlayIt (sFileName As String)
    
'--------------------------------------------------------------------------------------
'   SubName :   Playit
'   Author  :   Peter Wright
'   Date    :   1 March 1994
'
'   Params  :   sFileName - the name of the WAV or MID file to play
'
'   Notes   :   This code uses MCI to play either a WAV or MID file.
'
'--------------------------------------------------------------------------------------

    ' Declare a long variable to catch the return code from the MCI calls
    Dim lReturnCode As Long

    ' In the following code there is not much too worry about. If the file is the wrong
    ' format, or does not exist then MCI will catch the error and simply remain silent.
    ' For this reason there is no error checking in the code.

    ' If the file is a MID file then
    If UCase(Right$(sFileName, 4)) = ".MID" Then
        ' Tell MCI to open the file as a Sequencer file (MIDI), and call it SOUND
        lReturnCode = mciSendString("Open " & sFileName & " Type Sequencer Alias SOUND", "", 0, 0)
    
    ' Otherwise, if it is a WAV file, then
    ElseIf UCase(Right$(sFileName, 4)) = ".WAV" Then
        lReturnCode = mciSendString("Open " & sFileName & " Type WAVEAUDIO Alias SOUND", "", 0, 0)
    
    '...otherwise don't waste time, just quit
    Else
        Exit Sub

    End If
        
    ' OK, now tell MCI to play the file that we called SOUND, and WAIT until it has finished
    ' playing before returning.
    lReturnCode = mciSendString("Play SOUND Wait", "", 0, 0)

    ' Finally, tell MCI to close the file we called SOUND
    lReturnCode = mciSendString("Close SOUND", "", 0, 0)

End Sub

