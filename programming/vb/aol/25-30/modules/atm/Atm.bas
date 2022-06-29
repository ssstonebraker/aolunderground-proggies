Attribute VB_Name = "modATM"
' Offset variable
Public i As Integer

' Currency conversion table from USD
Public ConversionTable As Variant

' High level sound support API
#If Win32 Then
    Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
#Else
    Declare Function sndPlaySound Lib "MMSYSTEM.DLL" _
        (ByVal lpszSoundName As Any, ByVal wFlags As Integer) As Integer
#End If

Global Const SND_ASYNC = &H1     ' Play asynchronously
Global Const SND_NODEFAULT = &H2 ' Don't use default sound
Global Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file

Global SoundBuffer As String

' Mousepointer over command button.
Dim curSelect As StdPicture

Sub ConversionTable_Initialize()
    ConversionTable = Array(1@, 4.8635@, 1.3978@, 1614@, 119.07@, 89.075@)
End Sub

Sub BeginPlaySound(ByVal ResourceId As Integer)
    Dim Ret As Variant
    #If Win32 Then
        ' Important: The returned string is converted to Unicode
        SoundBuffer = StrConv(LoadResData(ResourceId, "ATM_SOUND"), vbUnicode)
    #Else
        SoundBuffer = LoadResData(ResourceId, "ATM_SOUND")
    #End If
    Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    ' Important: This function is neccessary for playing sound asynchronously
    DoEvents
End Sub

Sub EndPlaySound()
    Dim Ret As Variant
    Ret = sndPlaySound(0&, 0&)
End Sub

Sub Cursor_Initialize()
    Set curSelect = LoadResPicture(1, vbResCursor)
End Sub

Sub SetCursor(Button As CommandButton)
    Button.MousePointer = 99
    Button.MouseIcon = curSelect
End Sub
