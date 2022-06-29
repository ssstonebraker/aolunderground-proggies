VERSION 4.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "MCI Sample by Matt Hart vbhelp@matthart.com"
   ClientHeight    =   4800
   ClientLeft      =   1650
   ClientTop       =   1950
   ClientWidth     =   8055
   Height          =   5205
   Left            =   1590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8055
   Top             =   1605
   Width           =   8175
   Begin VB.CommandButton Command18 
      Caption         =   "Eject CD"
      Height          =   315
      Left            =   1980
      TabIndex        =   12
      Top             =   2820
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Next CD Track"
      Height          =   315
      Left            =   1620
      TabIndex        =   28
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Prev CD Track"
      Height          =   315
      Left            =   120
      TabIndex        =   27
      Top             =   3180
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6300
      ScaleHeight     =   975
      ScaleWidth      =   1755
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6240
      ScaleHeight     =   975
      ScaleWidth      =   1755
      TabIndex        =   25
      Top             =   2940
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Copy AVI Image"
      Height          =   495
      Left            =   1500
      TabIndex        =   15
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7560
      Top             =   4200
   End
   Begin VB.CommandButton Command14 
      Caption         =   "MID Vol UP"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "MID Vol DN"
      Height          =   375
      Left            =   1620
      TabIndex        =   9
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "WAV Vol DN"
      Height          =   375
      Left            =   1620
      TabIndex        =   5
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "WAV Vol UP"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Play Recorded WAV"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   1275
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Play CD && Record to WAV"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   2715
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movie!"
      Height          =   2655
      Left            =   3000
      TabIndex        =   16
      Top             =   60
      Width           =   4935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Stop CD"
      Height          =   315
      Left            =   1020
      TabIndex        =   11
      Top             =   2820
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Stop AVI"
      Height          =   495
      Left            =   1620
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Stop MID"
      Height          =   495
      Left            =   1620
      TabIndex        =   7
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop WAV"
      Height          =   495
      Left            =   1620
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play CD"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play AVI"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play MID"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play WAV"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.PictureBox ProgressBar4 
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   24
      Top             =   4320
      Width           =   4095
   End
   Begin VB.PictureBox ProgressBar3 
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   21
      Top             =   3840
      Width           =   4095
   End
   Begin VB.PictureBox ProgressBar2 
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   20
      Top             =   3360
      Width           =   4095
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   17
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "CD"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "MID"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "AVI"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WAV"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' MCI Sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' You don't need the MCI32.OCX file to create some great Multimedia
' VB apps - when you know the right commands.
' All you need is available through mciSendString.  Experiment
' with the calls and commands to see what you can do - it took me
' about 10 minutes of experimenting to finally get an AVI movie to play
' in a frame.  Enjoy!

Private Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Type lVolType
    v As Long
End Type

Private Type VolType
    lv As Integer
    rv As Integer
End Type

Const SM_CXBORDER = 5
Const SM_CYBORDER = 6
Const VK_SNAPSHOT As Byte = &H2C
Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Const MCI_SET = &H80D
Const MCI_SET_DOOR_OPEN = &H100&

Private Sub Command1_Click()        ' PLAY WAV
    Command5_Click
    Dim i As Long, RS As String, cb As Long, W$
    RS = Space$(128)
    W$ = "c:\windows\media\robot~17.wav"
    i = mciSendString("open waveaudio!" & W$ & " alias sound", RS, 128, cb)
    If i Then MsgBox "Error! Probably file not found. Modify the code to point to a .WAV file on your system."
    i = mciSendString("play sound", RS, 128, cb)
End Sub

Private Sub Command16_Click()       ' Prev CD Track
    Dim i As Long, RS As String, cb As Long, W$, s As Long
    RS = Space$(128)
    i = mciSendString("status cdaudio current track", RS, 128, cb)
    If Val(RS) Then
        s = Val(RS) - 1
        i = mciSendString("status cdaudio position track " & s, RS, 128, cb)
        s = Val(RS)
        If s Then
            i = mciSendString("play cdaudio from " & s, RS, 128, cb)
            ProgressBar4.Line (0, 0)-(ProgressBar4.Width, ProgressBar4.Height), ProgressBar4.BackColor, BF
        End If
    End If
End Sub

Private Sub Command17_Click()       ' Next CD Track
    Dim i As Long, RS As String, cb As Long, W$, s As Long
    RS = Space$(128)
    i = mciSendString("status cdaudio current track", RS, 128, cb)
    If Val(RS) Then
        s = Val(RS) + 1
        i = mciSendString("status cdaudio position track " & s, RS, 128, cb)
        s = Val(RS)
        If s Then
            i = mciSendString("play cdaudio from " & s, RS, 128, cb)
        End If
    End If
End Sub

Private Sub Command18_Click()       ' EJECT CD
    Command8_Click
    Dim i As Long, RS As String, cb As Long, id As Long
    RS = Space$(128)
    i = mciSendString("open cdaudio", RS, 128, cb)
    id = mciGetDeviceID("cdaudio")
    i = mciSendCommand(id, MCI_SET, MCI_SET_DOOR_OPEN, 0)
    Command8_Click
End Sub

Private Sub Command5_Click()        ' STOP WAV
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("stop sound", RS, 128, cb)
    i = mciSendString("close sound", RS, 128, cb)
End Sub

Private Sub Command3_Click()        ' PLAY AVI
    Command7_Click
    Dim i As Long, RS As String, cb As Long, A$, x As Long, y As Long
    RS = Space$(128)
    A$ = "d:\winvideo\matts\ktkd.avi"
    i = mciSendString("open AVIvideo!" & A$ & " alias movie parent " & Frame1.hWnd & " style child", RS, 128, cb)
'    i = mciSendString("put movie window client at 200 0 0 0", RS, 128, cb)
    If i Then MsgBox "Error! Probably file not found. Modify the code to point to an .AVI file on your system."
    i = mciSendString("play movie", RS, 128, cb)
End Sub

Private Sub Command7_Click()        ' STOP AVI
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("stop movie", RS, 128, cb)
    i = mciSendString("close movie", RS, 128, cb)
End Sub

Private Sub Command15_Click()       ' COPY AVI IMAGE
    Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
    DoEvents
    Picture1.Move 0, 0, Width, Height
    Picture2.Move 0, 0, Frame1.Width, Frame1.Height
    DoEvents
    Picture1.Picture = Clipboard.GetData
    ' The positions must be adjusted for Form caption height and border width for exact frame match!
    ' Also, the Frame size should be adjusted to match the AVI !
    Picture2.PaintPicture Picture1.Picture, 0, 0, Frame1.Width, Frame1.Height, Frame1.Left, Frame1.Top, Frame1.Width, Frame1.Height
    Clipboard.Clear
    Clipboard.SetData Picture2.Image
    MsgBox "Image copied to the clipboard."
End Sub

Private Sub Command10_Click()       ' PLAY RECORDED WAV
    Dim i As Long
    i = sndPlaySound("c:\cdtest.wav", 0)
    If i = 0 Then MsgBox "Error! Probably file not found. Modify the code to record and play a .WAV file on your system."
End Sub

Private Sub Command11_Click()       ' INCREASE WAV VOLUME
    Dim id As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    id = -1     ' the ALL DEVICE id - this will change the master WAVE volume!
    i = waveOutGetVolume(id, v)
    lVol.v = v
    LSet Vol = lVol
    lv = Vol.lv: rv = Vol.rv
    lv = lv + &HFFF
    rv = rv + &HFFF
    If lv > 32767 Then lv = lv - 65536
    If rv > 32767 Then rv = rv - 65536
    Vol.lv = lv
    Vol.rv = rv
    LSet lVol = Vol
    v = lVol.v
    i = waveOutSetVolume(id, v)
End Sub

Private Sub Command12_Click()       ' DECREASE WAV VOLUME
    Dim id As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    id = -1     ' the ALL DEVICE id - this will change the master WAVE volume!
    i = waveOutGetVolume(id, v)
    lVol.v = v
    LSet Vol = lVol
    lv = Vol.lv: rv = Vol.rv
    lv = lv - &HFFF
    rv = rv - &HFFF
    If lv < -32768 Then lv = 65535 + lv
    If rv < -32768 Then rv = 65535 + rv
    Vol.lv = lv
    Vol.rv = rv
    LSet lVol = Vol
    v = lVol.v
    i = waveOutSetVolume(id, v)
End Sub

Private Sub Command13_Click()       ' DECREASE MIDI VOLUME
    Dim id As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    id = mciGetDeviceID("midi")   ' I don't know the master MIDI id
    i = midiOutGetVolume(id, v)
    lVol.v = v
    LSet Vol = lVol
    lv = Vol.lv: rv = Vol.rv
    lv = lv - &HFFF
    rv = rv - &HFFF
    If lv < -32768 Then lv = 65535 + lv
    If rv < -32768 Then rv = 65535 + rv
    Vol.lv = lv
    Vol.rv = rv
    LSet lVol = Vol
    v = lVol.v
    i = midiOutSetVolume(id, v)
End Sub

Private Sub Command14_Click()       ' INCREASE MIDI VOLUME
    Dim id As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    id = mciGetDeviceID("midi")   ' I don't know the master MIDI id
    i = midiOutGetVolume(id, v)
    lVol.v = v
    LSet Vol = lVol
    lv = Vol.lv: rv = Vol.rv
    lv = lv + &HFFF
    rv = rv + &HFFF
    If lv > 32767 Then lv = lv - 65536
    If rv > 32767 Then rv = rv - 65536
    Vol.lv = lv
    Vol.rv = rv
    LSet lVol = Vol
    v = lVol.v
    i = midiOutSetVolume(id, v)
End Sub

Private Sub Command2_Click()           ' PLAY MIDI FILE
    Command6_Click
    Dim i As Long, RS As String, cb As Long, W$
    RS = Space$(128)
    W$ = "c:\sb16\samples\minuet.mid"
    i = mciSendString("open sequencer!" & W$ & " alias midi", RS, 128, cb)
    If i Then MsgBox "Error! Probably file not found. Modify the code to record and play a .MID file on your system."
    i = mciSendString("play midi", RS, 128, cb)
End Sub

Private Sub Command4_Click()            ' PLAY CD
    Command8_Click
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("open cdaudio", RS, 128, cb)
    i = mciSendString("set cdaudio time format milliseconds", RS, 128, cb)
    i = mciSendString("play cdaudio", RS, 128, cb)
End Sub

Private Sub Command6_Click()            ' STOP MIDI
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("stop midi", RS, 128, cb)
    i = mciSendString("close midi", RS, 128, cb)
End Sub

Private Sub Command8_Click()            ' STOP CD
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("stop cdaudio", RS, 128, cb)
    i = mciSendString("close cdaudio", RS, 128, cb)
End Sub

Private Sub Command9_Click()            ' RECORD WAV
    On Local Error Resume Next
    Dim i As Long, RS As String, cb As Long, t#
    RS = Space$(128)
    Kill "c:\cdtest.wav"
    Command4_Click
    i = mciSendString("open new type waveaudio alias capture", RS, 128, cb)
    i = mciSendString("record capture", RS, 128, cb)
    t# = Timer + 1: Do Until Timer > t#: DoEvents: Loop
    i = mciSendString("stop capture", RS, 128, cb)
    i = mciSendString("save capture c:\cdtest.wav", RS, 128, cb)
    i = mciSendString("close capture", RS, 128, cb)
    Command8_Click
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    
    i = mciSendString("stop sound", RS, 128, cb)
    i = mciSendString("stop midi", RS, 128, cb)
    i = mciSendString("stop movie", RS, 128, cb)
    i = mciSendString("stop cdaudio", RS, 128, cb)
    i = mciSendString("stop capture", RS, 128, cb)

    i = mciSendString("close sound", RS, 128, cb)
    i = mciSendString("close midi", RS, 128, cb)
    i = mciSendString("close movie", RS, 128, cb)
    i = mciSendString("close cdaudio", RS, 128, cb)
    i = mciSendString("close capture", RS, 128, cb)
End Sub

Private Sub Timer1_Timer()
    ' Uses Picture controls - comment the ProgressBar#.Line lines and uncomment the
    ' .Min .Max .Value ProgressBar# lines to use a real Win 95 Progress Bar
    ' (If you have that control with VB 4 or 5 Pro or Enterprise versions)
    Dim i As Long, RS As String, cb As Long, s As Single
    RS = Space$(128)
    i = mciSendString("status sound length", RS, 128, cb)
    If Val(RS) Then
        'ProgressBar1.Max = Val(RS)
        s = Val(RS)
        i = mciSendString("status sound position", RS, 128, cb)
        s = Val(RS) / s
        ProgressBar1.Line (0, 0)-(ProgressBar1.Width * s, ProgressBar1.Height), QBColor(4), BF
        'ProgressBar1.Value = Val(RS)
    Else
        'ProgressBar1.Value = 0
        ProgressBar1.Line (0, 0)-(ProgressBar1.Width, ProgressBar1.Height), ProgressBar1.BackColor, BF
    End If
    
    i = mciSendString("status movie length", RS, 128, cb)
    If Val(RS) Then
        'ProgressBar2.Max = Val(RS)
        s = Val(RS)
        i = mciSendString("status movie position", RS, 128, cb)
        s = Val(RS) / s
        ProgressBar2.Line (0, 0)-(ProgressBar2.Width * s, ProgressBar2.Height), QBColor(4), BF
        'ProgressBar2.Value = Val(RS)
    Else
        'ProgressBar2.Value = 0
        ProgressBar2.Line (0, 0)-(ProgressBar2.Width, ProgressBar2.Height), ProgressBar2.BackColor, BF
    End If
    
    i = mciSendString("status midi length", RS, 128, cb)
    If Val(RS) Then
        'ProgressBar3.Max = Val(RS)
        s = Val(RS)
        i = mciSendString("status midi position", RS, 128, cb)
        s = Val(RS) / s
        ProgressBar3.Line (0, 0)-(ProgressBar3.Width * s, ProgressBar3.Height), QBColor(4), BF
        'ProgressBar3.Value = Val(RS)
    Else
        'ProgressBar3.Value = 0
        ProgressBar3.Line (0, 0)-(ProgressBar3.Width, ProgressBar3.Height), ProgressBar3.BackColor, BF
    End If
    
    i = mciSendString("status cdaudio length", RS, 128, cb)
    If Val(RS) Then
        'ProgressBar4.Max = Val(RS)
        s = Val(RS)
        i = mciSendString("status cdaudio position", RS, 128, cb)
        s = Val(RS) / s
        ProgressBar4.Line (0, 0)-(ProgressBar4.Width * s, ProgressBar4.Height), QBColor(4), BF
        'ProgressBar4.Value = Val(RS)
    Else
        'ProgressBar4.Value = 0
        ProgressBar4.Line (0, 0)-(ProgressBar4.Width, ProgressBar4.Height), ProgressBar4.BackColor, BF
    End If
End Sub
