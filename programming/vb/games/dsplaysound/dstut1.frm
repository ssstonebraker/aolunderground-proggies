VERSION 5.00
Begin VB.Form DStutform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DS Play Sound"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "dstut1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chloop 
      Caption         =   "Loop Play"
      Height          =   195
      Left            =   3480
      TabIndex        =   14
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "&Stop"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   1200
      Width           =   750
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "P&ause"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   1200
      Width           =   750
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   1200
      Width           =   750
   End
   Begin VB.HScrollBar scrlPan 
      Height          =   255
      LargeChange     =   1000
      Left            =   600
      Max             =   10000
      Min             =   -10000
      SmallChange     =   500
      TabIndex        =   9
      Top             =   840
      Width           =   4215
   End
   Begin VB.HScrollBar scrlVol 
      Height          =   255
      LargeChange     =   20
      Left            =   840
      Max             =   0
      Min             =   -5000
      SmallChange     =   255
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      Height          =   195
      Index           =   3
      Left            =   4920
      TabIndex        =   10
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   270
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Panning:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   630
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Volume:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "An Example of a DX game made in VB."
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   2760
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"dstut1.frx":0442
      Height          =   1215
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Label lblLink1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.microsoft.com/directx"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   2
      Top             =   1680
      Width           =   2385
   End
   Begin VB.Label lbllink2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.parkstonemot.freeserve.co.uk/indexfw.htm"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      Top             =   2160
      Width           =   3900
   End
   Begin VB.Label lbllink3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mailto: Jollyjeffers@GreenOnions.netscapeonline.co.uk"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   1920
      Width           =   3900
   End
End
Attribute VB_Name = "DStutform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DIRECT X 7 is initialised using a reference type library, in VB5, open the PROJECT menu, then select
'REFERENCES. in the list there will be "DirectX 7 for visual basic type library", with a check next to it.
'When making your own projects you need to select this library.....

'Each DX app needs to declare a DX object, similiar to a control, sub objects such as DSound/DDraw are
'created from this master object:
Dim m_dx As New DirectX7
'Then there is the sub object, DirectSound:
Dim m_ds As DirectSound
'Sound is loaded into BUFFERS, these buffers represent different areas of memory on your sound card
'or system memory. You must have 1 buffer for each .wav file you load. Although you can keep reloading
'different files into the same buffer, you can only have one file in each buffer at any one time. For
'this example, we only need one buffer.
Dim m_dsBuffer As DirectSoundBuffer

Dim m_bLoaded As Boolean



'USED FOR THE LINKS, NOT FOR DX
#If Win32 Then
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
#Else
Private Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, ByVal lpszDir$, ByVal fsShowCmd%) As Integer
Private Declare Function GetDesktopWindow Lib "USER" () As Integer
#End If
Private Const SW_SHOWNORMAL = 1

Function StartDoc(DocName As String) As Long
      Dim Scr_hDC As Long
      Scr_hDC = GetDesktopWindow()
      StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function
Private Sub Form_Load()
    Me.Show
    On Local Error Resume Next
    'First we have to create a DSound object, this must be done before any features can be used.
    'It must also be done before we set the cooperativelevel or create any buffers.
    Set m_ds = m_dx.DirectSoundCreate("")
    'This checks for any errors, if there are no errors the user has got DX7 and a functional sound card
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound. Check to see that your sound card is properly installed"
        End
    End If
    'THIS MUST BE SET BEFORE WE CREATE ANY BUFFERS
    'associating our DS object with our window is important. This tells windows to stop
    'other sounds from interfering with ours, and ours not to interfere with other apps.
    'The sounds will only be played when the from has got focus....
    'DSSCL_PRIORITY=no cooperation, exclusive access to the sound card
        'Needed for games
    'DSSCL_NORMAL=cooperates with other apps, shares resources
        'Good for general windows multimedia apps.
    m_ds.SetCooperativeLevel Me.hwnd, DSSCL_PRIORITY
    
    
End Sub


Sub LoadWave(i As Integer, sfile As String)

    Dim bufferDesc As DSBUFFERDESC  'a new object that when filled in is passed to the DS object to describe
    'what sort of buffer to create
    Dim waveFormat As WAVEFORMATEX
    'These settings should do for almost any app....
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2    '2 channels
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    'this next line creates a buffer with the specified file in it, 'BufferDesc' and 'waveformat'
    'describe the properties of the new buffer. they can only be modified when creating a new buffer...
    Set m_dsBuffer = m_ds.CreateSoundBufferFromFile(sfile, bufferDesc, waveFormat)

    'checks for any errors
    If Err.Number <> 0 Then
        MsgBox "unable to find " + sfile
        End
    End If
    
    'check the panning and volume "properties"
    scrlPan_Change
    scrlVol_Change
    
End Sub



Private Sub cmdPlay_Click()
    'if there is no sound loaded, load the sound:
    If m_bLoaded = False Then
        m_bLoaded = True
        LoadWave 0, App.Path & "\info.wav"
    End If
            
    
    Dim flag As Long
    flag = 0
    If chloop.Value <> 0 Then flag = 1  'decide whether or not too loop the sound
    
    'the play statement has these possibilities
    'dsb_looping=1
    'dsb_default=0
    m_dsBuffer.Play flag

End Sub


Private Sub cmdStop_Click()
    If m_dsBuffer Is Nothing Then Exit Sub 'if the user clicks stop when nothing has been loaded
    'the stop function doesn't 'rewind' the sound back to the beginning
    m_dsBuffer.Stop
    'so we have to tell it to go back to the beginnning
    m_dsBuffer.SetCurrentPosition 0
    'this line ^^ can be used to start a sound 1/2 way through......
End Sub

Private Sub chLoop_Click()
    If chloop.Value = 0 Then
        cmdStop_Click
    End If
End Sub


Private Sub cmdPause_Click()

    If m_dsBuffer Is Nothing Then Exit Sub
    'as stated in the stop section, without the 'setcurrentposition' statement it will not
    'go back to the beggining
    m_dsBuffer.Stop
    
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLink1.FontUnderline = False
lbllink2.FontUnderline = False
lbllink3.FontUnderline = False
lblLink1.FontBold = False
lbllink2.FontBold = False
lbllink3.FontBold = False
End Sub

Private Sub lblLink1_Click()
Dim Z As Long
Z = StartDoc("http://www.microsoft.com/directx")
End Sub


Private Sub lblLink1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLink1.FontUnderline = True
lbllink2.FontUnderline = False
lbllink3.FontUnderline = False
lblLink1.FontBold = True
lbllink2.FontBold = False
lbllink3.FontBold = False
End Sub


Private Sub lbllink2_Click()
Dim Y As Long
Y = StartDoc("http://www.parkstonemot.freeserve.co.uk/indexfw.htm")
End Sub

Private Sub lbllink2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLink1.FontUnderline = False
lbllink2.FontUnderline = True
lbllink3.FontUnderline = False
lblLink1.FontBold = False
lbllink2.FontBold = True
lbllink3.FontBold = False
End Sub


Private Sub lbllink3_Click()
Dim X As Long
X = StartDoc("mailto: jollyjeffers@greenonions.netscapeonline.co.uk")
End Sub


Private Sub lbllink3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLink1.FontUnderline = False
lbllink2.FontUnderline = False
lbllink3.FontUnderline = True
lblLink1.FontBold = False
lbllink2.FontBold = False
lbllink3.FontBold = True
End Sub



Private Sub scrlVol_Change()
    If m_dsBuffer Is Nothing Then Exit Sub
    'you can't set the volume value without the buffer being created, so you must handle this
    'otherwise your app will crash.......
    m_dsBuffer.SetVolume scrlVol.Value
End Sub
Private Sub scrlVol_Scroll()
    If m_dsBuffer Is Nothing Then Exit Sub
    m_dsBuffer.SetVolume scrlVol.Value
End Sub




Private Sub scrlPan_Change()
    If m_dsBuffer Is Nothing Then Exit Sub
    'you can't set the panning value without the buffer being created, so you must handle this
    'otherwise your app will crash.......
    m_dsBuffer.SetPan scrlPan.Value
End Sub
Private Sub scrlPan_Scroll()
    If m_dsBuffer Is Nothing Then Exit Sub
    m_dsBuffer.SetPan scrlPan.Value
End Sub



