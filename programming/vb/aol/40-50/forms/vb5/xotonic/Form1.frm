VERSION 5.00
Object = "{24365B29-A3B5-11D1-B8B0-444553540000}#1.0#0"; "XFXFORMSHAPER.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG~2.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Xotonic 0.1 beta"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   2160
      TabIndex        =   11
      Top             =   600
      Width           =   3150
      Visible         =   0   'False
      _ExtentX        =   5556
      _ExtentY        =   582
      _Version        =   327681
      BorderStyle     =   0
      AutoEnable      =   0   'False
      BackVisible     =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.ListBox Playlist 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1215
      Left            =   1440
      TabIndex        =   10
      Top             =   4080
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.ListBox PlayPath 
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1440
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   1320
   End
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   120
      Top             =   3000
      _ExtentX        =   1852
      _ExtentY        =   1296
   End
   Begin VB.Label testpl 
      BackStyle       =   0  'Transparent
      Caption         =   "LP"
      Height          =   210
      Left            =   2400
      TabIndex        =   25
      ToolTipText     =   "load playlist"
      Top             =   2800
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   1320
      Picture         =   "Form1.frx":0442
      Top             =   3960
      Width           =   4305
   End
   Begin VB.Label listshow 
      BackStyle       =   0  'Transparent
      Caption         =   "PL"
      Height          =   210
      Left            =   6000
      TabIndex        =   24
      ToolTipText     =   "hide/show playlist"
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label listhide 
      BackStyle       =   0  'Transparent
      Caption         =   "PL"
      Height          =   210
      Left            =   4150
      TabIndex        =   23
      ToolTipText     =   "show/hide playlist"
      Top             =   2800
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   1455
      Left            =   1320
      Tag             =   "shaper"
      Top             =   3960
      Width           =   4320
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "off"
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   720
      TabIndex        =   22
      ToolTipText     =   "ontop off"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "on"
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   2760
      TabIndex        =   21
      ToolTipText     =   "ontop on"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label ontoplabel 
      BackStyle       =   0  'Transparent
      Caption         =   "ontop:"
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label quitbutton 
      BackStyle       =   0  'Transparent
      Caption         =   "quit"
      Height          =   210
      Left            =   4200
      TabIndex        =   19
      ToolTipText     =   "quit"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label minimized 
      BackStyle       =   0  'Transparent
      Caption         =   "min"
      Height          =   210
      Left            =   3840
      TabIndex        =   18
      ToolTipText     =   "minimize player"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label cDn 
      Caption         =   "cDn"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label cUp 
      Caption         =   "cUp"
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "add"
      Height          =   210
      Left            =   3480
      TabIndex        =   15
      ToolTipText     =   "add files to playlist"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lTot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3060
      TabIndex        =   14
      ToolTipText     =   "total lenght"
      Top             =   1830
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "current played"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label cDel 
      Caption         =   "Label1"
      Height          =   135
      Left            =   5040
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image iSlider 
      Height          =   120
      Left            =   2275
      Picture         =   "Form1.frx":1536
      ToolTipText     =   "slider - drag to adjust"
      Top             =   2250
      Width           =   345
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Height          =   365
      Index           =   7
      Left            =   3690
      TabIndex        =   8
      ToolTipText     =   " Next "
      Top             =   3030
      Width           =   345
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Height          =   365
      Index           =   6
      Left            =   2690
      TabIndex        =   7
      ToolTipText     =   " Prev "
      Top             =   3030
      Width           =   335
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Height          =   175
      Index           =   3
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   " Play "
      Top             =   3250
      Width           =   460
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Height          =   175
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   " Stop "
      Top             =   3000
      Width           =   500
   End
   Begin VB.Label pbc 
      BackStyle       =   0  'Transparent
      Caption         =   "open"
      ForeColor       =   &H00400040&
      Height          =   210
      Index           =   0
      Left            =   3010
      TabIndex        =   1
      ToolTipText     =   "open file"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lSong 
      BackStyle       =   0  'Transparent
      Caption         =   "Song Label"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "song - click to see ID3 tag"
      Top             =   1610
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   2120
      Shape           =   3  'Circle
      Tag             =   "shaper"
      Top             =   1155
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   2040
      Picture         =   "Form1.frx":1CE9
      Tag             =   "shaper"
      Top             =   1080
      Width           =   2670
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code made by : Peter Hollandare - audionow.home.ml.org
' Copyright (c) 1998 Peter Hollandare & audionow.home.ml.org
' You are *free* to modify this code as much as you whant
' if you give me the credit for it!
' You *MUST* have MediaPlayer for win95/98/NT installed
' in order to get this player to play!
' Future updates of this code can be found on the page above.
'
' ^Mouse on EfNet in channel #MPEG3



Private Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
'Private Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
'Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
'Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long


Public PrgName, Sect
Public DragFlag, SlideFlag, PlVisFlag
Public IX, IY, TX, TY, FX, FY
Public TNum, SongLen, Dev$
Public SongName As String, SongPath As String, SongTitle As String
Public AddName As String, AddPath As String, AddTitle As String
Public DefPath As String, Info As String
Public Intro As Boolean, STP As Boolean
Public Playing As Boolean, Paused As Boolean


Dim Title As String
Dim InBuf As String * 256
'Dim CDSong As String            ' string to hold the tracks
'Dim fastForwardSpeed As Long    ' seconds to seek for ff/rew
Dim fPlaying As Boolean         ' true if CD is currently playing
'Dim fCDLoaded As Boolean        ' true if CD is the the player
'Dim numTracks As Integer        ' number of tracks on audio CD
Dim trackLength() As String     ' array containing length of each track
Dim track As Integer            ' current track
Dim min As Integer              ' current minute on track
Dim sec As Integer              ' current second on track
Dim cmd As String               ' string to hold mci command strings
Dim mouseIsDown As Boolean
Dim cx As Single
Dim cy As Single

'*************
' Volume part
'*************
Private Type lVolType
   v As Long
End Type

Private Type VolType
    lv As Integer
    rv As Integer
End Type
Private Sub Command1_Click()
Dim ID As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    ID = -1     ' the ALL DEVICE id - this will change the master WAVE volume!
    i = waveOutGetVolume(ID, v)
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
    i = waveOutSetVolume(ID, v)
    
End Sub

Private Sub Command2_Click()
Dim ID As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    ID = -1     ' the ALL DEVICE id - this will change the master WAVE volume!
    i = waveOutGetVolume(ID, v)
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
    i = waveOutSetVolume(ID, v)
End Sub
Sub VolumeDown()

Dim ID As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    ID = -1     ' the ALL DEVICE id - this will change the master WAVE volume!
    i = waveOutGetVolume(ID, v)
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
    i = waveOutSetVolume(ID, v)
    
End Sub
Sub VolumeUp()

Dim ID As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double
    ID = -1     ' the ALL DEVICE id - this will change the master WAVE volume!
    i = waveOutGetVolume(ID, v)
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
    i = waveOutSetVolume(ID, v)
    
End Sub

Private Sub Command3_Click()
Form1.Show

End Sub

Private Sub Command4_Click()
Call Quit
End Sub

Private Sub Form_Activate()
    'Call AlwaysOnTop(Me, OptAlwaysOnTop)
AlwaysOnTop Form1, 1
End Sub

Private Sub Form_Load()
    'get window initial position
    
    'test this
    'frmOptions.Show 1
    'songtest.Visible = False
    
    PrgName = "xotonic": Sect = "config"
    'X = GetSetting(PrgName, Sect, "X", Me.Left)
    'Y = GetSetting(PrgName, Sect, "Y", Me.Top)
    DefPath = GetSetting(PrgName, Sect, "Path", "")
    'OptSnap = GetSetting(PrgName, Sect, "Snap", 0)
    'OptAlwaysOnTop = GetSetting(PrgName, Sect, "AlwaysOnTop", 0)
    
    Me.Move x, Y
    'Me.Height = 4000
    
        
    Call ClearInf
    
    'Make hotspot areas invisible
    'lTitlebar.BackStyle = 0: lClose.BackStyle = 0
    'cIDTag.BackStyle = 0: cPrefs.BackStyle = 0
    'cMixer.BackStyle = 0: cPlaylist.BackStyle = 0
    'cUp.BackStyle = 0: cDn.BackStyle = 0
    'cDel.BackStyle = 0: cLoad.BackStyle = 0
    'cClear.BackStyle = 0: cSave.BackStyle = 0
    'cAddFile.BackStyle = 0: cAddDir.BackStyle = 0
    'cSkin.BackStyle = 0: cWShade.BackStyle = 0
    'cMin.BackStyle = 0
    For J = 0 To 7: pbc(J).BackStyle = 0: Next
    
    'Set status lights
    'sStop.Visible = True: sPause.Visible = False
    'sPlay.Visible = False: sIntro.Visible = False: sSTP.Visible = False
    

    'Load frmIcon 'system tray icon/menu
    
    'Load playlist and set last track
    'F$ = App.Path + "\playlist.m3u"
    'If Dir$(F$) <> "" Then
      'ReadPL (F$)
      'TNum = Val(GetSetting(PrgName, Sect, "LastTrack", "0"))
      'Call PlayIt
    
    'Title = String(30, " ") + Playlist.List(TNum - 1)
    'Title = String(30, " ") + CommonDialog1.FileTitle
    
    'frmVBAmp.Caption = "Xotonic"
    'End If
    Title = String(30, " ") + "Xotonic 0.1 beta Copyright © 1998 Peter Hollandare"
    'Label3.Caption = "Xotonic 1.0 Copyright © 1998 Peter Hollandare"
 FormShaper1.ShapeIt "shaper"
 'Label1.Alignment = 2
 
 'iSlider.Visible = True

Playlist.Visible = False
FormShaper1.Tag = ""
Shape2.Visible = False
Shape2.Tag = ""
FormShaper1.ShapeIt "shaper"
listhide.Visible = False
listshow.Visible = True
listshow.Left = 4150
listshow.Top = 2800


End Sub

Private Sub Form_Resize()
'Timer4.Enabled = False
'frmVBAmp.Caption = "Mouse Player-3 version 1.06beta"
'Timer4.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
    'DoEvents
    
    MMControl1.Command = "close"
    'save window position
    'SaveSetting PrgName, Sect, "X", Str$(Me.Left)
    'SaveSetting PrgName, Sect, "Y", Str$(Me.Top)
    'SaveSetting PrgName, Sect, "Snap", OptSnap
    'SaveSetting PrgName, Sect, "AlwaysOnTop", OptAlwaysOnTop
    
    'Save playlist and current track
    'SaveSetting PrgName, Sect, "LastTrack", Str$(TNum)
    'WritePL (App.Path + "\playlist.m3u")
    'End
End Sub
Sub Quit()
    Unload Me
End Sub

Private Sub lClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    A$ = UCase$(Chr$(KeyAscii))
    P = InStr("XSZPIN-+ ?VLCADKTMQ12", A$)
    If P > 0 Then
        Select Case P
            Case 1 To 8: Call pbc_Click(P - 1)
            Case 9: Call PauseIt
            Case 10: Call ShowPrefs
            Case 11: Call TogglePL
            Case 12: Call LoadPL
            Case 13: Call ClearPL
            Case 14: Call AddFile
            Case 15: Call AddDir
            Case 16: Call LoadSkin
            Case 17: Call ShowInfo
            Case 18: Call ShowMixer
            Case 19: Call QuitIt
            Case 20: Call VolumeDown
            Case 21: Call VolumeUp
        End Select
    End If
    KeyAscii = 0
End Sub

Private Sub Image7_Click()
cAddFile_Click
End Sub

Private Sub Image8_Click()
Call LoadPL
End Sub

Private Sub Label2_Click()
cAddFile_Click
End Sub

Private Sub Label3_Click()
'Label3.Caption = "off"
AlwaysOnTop Form1, 0
Label3.Visible = False
Label4.Visible = True
Label4.Left = 2760
Label4.Top = 2520
End Sub

Private Sub Label4_Click()
'Label3.Caption = "off"
AlwaysOnTop Form1, 1
Label3.Visible = True
Label4.Visible = False
End Sub

Private Sub listhide_Click()
Playlist.Visible = False
FormShaper1.Tag = ""
Shape2.Visible = False
Shape2.Tag = ""
FormShaper1.ShapeIt "shaper"
listhide.Visible = False
listshow.Visible = True
listshow.Left = 4150
listshow.Top = 2800
End Sub

Private Sub listshow_Click()
Playlist.Visible = True
'FormShaper1.Tag = ""
Shape2.Visible = True
Shape2.Tag = "shaper"
FormShaper1.ShapeIt "shaper"
listhide.Visible = True
listshow.Visible = False

End Sub

Private Sub lSong_DblClick()
   Call ShowInfo
End Sub
Private Sub cMin_Click()
  Me.Visible = False
End Sub

Private Sub cPrefs_Click()
    Call ShowPrefs
End Sub
Private Sub cWShade_Click()
    Dim Tmpint As Integer

  
    If Me.Height = 1740 Then
        For Tmpint = 1740 To 200 Step -70
            DoEvents
            Me.Height = Tmpint
        Timer6.Enabled = True
        songtest.Enabled = True
        songtest.Visible = True
        'songtest.BackColor = &H8000E '&H80000012
        'songtest.ForeColor = &H8000E
          
        Next Tmpint
    Else
        For Tmpint = 200 To 1740 Step 70
            DoEvents
            Me.Height = Tmpint
        songtest.Enabled = False
        songtest.Visible = False
        Timer6.Enabled = False
        'songtest.BackColor = &H8000E '&H80000012
        'songtest.ForeColor = &H8000E
             
        Next Tmpint
    End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
mouseIsDown = True
    cx = x
    cy = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If mouseIsDown Then
        Move Left + (x - cx), Top + (Y - cy)
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
mouseIsDown = False
End Sub

Private Sub Image6_Click()
Unload Me
End Sub

Private Sub minimized_Click()
Dim x, Y    ' Declare variables.
'Dim frameHeight As Long
'Dim frameWidth As Long
   
    
     Form1.WindowState = vbMinimized
     
     'If vbMinimized Then
     'Timer4.Enabled = True
     'End If
     
     End Sub

Private Sub OnTop_Click()
AlwaysOnTop Form1, 1
End Sub

Private Sub pbc_Click(Index As Integer)
  Select Case Index
    Case 0: Call Eject
    Case 1: Call StopIt
    Case 2: Call PauseIt
    Case 3: Call PlayIt
    Case 4: Call ToggleIntro
    Case 5: Call ToggleSTP
    Case 6: Call PrevTrack
    Case 7: Call NextTrack
  End Select
End Sub
Private Sub Eject()
    Call GetOneFile
    If AddName <> "" Then
        SongTitle = AddTitle
        SongName = AddName
        SongPath = AddPath
        Call cClear_Click
        TNum = 1
        Playlist.AddItem AddTitle
        PlayPath.AddItem AddPath
        Call PlayIt
        Playlist.ListIndex = 0
    End If
    Exit Sub
End Sub
Sub StopIt()
    MMControl1.From = 0
    MMControl1.Command = "seek"
    MMControl1.Command = "stop"
    iSlider.Visible = False
    Label1.Caption = "00:00"
    Timer5.Enabled = False
    Call MP3Timer
    'frmVBAmp.Caption = "Mouse Player-3 version 1.5beta"
    'sStop.Visible = True: sPause.Visible = False: sPlay.Visible = False
    Paused = False: Playing = False
End Sub

Sub PauseIt()
    'If sStop.Visible = True Then Exit Sub
    If Paused = True Then
        MMControl1.Command = "play"
        
    Else
        MMControl1.Command = "stop"
               
    End If
    Paused = Not Paused
    'sStop.Visible = False
    'sPause.Visible = Paused
    'sPlay.Visible = Not Paused

End Sub
Sub QuitIt()
Call StopIt
Unload Me
End Sub
Sub NextTrack()
    TNum = TNum + 1: Call PlayIt
    End Sub
Sub PrevTrack()
    TNum = TNum - 1: Call PlayIt
End Sub
Sub PlayIt()
    'SendMCIString "stop cd wait", True
    'cmd = "seek cd to " & track
    'SendMCIString cmd, True
    'fPlaying = False
    'Update
    'Timer3.Enabled = False
    Call MP3Timer
    'test here
    
    Title = String(30, " ") + Playlist.List(TNum - 1)
    
    'end of test
    If Playlist.ListCount = 0 Then Exit Sub
    If TNum < 1 Then TNum = Playlist.ListCount
    If TNum > Playlist.ListCount Then TNum = 1
    Timer1.Enabled = False
    'cIDTag.Enabled = True LÄGG TILL DENNA SEDAN !!!!!!
    'sStop.Visible = False: sPause.Visible = False: sPlay.Visible = True
    
    'read MP3 file here and set details (simulate for now)
    SongPath = PlayPath.List(TNum - 1)
    
    If Dir$(SongPath) = "" Then
      x$ = PlayPath.List(TNum - 1)
      PlayPath.List(TNum - 1) = x$ + "*** FILE NOT FOUND ***"
      Dev$ = "": Exit Sub
    End If
    
    x$ = UCase$(Right$(SongPath, 5)): P = InStr(x$, ".")
    FileType$ = Mid$(x$, P + 1)
    
    Info = "Filename= " & SongPath 'Default info
    Dev$ = "ActiveMovie" 'default MCI device
    
    Select Case FileType$
        Case "MP3": Call GetIDTag
        Case "MOD", "MTM", "FAR", "669", "OKT", "STM", "S3M", "NST", "WOW", "XM": Dev$ = "M4W_MCI"
        Case "FLI", "FLC": Dev$ = "Animation"
        Case "AWA", "AWM": Dev$ = "Animation1"
        Case "MMM": Dev$ = "MMMovie"
        'Case "VIV": Dev$ = "MMMovie"
        'Case "CDA": Dev$ = "cdaudio" ': CDTrack = Val(Right$(SongPath, 6)) '**** TEST
        Case "BMP", "GIF", "JPG": GoSub DoBitmap: Exit Sub
    End Select
    
    'lKbps.Caption = ""
    'lKHz.Caption = ""
    'lstereo.Caption = "stereo"

    'reset elapsed slider and set titles
    iSlider.Move 2275, 2250: iSlider.Visible = True
    'sBlip.Visible = True
    lSong.Caption = Playlist.List(TNum - 1)
    'lTrack.Caption = Str$(TNum)
    'lTrack.Visible = True
    
    MMControl1.Command = "close"
    
    
    '****** CHECK WHY THE FUCK THESE LINES FUCKS UP THE PLAYING ON SOM FILES !!!!!!!!!!
    '****** Looks like it cant read some PATHS ?!? With the ' in the name..check tihs!!
    'Play the file
    MMControl1.DeviceType = Dev$ 'select MCI driver
    MMControl1.TimeFormat = 0
    '****** END OFF LINECHECK
    
    If Dev$ = "cdaudio" Then
    MMControl1.DeviceType = "CDAudio"
      Call StopIt
      'lTrack.Visible = False
      'Timer1.Enabled = False
      'Timer3.Enabled = True
      'Timer5.Enabled = False
      'Call Update
      
      MMControl1.Command = "open"
      MMControl1.Command = "play"
      'SendMCIString "play cd", True
      'fPlaying = True
            
      MMControl1.filename = ""
      'MMControl1.DeviceType = "CDAudio"
      'MMControl1.Command = "open"
      'MMControl1.Command = "play"
      'MMControl1.track = CDTrack
      PFrom! = MMControl1.To
      SongLen = MMControl1.trackLength
      MMControl1.From = PFrom!
      MMControl1.To = PFrom! + SongLen - 1
    Else
      MMControl1.filename = SongPath
      MMControl1.Command = "open"
      MMControl1.From = 0
    End If
    'Label2.Visible = False
    'cdstop_Click
    SongLen = MMControl1.trackLength ' WORKS OKI
    MMControl1.Command = "play": Playing = True ' WORKS OKI
    
    Playlist.ListIndex = TNum - 1
    Pos = SongLen / 1000
    min = Int(Pos / 60): sec = Int(Pos - min * 60)
    lTot.Caption = Format$(min, "00") + ":" + Format$(sec, "00")
    Timer1.Enabled = True
    Exit Sub
    
DoBitmap:
    Dev$ = ""
    Call LoadCover(SongPath)
    Playlist.ListIndex = TNum - 1
    Timer1.Enabled = True
    Return
      
    
End Sub
Private Sub ToggleIntro()
    Intro = Not Intro
    sIntro.Visible = Intro
End Sub
Private Sub ToggleSTP()
    STP = Not STP
    sSTP.Visible = STP
End Sub
Private Sub cAddFile_Click()
  Call AddFile
End Sub
Private Sub AddFile()
    Call GetOneFile
    If AddName <> "" Then
        If IsPic(AddPath) Then AddTitle = AddTitle + " (picture)"
        Playlist.AddItem AddTitle
        PlayPath.AddItem AddPath
        If TNum = 0 Then TNum = 1
    End If
    Call SetUpDn
End Sub
Sub SetUpDn()
    If Playlist.ListCount > 0 Then
        cUp.Enabled = True: cDn.Enabled = True: cDel.Enabled = True
    Else
        cUp.Enabled = False: cDn.Enabled = False: cDel.Enabled = True
    End If
End Sub
Private Sub cClear_Click()
    Call ClearPL
End Sub
Sub ClearPL()
    Playlist.Clear: PlayPath.Clear: TNum = 1
End Sub
Private Sub cDel_Click()
    N = Playlist.ListIndex
    If N >= 0 Then
        Playlist.RemoveItem (N)
        PlayPath.RemoveItem (N)
        If N > Playlist.ListCount - 1 Then N = Playlist.ListCount - 1
        Playlist.ListIndex = N
    End If
    Call SetUpDn
End Sub
Private Sub cUp_DblClick()
    Call MovePLe(-1)
End Sub
Private Sub cUp_Click()
    Call MovePLe(-1)
End Sub

Private Sub cDn_DblClick()
    Call MovePLe(1)
End Sub
Private Sub cDn_Click()
    Call MovePLe(1)
End Sub
Sub MovePLe(D)
    N = Playlist.ListIndex
    If (N + D) >= 0 And (N + D) < Playlist.ListCount Then
        T1$ = Playlist.List(N): T2$ = PlayPath.List(N)
        Playlist.List(N) = Playlist.List(N + D)
        PlayPath.List(N) = PlayPath.List(N + D)
        Playlist.List(N + D) = T1$
        PlayPath.List(N + D) = T2$
        Playlist.ListIndex = N + D
    End If
End Sub
Private Sub cMixer_Click()
    Call ShowMixer
End Sub
Private Sub cPlaylist_Click()
    Call TogglePL
End Sub

Private Sub TogglePL()
    PlVisFlag = 1 - PlVisFlag
    Me.Height = 1740 + PlVisFlag * 1545
End Sub
Sub lTitlebar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If DragFlag = 0 Then
        IX = x: IY = Y
        FX = Me.Left: FY = Me.Top
        DragFlag = 1
    End If
End Sub

Sub lTitlebar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If DragFlag = 1 Then
        Me.Move FX + (x - IX), FY + (Y - IY)
        FX = Me.Left: FY = Me.Top
    End If
End Sub

Sub lTitlebar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    DragFlag = 0
    
    'Call Snap2ViewPoint(Me)
End Sub

Private Sub GetOneFile()
    Static LastFilter
    If LastFilter = 0 Then LastFilter = 1
    AddName = ""
    On Error GoTo ErrHandler
    
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = DefPath
    CommonDialog1.DialogTitle = "Open Media File"
    CommonDialog1.FLAGS = cdlOFNHideReadOnly
    CommonDialog1.Filter = "MPEG Audio Files|*.MP?|ActiveMovie Files|*.MP?;*.MPEG;*.DAT;*.WAV;*.AU;*.MID;*.RMI;*.AIF?;*.MOV;*.QT;*.AVI;*.M1V;*.RA;*.RAM;*.RM;*.RMM|Music Modules|*.MOD;*.MTM;*FAR;*.669;*.OKT;*.STM;*.S3M;*.NST;*.WOW;*.XM|Bitmaps|*.BMP;*.GIF;*.JPG|All Files|*.*"
    CommonDialog1.FilterIndex = LastFilter
    CommonDialog1.ShowOpen
    LastFilter = CommonDialog1.FilterIndex
    
    AddName = CommonDialog1.FileTitle
     
    AddTitle = MakeTitle$(AddName)
    AddPath = CommonDialog1.filename
    DefPath = Left$(AddPath, Len(AddPath) - Len(AddName))
    SaveSetting PrgName, Sect, "Path", DefPath
        
ErrHandler:
    Exit Sub
End Sub

Private Sub Playlist_Click()
cDel.Enabled = True
End Sub

Private Sub Playlist_DblClick()
  TNum = Playlist.ListIndex + 1
  Call PlayIt
End Sub


Private Sub quitbutton_Click()
Unload Me
End Sub

Private Sub testpl_Click()
Call LoadPL
End Sub

Private Sub Timer1_Timer()
  
  DoEvents
  If Playing = False Then Exit Sub
  If Dev$ = "" Then TNum = TNum + 1: Call PlayIt: Exit Sub
  If SongLen > 0 Then
    Elapsed = MMControl1.Position
    Frame = Elapsed / 25: Pos = Elapsed / 1000
    min = Int(Pos / 60): sec = Int(Pos - min * 60)
    'lElapsed.Caption = Format$(min, "00") + ":" + Format$(sec, "00")
    
    'sBlip.Left = 12 + (Frame Mod 78)
    If SlideFlag = False Then
        iSlider.Left = 2275 + Int((Elapsed / SongLen) * 1900)
        End If
       
    If Intro = True Then
        If Elapsed > 10000 Then TNum = TNum + 1: Call PlayIt
    End If
    
    If Elapsed = SongLen Then
        TNum = TNum + 1: Call PlayIt
    End If
    
  Else
    If lElapsed.Caption <> ":" Then Call ClearInf
  End If

'Title = Mid(Title, 2) & Left(Title, 1)
'Label1 = Title


End Sub

Sub GetIDTag()
    If SongPath <> "" Then If Dir$(SongPath) <> "" Then GoSub GetID
    Exit Sub
        
GetID:
    Close
    Open SongPath For Binary As 1
    N& = LOF(1): If N& < 256 Then Close 1: Return
    Get #1, (N& - 256), InBuf:  Close 1
    A$ = "": Cr$ = Chr$(13)
    P = InStr(1, InBuf, "tag", 1)
    If P = 0 Then
        A$ = "No ID Tag in file!"
    Else
        A$ = A$ & Cr$ & "Title: " & Mid$(InBuf, P + 3, 30)
        A$ = A$ & Cr$ & "Artist: " & Mid$(InBuf, P + 33, 30)
        A$ = A$ & Cr$ & "Album: " & Mid$(InBuf, P + 63, 30)
        A$ = A$ & Cr$ & "Year: " & Mid$(InBuf, P + 93, 4)
        A$ = A$ & Cr$ & "Comment: " & Mid$(InBuf, P + 97, 30)
    End If
    Info = A$: A$ = ""
    Return
End Sub
Private Sub ShowInfo()
    MsgBox Info, vbInformation, "Track Info"
End Sub
Sub ClearInf()
    'lElapsed.Caption = ":": lTrack.Caption = "0"
    'lKbps.Caption = "": lKHz.Caption = ""
    'lstereo.Caption = "": lTot.Caption = ""
    'sBlip.Visible = False:
    iSlider.Visible = False
End Sub
Private Sub cLoad_Click()
    Call LoadPL
End Sub
Sub LoadPL()
    On Error GoTo ErrHandler2
    
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = DefPath
    CommonDialog1.DialogTitle = "Load Playlist"
    CommonDialog1.FLAGS = cdlOFNHideReadOnly
    CommonDialog1.Filter = "MP3 Playlists (*.M3U)|*.M3U|Playlists (*.PLS)|*.PLS"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
 
    F$ = CommonDialog1.filename
    DoEvents
    
    ReadPL (F$)
    
ErrHandler2:
    Exit Sub
End Sub
Sub ReadPL(F$)
    For J = Len(F$) To 1 Step -1
      If Mid$(F$, J, 1) = "." Then Exit For
    Next
    ff$ = Left$(F$, J - 1) 'path+filename without extension
    
    
    A$ = F$: GoSub SplitPF: P1$ = P2$ 'Get path of playlist as base for entries
        
    Open F$ For Input As 1
    Call ClearPL: N = 0
    Select Case UCase$(Right$(F$, 3))
      Case "M3U": GoSub LoadM3U
      Case "PLS": GoSub LoadPLS
    End Select
    Close 1
    TNum = 1: Call SetUpDn: Call PlayIt
    Exit Sub
    
LoadM3U:
    While Not EOF(1)
        Line Input #1, AA$: A$ = LTrim$(AA$)
        If N < 32766 Then GoSub AddIt
    Wend
    Return
  
LoadPLS:
    While Not EOF(1)
        Line Input #1, AA$: AA$ = LTrim$(A$)
        If N < 32766 And Left$(AA$, 4) = "File" Then
            i = InStr(AA$, "=")
            If i > 0 Then A$ = Mid$(AA$, i + 1): GoSub AddIt
        End If
    Wend
    Return

AddIt:
    GoSub SplitPF: N = N + 1
    x$ = MakeTitle$(B$)
    If IsPic(B$) Then x$ = x$ + " (picture)"
    Playlist.AddItem x$
    If Mid$(A$, 2, 1) = ":" Then
      PlayPath.AddItem A$
    ElseIf Left$(A$, 1) = "\" Then
      PlayPath.AddItem A$
    Else
      PlayPath.AddItem P1$ + "\" + P2$ + B$
    End If
    Return

SplitPF:
    For J = Len(A$) To 1 Step -1
        If Mid$(A$, J, 1) = "\" Then Exit For
    Next
    P2$ = Left$(A$, J)
    B$ = Mid$(A$, J + 1)
    If Left$(P2$, 1) = "\" Then P2$ = Left$(P2$, 2)
    Return

End Sub
Private Sub cSave_Click()
  Call SavePL
End Sub
Private Sub SavePL()
    On Error GoTo ErrHandler4
    
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = DefPath
    CommonDialog1.DialogTitle = "Save Playlist"
    CommonDialog1.FLAGS = cdlOFNHideReadOnly
    CommonDialog1.Filter = "MP3 Playlists (*.M3U)|*.M3U"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.filename = ""
    CommonDialog1.ShowSave
 
    F$ = CommonDialog1.filename
    Call WritePL(F$)
    
ErrHandler4:
    Exit Sub
End Sub
Sub WritePL(F$)
    On Error GoTo ErrHandler5
    Open F$ For Output As 1
        
    For J = 1 To PlayPath.ListCount
        Print #1, PlayPath.List(J - 1)
    Next
    Close 1: Exit Sub
ErrHandler5:
    Close
    MsgBox "Unable to write Playlist!"
    Exit Sub

End Sub
Private Sub iSlider_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SlideFlag = False Then
        IX = x: FX = iSlider.Left
        TX = Screen.TwipsPerPixelX
        SlideFlag = True
    End If
End Sub

Private Sub iSlider_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SlideFlag = True Then
        Pos = FX + (x - IX) / TX
        If Pos < 2275 Then Pos = 2275
        If Pos > 4000 Then Pos = 4000
        FX = Pos: iSlider.Left = Pos
    End If
End Sub

Private Sub iSlider_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Kolla vad i helvete den gör här !!!
    
    P! = Int(((iSlider.Left - 2275) / 1900) * SongLen)
    MMControl1.Command = "stop"
    MMControl1.From = P!
    MMControl1.Command = "play"
    SlideFlag = False
End Sub
Private Sub cIDTag_Click()
    Call ShowInfo
End Sub

Private Sub cSkin_Click()
    Call LoadSkin
End Sub
Private Sub LoadSkin()

    On Error GoTo ErrHandler3
    
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Select Skin"
    CommonDialog1.FLAGS = cdlOFNHideReadOnly
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Skin control file (*.skin)|*.skin"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    
    F$ = CommonDialog1.filename
    
    'load colours then bitmaps
    Open F$ For Input As 1
    GoSub GetColour: GoSub SetText
    GoSub GetColour: Playlist.ForeColor = v
    GoSub GetColour: Playlist.BackColor = v
    GoSub GetColour: sBlip.FillColor = v
    
    Timer1.Enabled = False
    Input #1, A$: Me.Picture = LoadPicture(A$)
    Input #1, A$: iSlider.Picture = LoadPicture(A$)
    Close
    
ErrHandler3:
    Timer1.Enabled = True
    Exit Sub
    
GetColour:
  Input #1, r, g, B
  v = (r * 65536) + (g * 256) + B: Return

SetText:
  lSong.ForeColor = v
  lTrack.ForeColor = v
  lElapsed.ForeColor = v
  lTot.ForeColor = v
  lKbps.ForeColor = v
  lKHz.ForeColor = v
  lstereo.ForeColor = v
  Return
  
End Sub

Private Sub cAddDir_Click()
    'Call AddDir
MsgBox "Option not added yet!"
End Sub
Private Sub AddDir()

End Sub
Private Function MakeTitle$(A$)
  P = Len(A$): l = P + 1
  For J = 1 To P
    T$ = Mid$(A$, J, 1)
    If T$ = "_" Then Mid$(A$, J, 1) = " "
    If T$ = "." Then l = J
  Next J
  
  MakeTitle$ = Left$(A$, l - 1)
  
End Function
Sub LoadCover(A$)
  C$ = ""
  If InStr(A$, ".") = 0 Then
    'filename without extension, so try different types
    Ext$ = ".BMP": GoSub TestIt
    Ext$ = ".GIF": GoSub TestIt
    Ext$ = ".JPG": GoSub TestIt
  Else
    C$ = A$ 'full filename, so just use it
  End If
  
  If C$ <> "" Then
    frmAlbum.Visible = True
    frmAlbum.Cover.Picture = LoadPicture(C$)
  Else
    frmAlbum.Cover.Picture = Nothing
    frmAlbum.Visible = False
  End If
  Exit Sub
  
TestIt:
 If Dir$(A$ + Ext$) <> "" Then C$ = A$ + Ext$
 Return
 
End Sub
Sub ShowPrefs()
    If frmIcon.Visible = False Then
      Call AlwaysOnTop(Me, False)
      frmOptions.Show 1
    End If
End Sub
    
Sub ShowMixer()
    Shell "sndvol32.exe", vbNormalFocus
End Sub
Function IsPic(A$) As Boolean

    x$ = UCase$(Right$(A$, 4))
    If x$ = ".BMP" Or x$ = ".GIF" Or x$ = ".JPG" Then
      IsPic = True
    Else
      IsPic = False
    End If
End Function

Private Sub Timer2_Timer()
Title = Mid(Title, 2) & Left(Title, 1)
lSong = Title
End Sub


Sub MP3Timer()
'Timer4.Enabled = False
Timer2.Enabled = False
Timer5.Enabled = True

End Sub


Private Sub Timer4_Timer()
'frmVBAmp.Caption = " FILE : " & "[" & Str$(TNum) & "]  " & Format$(min, "00") + ":" + Format$(sec, "00") _
            & " : " & Playlist.List(TNum - 1)

End Sub

Private Sub Timer5_Timer()
'lTrack.Visible = False
'lTot.Visible = True


Label1.Caption = " " & Format$(min, "00") + ":" + Format$(sec, "00") _
            '& " : " & Playlist.List(TNum - 1)
            
'Label1.Caption = " & " Format$(min, "00") + ":" + Format$(sec, "00") _
            '& " : " & Playlist.List(TNum - 1)
            
Title = Mid(Title, 2) & Left(Title, 1)
lSong.Caption = Playlist.List(TNum - 1)
lSong = Title

End Sub

Private Sub Timer6_Timer()
Title = Mid(Title, 2) & Left(Title, 1)
songtest.Caption = Playlist.List(TNum - 1)
songtest = Title
End Sub





