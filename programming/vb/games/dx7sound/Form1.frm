VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRndPanRndSnd 
      Caption         =   "RS, RP, RB"
      Height          =   315
      Left            =   7560
      TabIndex        =   28
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdRndPan 
      Caption         =   "RP, RB"
      Height          =   315
      Left            =   7560
      TabIndex        =   27
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayRndSound 
      Caption         =   "RS, RB"
      Height          =   315
      Left            =   3120
      TabIndex        =   25
      Top             =   2460
      Width           =   1275
   End
   Begin VB.CommandButton cmdAnyBuffer 
      Caption         =   "RB"
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   24
      Top             =   1020
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Index           =   3
      ItemData        =   "Form1.frx":0000
      Left            =   4500
      List            =   "Form1.frx":0016
      TabIndex        =   21
      Top             =   1560
      Width           =   1755
   End
   Begin VB.HScrollBar hsbVolume 
      Height          =   195
      Index           =   3
      LargeChange     =   5
      Left            =   6360
      Max             =   100
      TabIndex        =   20
      Top             =   1800
      Value           =   50
      Width           =   2415
   End
   Begin VB.HScrollBar hsbPan 
      Height          =   195
      Index           =   3
      LargeChange     =   5
      Left            =   6360
      Max             =   100
      TabIndex        =   19
      Top             =   2220
      Value           =   50
      Width           =   2415
   End
   Begin VB.CommandButton cmdBuffer 
      Caption         =   "Play Buffer 3"
      Height          =   315
      Index           =   3
      Left            =   6360
      TabIndex        =   18
      Top             =   2460
      Width           =   1155
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Index           =   2
      ItemData        =   "Form1.frx":005E
      Left            =   4500
      List            =   "Form1.frx":0074
      TabIndex        =   15
      Top             =   120
      Width           =   1755
   End
   Begin VB.HScrollBar hsbVolume 
      Height          =   195
      Index           =   2
      LargeChange     =   5
      Left            =   6360
      Max             =   100
      TabIndex        =   14
      Top             =   360
      Value           =   50
      Width           =   2415
   End
   Begin VB.HScrollBar hsbPan 
      Height          =   195
      Index           =   2
      LargeChange     =   5
      Left            =   6360
      Max             =   100
      TabIndex        =   13
      Top             =   780
      Value           =   50
      Width           =   2415
   End
   Begin VB.CommandButton cmdBuffer 
      Caption         =   "Play Buffer 2"
      Height          =   315
      Index           =   2
      Left            =   6360
      TabIndex        =   12
      Top             =   1020
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuffer 
      Caption         =   "Play Buffer 1"
      Height          =   315
      Index           =   1
      Left            =   1980
      TabIndex        =   11
      Top             =   2460
      Width           =   1095
   End
   Begin VB.HScrollBar hsbPan 
      Height          =   195
      Index           =   1
      LargeChange     =   5
      Left            =   1980
      Max             =   100
      TabIndex        =   8
      Top             =   2220
      Value           =   50
      Width           =   2415
   End
   Begin VB.HScrollBar hsbVolume 
      Height          =   195
      Index           =   1
      LargeChange     =   5
      Left            =   1980
      Max             =   100
      TabIndex        =   7
      Top             =   1800
      Value           =   50
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Index           =   1
      ItemData        =   "Form1.frx":00BC
      Left            =   120
      List            =   "Form1.frx":00D2
      TabIndex        =   6
      Top             =   1560
      Width           =   1755
   End
   Begin VB.CommandButton cmdBuffer 
      Caption         =   "Play Buffer 0"
      Height          =   315
      Index           =   0
      Left            =   1980
      TabIndex        =   5
      Top             =   1020
      Width           =   1095
   End
   Begin VB.HScrollBar hsbPan 
      Height          =   195
      Index           =   0
      LargeChange     =   5
      Left            =   1980
      Max             =   100
      TabIndex        =   2
      Top             =   780
      Value           =   50
      Width           =   2415
   End
   Begin VB.HScrollBar hsbVolume 
      Height          =   195
      Index           =   0
      LargeChange     =   5
      Left            =   1980
      Max             =   100
      TabIndex        =   1
      Top             =   360
      Value           =   50
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Index           =   0
      ItemData        =   "Form1.frx":011A
      Left            =   120
      List            =   "Form1.frx":0130
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "RS = Random Sound :      RB  = Random Buffer :     RP = Random Pan"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Width           =   8775
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   2040
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   7560
      X2              =   7560
      Y1              =   600
      Y2              =   720
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3180
      X2              =   3180
      Y1              =   2040
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3180
      X2              =   3180
      Y1              =   600
      Y2              =   720
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume"
      Height          =   195
      Index           =   3
      Left            =   6360
      TabIndex        =   23
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Pan"
      Height          =   195
      Index           =   3
      Left            =   6360
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume"
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Pan"
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   16
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Pan"
      Height          =   195
      Index           =   1
      Left            =   1980
      TabIndex        =   10
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume"
      Height          =   195
      Index           =   1
      Left            =   1980
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Pan"
      Height          =   195
      Index           =   0
      Left            =   1980
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume"
      Height          =   195
      Index           =   0
      Left            =   1980
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Code and Module was created by D.R Hall
'For more Information and latest version
'E-mail me, derek.hall@virgin.net

'This is a demo of how to use my module
'to make playing DirectX7 Sound easier

Option Explicit
Const cSoundBuffers = 10

Private Sub cmdAnyBuffer_Click(Index As Integer)
  PlaySoundAnyBuffer List1(Index), hsbVolume(Index).Value, hsbPan(Index).Value
End Sub

Private Sub cmdBuffer_Click(Index As Integer)
  PlaySoundWithPan Index, List1(Index), hsbVolume(Index).Value, hsbPan(Index).Value ' Play the buffer for this index
End Sub

Private Sub cmdPlayRndSound_Click()
  List1(1).ListIndex = (5 * Rnd)
  PlaySoundAnyBuffer List1(1), hsbVolume(1).Value, hsbPan(1).Value
End Sub

Private Sub cmdRndPan_Click()
  hsbPan(2).Value = (Rnd * 100) '+ 25
  PlaySoundAnyBuffer List1(2), hsbVolume(2).Value, hsbPan(2).Value
End Sub

Private Sub cmdRndPanRndSnd_Click()
  List1(3).ListIndex = (5 * Rnd)
  hsbPan(3).Value = (Rnd * 100) '+ 25
  PlaySoundAnyBuffer List1(3), hsbVolume(3).Value, hsbPan(3).Value
End Sub

Private Sub Form_Load()
'This Code was created by D.R Hall
'For more Information and latest version
'E-mail me, derek.hall@virgin.net
'To set up the DX7 sound module, just call these routines 3 routines
  
  SetupDX7Sound Me              ' Assign DX7 to this Application
  SoundDir App.Path & "\Sound"  'Where is the applications sound stored

  CreateBuffers cSoundBuffers, "Default.wav" ' How many Channels/buffers (I used 10)
                                              'and assign a default sound,
                                              'change sound later.
                                              'To stop errors you must supply
                                              'a default wave to set up a buffer
                                              'select your smallest wave file
'**** Make a selection for each listbox on this form
  List1(0).ListIndex = 4
  List1(1).ListIndex = 5
  List1(2).ListIndex = 0
  List1(3).ListIndex = 2
'**************************************
End Sub

Private Sub hsbPan_Change(Index As Integer)
  PanSound Index, hsbPan(Index).Value 'value must be 0 to 100, 50 is centered
End Sub

Private Sub hsbVolume_Change(Index As Integer)
  VolumeLevel Index, hsbVolume(Index).Value 'value must be 0 to 100, 0 no sound,
End Sub


