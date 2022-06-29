VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectX 7 for VB: Direct input (Keyboard)"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGame 
      BackColor       =   &H00000000&
      Height          =   2310
      Left            =   3480
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   600
      Width           =   2310
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H00FF0000&
         FillStyle       =   7  'Diagonal Cross
         Height          =   750
         Left            =   750
         Top             =   750
         Width           =   750
      End
   End
   Begin VB.Timer tmrKey 
      Left            =   120
      Top             =   1560
   End
   Begin VB.ListBox lstKeys 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Page, with game: HTTP://www.parkstonemot.freeserve.co.uk/indexFW.htm"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   5685
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMail: JollyJeffers@GreenOnions.NetscapeOnline.Co.Uk"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4020
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Try these Contacts:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1380
   End
   Begin VB.Label lblmisc 
      BackStyle       =   0  'Transparent
      Caption         =   "The picturebox to the side responds to UP/DOWN/LEFT/RIGHT. It shows how you could create a simple game...."
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Keys:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dx As New DirectX7  'the directX object.
Dim di As DirectInput   'the directInput object.
Dim diDEV As DirectInputDevice  'the sub device of DirectInput.
Dim diState As DIKEYBOARDSTATE  'the key states.
Dim iKeyCounter As Integer
Dim aKeys(255) As String    'key names


Private Sub Form_Load()
    Set di = dx.DirectInputCreate() 'create the object, must be done before anything else
    If Err.Number <> 0 Then 'if err=0 then there are no errors.
        MsgBox "Error starting Direct Input, please make sure you have DirectX installed", vbApplicationModal
        End
    End If
    Set diDEV = di.CreateDevice("GUID_SysKeyboard") 'Create a keyboard object off the Input object
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD 'specify it as a normal keyboard, not mouse or joystick
    diDEV.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    ' ^ set coop level. Defines how it interacts with other applications, whether it will share with other
    ' apps. DISCL_NONEXCLUSIVE means that it's multi-tasking friendly
    Me.Show 'show the form
    diDEV.Acquire   'aquire the keystates.
    tmrKey.Interval = 10    'sensitivity, in this case the repeat rate of the keyboard
    tmrKey.Enabled = True   'enable the timer, this has the key detecting code in it
End Sub

Private Sub Form_Unload(Cancel As Integer)
    diDEV.Unacquire
End Sub

Private Sub lstKeys_GotFocus()
frmMain.SetFocus
End Sub


Private Sub lstKeys_Scroll()
frmMain.SetFocus
End Sub


Private Sub tmrKey_Timer()
    diDEV.GetDeviceStateKeyboard diState    'get all the key states.
    For iKeyCounter = 0 To 255  ' goes through all the 255 differnet keys.
        If diState.Key(iKeyCounter) <> 0 Then   'if it =0 then it's not pressed. Anything else means it is pressed
        lstKeys.Clear
            lstKeys.AddItem KeyNames(iKeyCounter), 0    'add an item to the top of the list
        End If
    Next
    DoEvents    'doevents. Lets windows do anything it needs to do. Required
    'otherwise you can get it doing more things than it's capable of.
'This stuff is for the little game window:
'200=up
'203=left
'205=right
'208=down
If diState.Key(200) <> 0 Then
    If Shape1.Top > 0 Then
    Shape1.Top = Shape1.Top - 50
    End If
End If
If diState.Key(208) <> 0 Then
    If Shape1.Top < 100 Then
    Shape1.Top = Shape1.Top + 50
    End If
End If
If diState.Key(203) <> 0 Then
    If Shape1.Left > 0 Then
    Shape1.Left = Shape1.Left - 50
    End If
End If
If diState.Key(205) <> 0 Then
    If Shape1.Left < 100 Then
    Shape1.Left = Shape1.Left + 50
    End If
End If
End Sub





Function KeyNames(iNum As Integer) As String
'DIK=DirectInputKey, just a prefix, not actually required.
'Each key has a number, like the "Keycode=VbKeyEnter or chr$(13)"
    aKeys(1) = "DIK_ESCAPE"
    aKeys(2) = "DIK_1  On main keyboard"
    aKeys(3) = "DIK_2  On main keyboard"
    aKeys(4) = "DIK_3  On main keyboard"
    aKeys(5) = "DIK_4  On main keyboard"
    aKeys(6) = "DIK_5  On main keyboard"
    aKeys(7) = "DIK_6  On main keyboard"
    aKeys(8) = "DIK_7  On main keyboard"
    aKeys(9) = "DIK_8  On main keyboard"
    aKeys(10) = "DIK_9  On main keyboard"
    aKeys(11) = "DIK_0  On main keyboard"
    aKeys(12) = "DIK_MINUS  On main keyboard"
    aKeys(13) = "DIK_EQUALS  On main keyboard"
    aKeys(14) = "DIK_BACK BACKSPACE"
    aKeys(15) = "DIK_TAB"
    aKeys(16) = "DIK_Q"
    aKeys(17) = "DIK_W"
    aKeys(18) = "DIK_E"
    aKeys(19) = "DIK_R"
    aKeys(20) = "DIK_T"
    aKeys(21) = "DIK_Y"
    aKeys(22) = "DIK_U"
    aKeys(23) = "DIK_I"
    aKeys(24) = "DIK_O"
    aKeys(25) = "DIK_P"
    aKeys(26) = "DIK_LBRACKET  ["
    aKeys(27) = "DIK_RBRACKET  ]"
    aKeys(28) = "DIK_RETURN  ENTER on main keyboard"
    aKeys(29) = "DIK_LCONTROL  Left CTRL Key"
    aKeys(30) = "DIK_A"
    aKeys(31) = "DIK_S"
    aKeys(32) = "DIK_D"
    aKeys(33) = "DIK_F"
    aKeys(34) = "DIK_G"
    aKeys(35) = "DIK_H"
    aKeys(36) = "DIK_J"
    aKeys(37) = "DIK_K"
    aKeys(38) = "DIK_L"
    aKeys(39) = "DIK_SEMICOLON"
    aKeys(40) = "DIK_APOSTROPHE"
    aKeys(41) = "DIK_GRAVE  Grave accent (`)"
    aKeys(42) = "DIK_LSHIFT  Left SHIFT"
    aKeys(43) = "DIK_BACKSLASH"
    aKeys(44) = "DIK_Z"
    aKeys(45) = "DIK_X"
    aKeys(46) = "DIK_C"
    aKeys(47) = "DIK_V"
    aKeys(48) = "DIK_B"
    aKeys(49) = "DIK_N"
    aKeys(50) = "DIK_M"
    aKeys(51) = "DIK_COMMA"
    aKeys(52) = "DIK_PERIOD  On main keyboard"
    aKeys(53) = "DIK_SLASH  Forward slash (/)on main keyboard"
    aKeys(54) = "DIK_RSHIFT  Right SHIFT"
    aKeys(55) = "DIK_MULTIPLY  Asterisk on numeric keypad"
    aKeys(56) = "DIK_LMENU  Left ALT"
    aKeys(57) = "DIK_SPACE Spacebar"
    aKeys(58) = "DIK_CAPITAL  CAPS LOCK"
    aKeys(59) = "DIK_F1"
    aKeys(60) = "DIK_F2"
    aKeys(61) = "DIK_F3"
    aKeys(62) = "DIK_F4"
    aKeys(63) = "DIK_F5"
    aKeys(64) = "DIK_F6"
    aKeys(65) = "DIK_F7"
    aKeys(66) = "DIK_F8"
    aKeys(67) = "DIK_F9"
    aKeys(68) = "DIK_F10"
    aKeys(69) = "vDIK_NUMLOCK"
    aKeys(70) = "DIK_SCROLL  SCROLL LOCK"
    aKeys(71) = "DIK_NUMPAD7"
    aKeys(72) = "DIK_NUMPAD8"
    aKeys(73) = "DIK_NUMPAD9"
    aKeys(74) = "DIK_SUBTRACT  Hyphen (minus sign) on numeric keypad"
    aKeys(75) = "DIK_NUMPAD4"
    aKeys(76) = "DIK_NUMPAD5"
    aKeys(77) = "DIK_NUMPAD6"
    aKeys(78) = "DIK_ADD  Plus sign on numeric keypad"
    aKeys(79) = "DIK_NUMPAD1"
    aKeys(80) = "DIK_NUMPAD2"
    aKeys(81) = "DIK_NUMPAD3"
    aKeys(82) = "DIK_NUMPAD0"
    aKeys(83) = "DIK_DECIMAL  Period (decimal point) on numeric keypad"
    aKeys(87) = "DIK_F11"
    aKeys(88) = "DIK_F12"
    aKeys(86) = "DIK_F13"
    aKeys(84) = "DIK_F14"
    aKeys(85) = "DIK_F15"
    aKeys(156) = "DIK_NUMPADENTER"
    aKeys(157) = "DIK_RCONTROL  Right CTRL key"
    aKeys(91) = "DIK_NUMPADCOMMA Comma on NEC PC98 numeric keypad"
    aKeys(181) = "DIK_DIVIDE  Forward slash (/)on numeric keypad"
    aKeys(183) = "DIK_SYSRQ"
    aKeys(184) = "DIK_RMENU  Right ALT"
    aKeys(199) = "DIK_HOME"
    aKeys(200) = "DIK_UP  Up arrow"
    aKeys(201) = "DIK_PRIOR  PAGE UP"
    aKeys(203) = "DIK_LEFT  Left arrow"
    aKeys(205) = "DIK_RIGHT  Right arrow"
    aKeys(207) = "DIK_END"
    aKeys(208) = "DIK_DOWN  Down arrow"
    aKeys(209) = "DIK_NEXT  PAGE DOWN"
    aKeys(210) = "DIK_INSERT"
    aKeys(211) = "DIK_DELETE"
    aKeys(219) = "DIK_LWIN  Left Windows key"
    aKeys(220) = "DIK_RWIN  Right Windows key"
    aKeys(221) = "DIK_APPS  Application key"
    aKeys(116) = "DIK_PAUSE"

    KeyNames = aKeys(iNum)

End Function

