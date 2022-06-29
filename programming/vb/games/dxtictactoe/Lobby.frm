VERSION 5.00
Begin VB.Form Lobby 
   Caption         =   "Tic Tac Lobby"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox GameName 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton startgame 
      Caption         =   "Start Game"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton join 
      Caption         =   "Join"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox playerlist 
      Height          =   1230
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Lobby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Hide
 ' Unload frmMultiplayer  ' Force reset of DPlay
  'frmMainMenu.Show
End Sub

Private Sub cmdRefresh_Click()
    Screen.MousePointer = vbHourglass
    UpdateSessionList
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
  cmdCancel_Click
End Sub

Private Sub cmdCreate_Click()
  Hide
 ' frmCreateGame.Show
    ' Put focus on name
  'frmCreateGame.txtGameName.SetFocus

End Sub

' Join game

Private Sub cmdJoin_Click()

  Dim SessionData As DirectPlaySessionData
  
  ' Join selected session
  'Set SessionData = gObjDPEnumSessions.GetItem(lstSessions.ListIndex + 1)
  On Error GoTo NOSESSION
 ' Call gObjDPlay.Open(SessionData, DPOPEN_JOIN)
  On Error GoTo 0
  
  ' Create player
  Dim PlayerHandle As String
  Dim PlayerName As String
  
  'PlayerName = frmMultiplayer.txtYourName.Text
  PlayerHandle = "Player"    ' We don't use this
  
  ' Create the player
  On Error GoTo FAILEDCREATE
  'gMyPlayerID = gObjDPlay.CreatePlayer(PlayerName, PlayerHandle, 0, 0)
  On Error GoTo 0
  
  ' Only host can start game, so disable Start button
  'frmWaiting.cmdStart.Enabled = False
  
  ' Hide this and show status window
  Hide
  'frmWaiting.Show
  'frmWaiting.UpdateWaiting  ' show the player list
  Exit Sub

' Error handlers

NOSESSION:
  MsgBox ("Failed to join game.")
  UpdateSessionList
  Exit Sub
  
FAILEDCREATE:
  MsgBox ("Failed to create player.")
  Exit Sub
  
End Sub


' Update session listbox.

Public Function UpdateSessionList() As Boolean
  
  Dim SessionCount As Integer, x As Integer
  Dim SessionData As DirectPlaySessionData
  Dim Details As String
  
  ' Delete the old list
  playerlist.Clear
  
  Set SessionData = dxplay.CreateSessionData
  
  ' Enumerate the sessions in synchronous mode.
  Call SessionData.SetGuidApplication(AppGuid)
  Call SessionData.SetSessionPassword("")
  
  If usermode = "host" Then
  Set SessionData = dxplay.CreateSessionData
' Finish describing the session
  Call SessionData.SetSessionName(GameName.Text)
  Call SessionData.SetGuidApplication(AppGuid)
  Call SessionData.SetFlags(DPSESSION_MIGRATEHOST)
  
  ' Create (and join) the session.
  
  ' Failure can result from the user cancelling out of the service provider dialog.
  ' In the case of the modem, this is the "answer" dialog.
  Call dxplay.Open(SessionData, DPOPEN_CREATE)
  
  ' Describe the host player
  Dim PlayerHandle As String
  Dim PlayerName As String
  
  PlayerName = connectfrm.PlayersName.Text
  PlayerHandle = "Player 1 (Host)"
  
  ' Create the host player
  MyPlayer = dxplay.CreatePlayer(PlayerName, PlayerHandle, 0, 0)
  On Error GoTo 0
    
  dxHost = True
  End If
  
  
  
  
  
  
  
  
  On Error GoTo USERCANCEL
  Set EnumSessions = dxplay.GetDPEnumSessions(SessionData, 0, _
          DPENUMSESSIONS_AVAILABLE)
  On Error GoTo 0
 
  ' List info for enumerated sessions: name, players, max. players
  
  On Error GoTo ENUM_ERROR
  SessionCount = EnumSessions.GetCount
  For x = 1 To SessionCount
    Set SessionData = EnumSessions.GetItem(x)
    Details = SessionData.GetSessionName & " (" & SessionData.GetCurrentPlayers _
            & "/" & SessionData.GetMaxPlayers & ")"
    playerlist.AddItem (Details)
  Next x
  
  ' Update Join button
  If SessionCount = 0 Then
   join.Enabled = False
  Else
    join.Enabled = True
  End If
  
  ' Initialize selection
  If playerlist.ListCount > 0 Then playerlist.ListIndex = 0
  
  UpdateSessionList = True
  Exit Function
  
  ' Error handlers
  ' User cancelled out of service provider dialog, e.g. for modem connection.
  ' We can't enumerate sessions but the user can still host a game.
USERCANCEL:
  UpdateSessionList = False
  join.Enabled = False
  Exit Function
  
ENUM_ERROR:
  UpdateSessionList = False
  MsgBox ("Error in enumeration functions.")
  Exit Function
  
End Function
Private Sub startgame_Click()
 Dim ConnectionMade As Boolean
Dim dxAddress As DirectPlayAddress
 Dim SessionData As DirectPlaySessionData
Set SessionData = dxplay.CreateSessionData
' Finish describing the session
  Call SessionData.SetSessionName(GameName.Text)
  Call SessionData.SetGuidApplication(AppGuid)
  Call SessionData.SetFlags(DPSESSION_MIGRATEHOST)
  
  ' Create (and join) the session.
  Dim cindex As Long
  
  ' Failure can result from the user cancelling out of the service provider dialog.
  ' In the case of the modem, this is the "answer" dialog.
  cindex = connectfrm.connectiontype.ListIndex + 1
  Set dxAddress = EnumConnections.GetAddress(cindex)
  Call dxplay.InitializeConnection(dxAddress)
  ConnectionMade = Lobby.UpdateSessionList
  If ConnectionMade Then
    Hide
    Lobby.Show
  Else
    InitDPlay
  End If
  Exit Sub
  ' Error handlers
INITIALIZEFAILED:
  If Err.Number <> DPERR_ALREADYINITIALIZED Then
    MsgBox ("Failed to initialize connection.")
    Exit Sub
  End If
  
  'Call dxplay.Open(SessionData, DPOPEN_CREATE)
  ' Describe the host player
  Dim PlayerHandle As String
  Dim PlayerName As String
  
  PlayerName = connectfrm.PlayersName.Text
  PlayerHandle = "Player 1 (Host)"
  
  ' Create the host player
  MyPlayer = dxplay.CreatePlayer(PlayerName, PlayerHandle, 0, 0)
  On Error GoTo 0
    
  dxHost = True
End Sub
