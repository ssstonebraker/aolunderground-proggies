Attribute VB_Name = "Module1"
Global usermode As String 'sets usermode host or client
Global multiplayermode As Boolean 'Sets multiplayer yes no
Global MyTurn As Boolean 'My turn switch
Global profilename As Variant 'name for your machine
Global opponentsname As Variant 'name for remote machine
Global score As Integer ' keeps track of game score
Global profilenamescore As Integer 'your score
Global opponentsscore As Integer 'remote score
Global sw As Boolean 'set whether x or o goes first
' Constants
Public Const MaxPlayers = 2
Public Const MChatString = 60

' DirectPlay stuff
Public dx7 As New DirectX7
Public dxplay As DirectPlay4
Public EnumConnect As DirectPlayEnumConnections
Public onconnect As Boolean
Public gNumPlayersWaiting As Byte
Public MyPlayer As Long
Public EnumSession As DirectPlayEnumSessions
Public numplayers As Byte
Public dxHost As Boolean
Public CurrentPlayer As Integer
Public PlayerScores(MaxPlayers) As Byte
Public PlayerIDs(MaxPlayers) As Long
Public dxMyTurn As Integer
Public GameUnderway As Boolean
Public connectionmade As Boolean

'The appguid number was generated with the utility provide with DX7 SDK.
Public Const AppGuid = "{D4D5D10B-7D04-11D3-8E64-00A0C9E01368}"

'This defines the msgtype you will send with DXplay.send
Public Enum MSGTYPES
    MSG_STOP 'Handles user diconnect
    MSG_STARTGAME 'Startgame
    MSG_CHAT_ON 'Chat on or off
    MSG_CHAT    'chat input
    MSG_RESTART 'Restart Game
    MSG_XORO 'Select if X or O Starts game
    MSG_MOVE 'What square selected
End Enum
Public Sub CloseDownDPlay() 'this shuts down directplay
  dxHost = False
  GameUnderway = False
  Set EnumConnect = Nothing
  Set EnumSession = Nothing
  Set dxplay = Nothing
End Sub
' Main procedure. This is where we poll for DirectPlay messages in idle time.
Public Sub Main()
MainBoard.Show
  Do While DoEvents()  ' allow event processing while any windows open
    DPInput
  Loop
End Sub
' Receive and process DirectPlay Messages
Public Sub DPInput()
  Dim FromPlayer As Long
  Dim ToPlayer As Long
  Dim msgsize As Long
  Dim msgtype As Long
  Dim dpmsg As DirectPlayMessage
  Dim MsgCount As Long
  Dim msgdata() As Byte
  Dim x As Integer
  Dim fromplayername As String
     
  If dxplay Is Nothing Then Exit Sub 'IF  single player then exit
    
  On Error GoTo NOMESSAGE
  ' If this call fails, presumably it's because there's no session or
  ' no player.
  MsgCount = dxplay.GetMessageCount(MyPlayer) 'Get number of messages.
  On Error GoTo MSGERROR
  Do While MsgCount > 0 'Read all messages
    Set dpmsg = dxplay.Receive(FromPlayer, ToPlayer, DPRECEIVE_ALL) 'Read DXINput
    msgtype = dpmsg.ReadLong() 'Read DXinput msg TYPE
    MsgCount = MsgCount - 1
    'Direct X System Only Messages not user defineable
    If FromPlayer = DPID_SYSMSG Then
    
      Select Case msgtype
      ' New player, update player list
        Case DPSYS_DESTROYPLAYERORGROUP, _
             DPSYS_CREATEPLAYERORGROUP
         
          If Connect.Visible Then Connect.UpdateWaiting 'update connection sessions list
          
          
          Case DPSYS_HOST 'either lost connection or changed you to host
            dxHost = True
             If Connect.Visible Then
               MsgBox ("You are now the host.")
               Connect.UpdateWaiting   ' make sure Start button is enabled
            End If
          
      End Select
' ---------------------------------------------------------------------------------------
    
    ' User specified Message Structure TYPES
    
    Else
    
      ' Get name of sending player
      If onconnect = False Then
      fromplayername = dxplay.GetPlayerFriendlyName(FromPlayer) 'Gets name
      opponentsname = fromplayername 'changes to games variable
            'Updates status bars and labels.
            If usermode = "host" Then
                MainBoard.playerdisplaylabel.Caption = opponentsname & " Has Joined The Game"
                MainBoard.StatusBar1.SimpleText = opponentsname & "Is Ready To Play,  Start Game"
            End If
            If usermode = "client" Then
                MainBoard.playerdisplaylabel.Caption = "You Have Joined " & opponentsname & "'s Game"
                MainBoard.StatusBar1.SimpleText = opponentsname & " Will Start The Game"
            End If
         End If
         onconnect = True
         Select Case msgtype
     'Below is where you define your message structure types and add responding code, cool.
     Case MSG_STARTGAME
        onconnect = True
          multiplayermode = True
          ' Number of players
          numplayers = dpmsg.ReadByte
          ' Player IDs,
            MyPlayer = dpmsg.ReadLong
          ' Show the game board.
            Connect.Hide
            MainBoard.Enabled = True
            MainBoard.Show
            MainBoard.hostagame.Enabled = False
            MainBoard.joinagame.Enabled = False
            MainBoard.mnudisconnect.Enabled = True
        
     Case MSG_MOVE 'Sent when square is click
            Dim t As Byte
                t = dpmsg.ReadByte
                
        Select Case t
            Case 0
                Call MainBoard.layer_A_online(0)
            Case 1
                Call MainBoard.layer_A_online(1)
            Case 2
                Call MainBoard.layer_A_online(2)
            Case 3
                Call MainBoard.layer_A_online(3)
            Case 4
                Call MainBoard.layer_A_online(4)
            Case 5
                Call MainBoard.layer_A_online(5)
            Case 6
                Call MainBoard.layer_A_online(6)
            Case 7
                Call MainBoard.layer_A_online(7)
            Case 8
                Call MainBoard.layer_A_online(8)
            End Select
     MyTurn = True
     
       Case MSG_CHAT_ON          'Handles Turn chat on off
            Call MainBoard.chatswitch
         
        Case MSG_XORO 'Selects who goes first X or O
        Dim thing As Byte
        thing = dpmsg.ReadByte
        If thing = 1 Then
            Call MainBoard.x_Click
        End If
        If thing = 2 Then
            Call MainBoard.o_Click
        End If
     
        Case MSG_RESTART 'handles input for restart
               multiplayermode = True
                MainBoard.playerdisplaylabel.Caption = opponentsname & " has restarted the game."
                        If sw = True Then
                            MyTurn = False
                        Else
                            MyTurn = True
                        End If
                         Call MainBoard.restart_Click
                        
        Case MSG_CHAT 'Handles Chat String input
          Dim chatin As String
                chatin = dpmsg.ReadString()
               If MainBoard.chatlabel.Text = "" Then
                 MainBoard.chatlabel.Text = opponentsname & ": " & chatin
                    Else
                 MainBoard.chatlabel.Text = MainBoard.chatlabel.Text & vbCrLf & opponentsname & ": " & chatin
                End If
                
         Case MSG_STOP 'Handles player disconnected.
          MsgBox opponentsname & " has left the game.", vbOKOnly, "Tic Tac Oops"
                MainBoard.mnudisconnect.Enabled = False
                MainBoard.newgame.Enabled = True
                MainBoard.hostagame.Enabled = True
                MainBoard.joinagame.Enabled = True
                multiplayermode = False
                usermode = "host"
                Call CloseDownDPlay
                Unload Connect
                onconnect = False
    End Select
   
End If
    
  Loop
  Exit Sub
  
' Error handlers
MSGERROR:
  MsgBox ("Error reading message.")
  CloseDownDPlay
  End
NOMESSAGE:
  Exit Sub
End Sub


