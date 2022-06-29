VERSION 5.00
Begin VB.Form Connect 
   Caption         =   "Connections"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form2"
   ScaleHeight     =   4740
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Connections"
      Height          =   2655
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton join 
         Caption         =   "Find a Game"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton host 
         Caption         =   "Host"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton lantype 
         Caption         =   "Modem Play"
         Height          =   495
         Index           =   2
         Left            =   1680
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton lantype 
         Caption         =   "IP/TCP or Internet"
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton lantype 
         Caption         =   "IPX Connection"
         Height          =   495
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select Connection Type and Establish a Conneciton"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.ListBox lstPlayers 
      Height          =   450
      Left            =   4800
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton start 
      Caption         =   "start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox playersname 
      Height          =   285
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton joingame 
      Caption         =   "joingame"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox gamename 
      Height          =   285
      Left            =   4680
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label labeljoined 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   4200
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "List of Connected Players"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label label7 
      Alignment       =   2  'Center
      Caption         =   "Available Game"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Your Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Game Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label gameopen 
      Alignment       =   2  'Center
      Caption         =   "No Games Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lanchoice As Long 'address
Public details As String 'names
Public connected As Boolean 'if connected
Private Sub Form_Load()
Connect.Icon = LoadResPicture("ictac", vbResIcon) 'form icon
If usermode = "host" Then
join.Enabled = False
Else
host.Enabled = False
gamename.Visible = False
Label5.Visible = False
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'call on form cancel or exit by control box on form
If connectionmade = False Then
MainBoard.hostagame.Enabled = True
MainBoard.joinagame.Enabled = True
Call CloseDownDPlay
multiplayermode = False
End If
MainBoard.Enabled = True
End Sub
Private Sub host_Click()
On Error GoTo NO_Hosting ' error handler in case creating host fails
If playersname = "" Or gamename = "" Then
MsgBox "You must enter a Players name and Game Name", vbOKOnly, "Tic Tac Oops"
Exit Sub
End If
Call goplay 'starts direct play object
Dim address As DirectPlayAddress
'Selects which choice was made for lan
Set address = EnumConnect.GetAddress(lanchoice)
'Binds address to directplay connection
Call dxplay.InitializeConnection(address)
'Starts sessiondata information
Dim SessionData As DirectPlaySessionData
Set SessionData = dxplay.CreateSessionData
Call SessionData.SetMaxPlayers(2)
Call SessionData.SetSessionName(gamename.Text)
Call SessionData.SetFlags(DPSESSION_MIGRATEHOST)
Call SessionData.SetGuidApplication(AppGuid)
'Starts a new session initializes connection
Call dxplay.Open(SessionData, DPOPEN_CREATE)
'Create Player profile
Dim PlayerName As String
Dim playerhandle As String
PlayerName = playersname.Text
profilename = PlayerName
playerhandle = "Player(Host)"
MyPlayer = dxplay.CreatePlayer(PlayerName, playerhandle, 0, 0)
dxHost = True
gameopen.Caption = gamename.Text
Call updatedisplay 'Updates game list
Label8.Caption = "Waiting for other Players"
Exit Sub
NO_Hosting:
    MsgBox "Could not Host Game", vbOKOnly, "Try Again"
End Sub
Private Sub join_Click()
On Error GoTo Oops
Call goplay
Dim address As DirectPlayAddress
Set address = EnumConnect.GetAddress(lanchoice)
Call dxplay.InitializeConnection(address)
Dim details2 As Byte
Dim SessionData As DirectPlaySessionData
Set SessionData = dxplay.CreateSessionData
'Gets Session any open session info
Set EnumSession = dxplay.GetDPEnumSessions(SessionData, 0, DPENUMSESSIONS_AVAILABLE)
Set SessionData = EnumSession.GetItem(1)
'Get open session name
details = SessionData.GetSessionName
If details > "" And usermode = "client" Then
joingame.Enabled = True
End If
Call updatedisplay
gameopen.Caption = details
Exit Sub
Oops:
    MsgBox "Connection Failed", vbOKOnly, "Tic Tac Oops"
    Exit Sub
End Sub
Public Function goplay()
Set dxplay = dx7.DirectPlayCreate("") 'open directplay object
'gets connection types
Set EnumConnect = dxplay.GetDPEnumConnections("", DPCONNECTION_DIRECTPLAY)
End Function
Private Sub joingame_Click()
On Error GoTo Joinfailed
If playersname = "" Then
MsgBox "You must enter a Players name", vbOKOnly, "Tic Tac Oops"
Exit Sub
End If
Dim SessionData As DirectPlaySessionData
Set SessionData = EnumSession.GetItem(1)
'Joins open session
Call dxplay.Open(SessionData, DPOPEN_JOIN)
'creats and sends player info
PlayerName = playersname.Text
profilename = PlayerName
playerhandle = "Player(Client)"
MyPlayer = dxplay.CreatePlayer(PlayerName, playerhandle, 0, 0)
Call UpdateWaiting
joingame.Enabled = False
playersname.Enabled = False
MainBoard.mnuchat.Enabled = True
Exit Sub
Joinfailed:
    MsgBox "Joining Session Failed", vbOKOnly, "No Session Found"
    Exit Sub
End Sub
Public Sub UpdateWaiting()
  Dim StatusMsg As String
  Dim x As Integer
  Dim objDPEnumPlayers As DirectPlayEnumPlayers
  Dim SessionData As DirectPlaySessionData
  ' Enumerate players
  On Error GoTo ENUMERROR
  Set objDPEnumPlayers = dxplay.GetDPEnumPlayers("", 0)
  gNumPlayersWaiting = objDPEnumPlayers.GetCount
  ' Update label
  Set SessionData = dxplay.CreateSessionData
  Call dxplay.GetSessionDesc(SessionData)
  StatusMsg = gNumPlayersWaiting & " of " & SessionData.GetMaxPlayers _
          & " players ready..."
  Label8.Caption = StatusMsg
     If gNumPlayersWaiting = SessionData.GetMaxPlayers And usermode = "host" Then
        start.Enabled = True
        Label8.Caption = "Everyone is here Click Start"
        End If
       If gNumPlayersWaiting = SessionData.GetMaxPlayers And usermode = "client" Then
        start.Enabled = False
        Label8.Caption = "Waiting For Host To Start Session"
        End If
  ' Update listbox
  Dim PlayerName As String
  For x = 1 To gNumPlayersWaiting
    PlayerName = objDPEnumPlayers.GetShortName(x)
    If PlayerName <> playersname.Text Then
        labeljoined.Caption = PlayerName & " has joined the game."
        opponentsname = PlayerName
    End If
    Call lstPlayers.AddItem(PlayerName)
  Next x
  Exit Sub
ENUMERROR:
  MsgBox ("No Players Found")
  Exit Sub
  
End Sub
Private Sub lantype_Click(Index As Integer)
lanchoice = Index + 1
host.Visible = True
join.Visible = True
End Sub
Private Sub start_Click()
On Error GoTo CouldNotStart
Const msgsize = 21
Dim tnumplayers As DirectPlayEnumPlayers
Dim SessionData As DirectPlaySessionData
  ' Disable joining, in case we start before maximum no. of players reached. We
  ' don't want anyone slipping in at the last moment.
  Set SessionData = dxplay.CreateSessionData
  Call dxplay.GetSessionDesc(SessionData)    ' necessary?
  Call SessionData.SetFlags(SessionData.GetFlags + DPSESSION_JOINDISABLED)
  Call dxplay.SetSessionDesc(SessionData)
  ' Set global player count. This mustn't be done earlier, because someone might
  ' have dropped out or joined just as the host clicked Start.
Set tnumplayers = dxplay.GetDPEnumPlayers("", 0)
  numplayers = CByte(tnumplayers.GetCount)
Dim dpmsg As DirectPlayMessage
Dim pID As Long
Dim msgtype As Long
Dim x As Byte
        Set dpmsg = dxplay.CreateMessage
        dpmsg.WriteLong (MSG_STARTGAME) 'case selector
        dpmsg.WriteByte (numplayers) 'number of players
        Dim PlayerID As Long
For x = 0 To numplayers - 1
    PlayerID = tnumplayers.GetDPID(x + 1)
    dpmsg.WriteLong (PlayerID)
    ' Keep local copy of player IDs
    PlayerIDs(x) = PlayerID
    ' Assign place in order to the host
    If PlayerID = MyPlayer Then dxMyTurn = x
  Next x
Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpmsg)
    Hide
    MainBoard.Enabled = True
        MainBoard.Show
        MainBoard.playerdisplaylabel.Caption = opponentsname & " Has Joined The Game"
        MainBoard.StatusBar1.SimpleText = opponentsname & "Is Ready To Play,  Start Game"
        MainBoard.mnudisconnect.Enabled = True
        connectionmade = True
        multiplayermode = True
        MainBoard.mnuchat.Enabled = True
        onconnect = True
        Exit Sub
CouldNotStart:
    MsgBox "Could not start game.", vbOKOnly, "System"
End Sub
Private Function updatedisplay()
label7.Visible = True
gameopen.FontUnderline = False
gameopen.ForeColor = vbBlue
host.Enabled = False
join.Enabled = False
Dim Y As Byte
Y = 0
For Y = 0 To 2 Step 1
 lantype(Y).Enabled = False
      Next Y
End Function

