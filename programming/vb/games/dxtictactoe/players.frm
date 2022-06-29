VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MainBoard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8175
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000040&
   Icon            =   "players.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "players.frx":000C
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox chatbox 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   17
      Text            =   "Type Message Here"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton send_chat 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tic Tac Chat"
      Height          =   2775
      Left            =   360
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   3975
      Begin RichTextLib.RichTextBox chatlabel 
         Height          =   1815
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3201
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"players.frx":6AA8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5250
      Top             =   510
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   5970
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton restart 
      Caption         =   "Restart Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3150
      TabIndex        =   2
      Top             =   2235
      Width           =   1245
   End
   Begin VB.Label Game_Over 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6060
      TabIndex        =   13
      Top             =   5100
      Width           =   135
   End
   Begin VB.Label Layer_A 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   8
      Left            =   6675
      TabIndex        =   12
      Top             =   4095
      Width           =   555
   End
   Begin VB.Label Layer_A 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   7
      Left            =   5835
      TabIndex        =   11
      Top             =   4095
      Width           =   555
   End
   Begin VB.Label Layer_A 
      BackColor       =   &H00000000&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   4950
      TabIndex        =   10
      Top             =   4095
      Width           =   555
   End
   Begin VB.Label Layer_A 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   6675
      TabIndex        =   9
      Top             =   3180
      Width           =   555
   End
   Begin VB.Label Layer_A 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   5835
      TabIndex        =   8
      Top             =   3180
      Width           =   555
   End
   Begin VB.Label Layer_A 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   4950
      TabIndex        =   7
      Top             =   3180
      Width           =   555
   End
   Begin VB.Label Layer_A 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   6675
      TabIndex        =   6
      Top             =   2265
      Width           =   555
   End
   Begin VB.Label Layer_A 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   1
      Left            =   5835
      TabIndex        =   5
      Top             =   2265
      Width           =   555
   End
   Begin VB.Label Layer_A 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Index           =   0
      Left            =   4950
      MouseIcon       =   "players.frx":6B8E
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Top             =   2250
      Width           =   555
   End
   Begin VB.Label playerdisplaylabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   345
      TabIndex        =   1
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Out_Box 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   6030
      TabIndex        =   0
      Top             =   1320
      Width           =   150
   End
   Begin VB.Menu Multiplay 
      Caption         =   "&MultiPlayer"
      Begin VB.Menu hostagame 
         Caption         =   "&Host a Game"
      End
      Begin VB.Menu joinagame 
         Caption         =   "&Join a Game"
      End
      Begin VB.Menu mnudisconnect 
         Caption         =   "&Disconnect"
      End
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Begin VB.Menu mnuchat 
         Caption         =   "&Chat On"
         Enabled         =   0   'False
      End
      Begin VB.Menu xxxxxxx 
         Caption         =   "Who's First"
         Begin VB.Menu x 
            Caption         =   "X's"
         End
         Begin VB.Menu o 
            Caption         =   "O's"
         End
      End
   End
   Begin VB.Menu game 
      Caption         =   "&Game"
      Begin VB.Menu newgame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "MainBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'*   Multiplayer Tic Tac Toe                                 *
'*      Mike Altmanshofer-Masher2@ibm.net                    *
'*         Distribute Freely                                 *
'*           Please Leave this Header description intact.    *
'*          Created 06-01-99 Last Modified 10-10-99(DX7)     *
'*                                                           *
'*                                                           *
'*                                                           *
'*************************************************************
Option Explicit
Dim a(9) As Integer
Dim Player_A(9) As Integer 'Initialize X array
Dim Computer_A(9) As Integer 'Initialize O array
Dim Test_Result(8) As Integer
Dim Win(3) As Integer ' Spots won to marked
Dim m, Token, first_turn, temp1 As Integer
Dim Temp As Boolean  'check whether player won
Dim Sq_Left, n1, mark As Integer
Dim tr As String 'string passed on win to mark routine
Dim Begin As Boolean 'continue winning spots flashing
Dim sw As Boolean  'Sets whether X or O starts game
Public Sub Initialize()
' select who's turn
If usermode = "host" And multiplayermode = True Then
' set o or x first
    If sw = True Then
        MyTurn = True
    Else
        MyTurn = False
    End If
End If
If multiplayermode = False Then
   MyTurn = True
End If
Begin = False   ' cancel marking routine
score = score + 1 'adds one to gamecount
If multiplayermode = True Then
If usermode = "client" And sw = True Then
       MyTurn = False
ElseIf usermode = "client" And sw = False Then
       MyTurn = True
End If
End If

'Start SW true mode**********************************
      'initialize game settings
If sw = True Then
    StatusBar1.SimpleText = "New Game Initialized" & "           X's Turn"
    Debug.Print "Turn Status " & MyTurn
    Debug.Print "SW Value is " & sw
    Dim u As Integer
    u = 0
    Sq_Left = 9
    Token = 10
        For u = 0 To 8
            Layer_A(u).MousePointer = vbCustom
            'select starting icon and characteristics****************************
                If usermode = "host" Then
                    Layer_A(u).MouseIcon = LoadResPicture("x", vbResIcon)
                Else
                    Layer_A(u).MouseIcon = LoadResPicture("nyt", vbResIcon)
                End If
            Layer_A(u).FontSize = 28
            Layer_A(u).FontBold = True
            Layer_A(u).Caption = ""
            Layer_A(u).BackStyle = 0
            Layer_A(u).Alignment = 2
            Player_A(u) = 0
            Computer_A(u) = 0
            Layer_A(u).Enabled = True
        Next u
    'update statusbar and display routine******************************
    If usermode = "host" And multiplayermode = True Then
        StatusBar1.SimpleText = "New Game Initialized          " & profilename & "'s Turn"
        Out_Box.Caption = profilename & "'s Turn."
    End If
    If usermode = "client" And multiplayermode = True Then
        StatusBar1.SimpleText = "New Game Initialized          " & opponentsname & "'s Turn"
        Out_Box.Caption = opponentsname & "'s Turn."
    End If
    If multiplayermode = False Then
        Out_Box.Caption = "X Goes First"
    End If
End If
'End sw true*********************************************
'set starting icon*****************
If sw = False Then
    StatusBar1.SimpleText = "New Game Initialized" & "           O's Turn"
    Debug.Print "Turn Status " & MyTurn
    Debug.Print "SW Value is " & sw
    u = 0
    Sq_Left = 9
    Token = 10
        For u = 0 To 8
            Layer_A(u).MousePointer = vbCustom
                If usermode = "host" And multiplayermode = True Then
                    Layer_A(u).MouseIcon = LoadResPicture("nyt", vbResIcon)
                Else
                    Layer_A(u).MouseIcon = LoadResPicture("o", vbResIcon)
                End If
            Layer_A(u).FontSize = 28
            Layer_A(u).FontBold = True
            Layer_A(u).Caption = ""
            Layer_A(u).BackStyle = 0
            Layer_A(u).Alignment = 2
            Player_A(u) = 0
            Computer_A(u) = 0
            Layer_A(u).Enabled = True
        Next u
    Temp = False  'initiate no win
    'Update Statusbar and outbox display********************8
        If usermode = "client" And multiplayermode = True Then
            StatusBar1.SimpleText = "New Game Initialized          " & profilename & "'s Turn"
            Out_Box.Caption = profilename & " 's Turn."
        End If
        If usermode = "host" And multiplayermode = True Then
            StatusBar1.SimpleText = "New Game Initialized          " & opponentsname & "'s Turn"
            Out_Box.Caption = opponentsname & " 's Turn."
        End If
        If multiplayermode = False Then
            Out_Box.Caption = "O Goes First"
        End If
End If
'End sw false*********************************************
Debug.Print "Ran Initialization Myturn status is " & MyTurn
Game_Over.Caption = "New Game"
End Sub
Private Sub exit_Click()
If onconnect = True Then 'checks for connection
On Error GoTo NoDx 'error to handle dxplay not initialized
Dim dpmsg As DirectPlayMessage
Set dpmsg = dxplay.CreateMessage
Call dpmsg.WriteLong(MSG_STOP) 'Sends player quit message to other player
Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpmsg)
Call CloseDownDPlay 'shuts down dxplay
End If
Unload Connect 'unloads connect form if connect frees memory
Unload MainBoard 'unloads board before ending to free memory
End
NoDx:
    MsgBox "Could not stop DXPlay.", vbOKOnly, "System"
    End
End Sub
Private Sub Form_Load()
On Error GoTo NoLoad 'Handles errors in case form won't load
MainBoard.Icon = LoadResPicture("ictac", vbResIcon) 'form icon
restart.Visible = False 'restart button not seen on single player or client mode
mnudisconnect.Enabled = False 'set menu item to no connect state
onconnect = False 'Sets connection status to false by default
sw = True 'set starting Player to x
x.Checked = True 'set menuitem X to x checked
multiplayermode = False 'initiate mode to false
Call deinitialize  'disables all squares until gamemode and multiplayer mode is decided
score = 0  'sets game count to 0
Exit Sub
NoLoad:
MsgBox "Could Not Load Form", vbOKOnly, "Quitting"
End

End Sub
Private Sub deinitialize()
'Disables all squares until game selection is made
Dim m As Integer
    For m = 0 To 8
        Layer_A(m).MousePointer = vbCustom
            If sw = True Then 'sets mouse pointer to x for x first
                Layer_A(m).MouseIcon = LoadResPicture("x", vbResIcon)
            Else 'sets mouse pointer to O for O first
                Layer_A(m).MouseIcon = LoadResPicture("o", vbResIcon)
            End If
        Layer_A(m).FontSize = 28
        Layer_A(m).FontBold = True
        Layer_A(m).Caption = ""
        Layer_A(m).BackStyle = 0
        Layer_A(m).Alignment = 2
        Layer_A(m).Enabled = False
    Next m
'Update Status Bar
StatusBar1.SimpleText = "Select Game- New Game or Multiplayer option to start game"
Out_Box.Caption = "Start New Game."
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If onconnect = True Then
    On Error GoTo NoDx
    Dim dpmsg As DirectPlayMessage
    Set dpmsg = dxplay.CreateMessage
    Call dpmsg.WriteLong(MSG_STOP)
    Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpmsg)
    Call CloseDownDPlay
End If
Unload Connect
Unload MainBoard
End
NoDx:
    MsgBox "Could not stop DXPlay.", vbOKOnly, "System"
    End
End Sub

Private Sub hostagame_Click()
usermode = "host" 'Sets usermode to host
Connect.Show 'starts connect form
MainBoard.Enabled = False 'disable form so user cannot select while connect form is up
hostagame.Enabled = False 'disables menu host button.
joinagame.Enabled = False ' disables menu join button
multiplayermode = True 'sets multiplayer to true
End Sub
Private Sub joinagame_Click()
usermode = "client" 'Sets usermode to client
Connect.Show
MainBoard.Enabled = False
multiplayermode = True
End Sub
Private Sub Layer_A_Click(Index As Integer)
playerdisplaylabel.Caption = ""
'Used For single player board selection or multiplayer your turn selection
Debug.Print "Layer A Click Turn Status " & MyTurn
Debug.Print "Layer A Multiplayer Mode Status " & multiplayermode
If multiplayermode = True And MyTurn = False Then  'Easy way to exit if not your turn
    Exit Sub
End If
If Sq_Left Mod 2 = 1 Then 'check remainder of squares left divided by 2
        If sw = True Then ' sets who goes first X or O
            Layer_A(Index).Caption = "X"
        Else
            Layer_A(Index).Caption = "O"
        End If
    Layer_A(Index).Enabled = False 'Sets selected square to not available
    Player_A(Index) = 1
    Computer_A(Index) = -Token
    LoadPlayer
        If multiplayermode = True And MyTurn = True Then 'checks for multiplayer and turn status
            'This routine below packs message to send
            'to other player to select the square chosen.
            Dim dpmsg As DirectPlayMessage 'alot direct playmessage
            Set dpmsg = dxplay.CreateMessage 'set and create the message
            Call dpmsg.WriteLong(MSG_MOVE) 'pack message structure and identify type
            Call dpmsg.WriteByte(Index) 'Packs case selection number to msgtype.
            'This sends the pack message structure
            Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpmsg)
        End If
        If multiplayermode = True Then 'Sets routines to not your turn on multiplayer
            Dim Y As Integer
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("nyt", vbResIcon)
                Next Y
            'Update Status displays
            StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & opponentsname & "'s Turn"
            Out_Box.Caption = opponentsname & "'s Turn."
        End If
    'Everything below until mod else statement is single player
        If multiplayermode = False Then 'Sets X or O turn status on single player
            If sw = True Then
                StatusBar1.SimpleText = "New Game Initialized          O's Turn"
            Else
                StatusBar1.SimpleText = "New Game Initialized         X's Turn"
            End If
        If sw = True Then
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("o", vbResIcon)
                Next Y
        Else
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("x", vbResIcon)
                Next Y
        End If
        If sw = True Then
            Out_Box.Caption = "O's Turn"
        Else
            Out_Box.Caption = "X's Turn"
        End If
    End If
Else
    'Mod else*********************************
        If sw = True Then
            Layer_A(Index).Caption = "O"
        Else
            Layer_A(Index).Caption = "X"
        End If
    Layer_A(Index).Enabled = False
    Player_A(Index) = -Token
    Computer_A(Index) = 1
        If multiplayermode = True Then
            StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & opponentsname & "'s Turn"
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("nyt", vbResIcon)
                Next Y
            Out_Box.Caption = opponentsname & "'s Turn."
        End If
        If multiplayermode = False Then
            If sw = True Then
                StatusBar1.SimpleText = "New Game Initialized          X's Turn"
            Else
                StatusBar1.SimpleText = "New Game Initialized          O's Turn"
            End If
        If sw = True Then
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("x", vbResIcon)
                Next Y
            Out_Box.Caption = "X's Turn"
        Else
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("o", vbResIcon)
                Next Y
            Out_Box.Caption = "O's Turn"
        End If
    End If
    LoadComputer
        If multiplayermode = True And MyTurn = True Then
          'Same as above packs message and sends move to other player
            Dim dpmsg2 As DirectPlayMessage
            Set dpmsg2 = dxplay.CreateMessage
            Call dpmsg2.WriteLong(MSG_MOVE)
            Call dpmsg2.WriteByte(Index)
            Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpmsg2)
        End If
End If
Sq_Left = Sq_Left - 1
EvalNextMove
MyTurn = False
End Sub
Public Function layer_A_online(Index As Integer)
playerdisplaylabel.Caption = ""
'This routine is called to mark sqares when remote computer
'sends a move made command.
'Same as above with some redundant routines removed
If Sq_Left Mod 2 = 1 Then
    If sw = True Then
        Layer_A(Index).Caption = "X"
    Else
        Layer_A(Index).Caption = "O"
    End If
    Layer_A(Index).Enabled = False
    Player_A(Index) = 1
    Computer_A(Index) = -Token
        If multiplayermode = True Then
            If sw = True Then
                StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & profilename & "'s Turn"
                Out_Box.Caption = profilename & "'s Turn."
                Dim Y As Integer
                    For Y = 0 To 8
                        Layer_A(Y).MouseIcon = LoadResPicture("o", vbResIcon)
                    Next Y
            Else
                StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & profilename & "'s Turn"
                Out_Box.Caption = profilename & "'s Turn."
                Y = 0
                    For Y = 0 To 8
                        Layer_A(Y).MouseIcon = LoadResPicture("x", vbResIcon)
                    Next Y
            End If
        End If
        If multiplayermode = False Then
            If sw = True Then
                Y = 0
                    For Y = 0 To 8
                        Layer_A(Y).MouseIcon = LoadResPicture("o", vbResIcon)
                        Out_Box.Caption = "O's Turn"
                    Next Y
            Else
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("x", vbResIcon)
                    Out_Box.Caption = "X's Turn"
            Next Y
            End If
        End If
    LoadPlayer
Else
    If sw = True Then
        Layer_A(Index).Caption = "O"
    Else
        Layer_A(Index).Caption = "X"
    End If
    Layer_A(Index).Enabled = False
    Player_A(Index) = -Token
    Computer_A(Index) = 1
    If multiplayermode = True Then
        If sw = True Then
            StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & profilename & "'s Turn"
            Out_Box.Caption = profilename & "'s Turn."
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("x", vbResIcon)
                Next Y
        Else
            StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & profilename & "'s Turn"
            Out_Box.Caption = profilename & "'s Turn."
            Y = 0
                For Y = 0 To 8
                    Layer_A(Y).MouseIcon = LoadResPicture("o", vbResIcon)
            Next Y
        End If
    End If
    If multiplayermode = False Then
        If sw = True Then
            StatusBar1.SimpleText = "New Game Initialized          X's Turn"
        Else
            StatusBar1.SimpleText = "New Game Initialized          O's Turn"
    End If
    If sw = True Then
        Y = 0
            For Y = 0 To 8
                Layer_A(Y).MouseIcon = LoadResPicture("x", vbResIcon)
            Next Y
            Out_Box.Caption = "X's Turn"
    Else
        Y = 0
            For Y = 0 To 8
                Layer_A(Y).MouseIcon = LoadResPicture("o", vbResIcon)
            Next Y
        Out_Box.Caption = "O's Turn"
    End If
    End If
        LoadComputer
End If
Sq_Left = Sq_Left - 1
EvalNextMove
End Function
Private Sub scan_3() '*****************************************
Dim r As Integer
For r = 0 To 7
    If Test_Result(r) = 3 Then
    Temp = True
    End If
Next r
End Sub
Private Sub EvalNextMove() '***********************************
test
scan_3
Debug.Print "Squares Left Value on Evaluate Next Move " & Sq_Left
Debug.Print "Boolean Temp Value on Evaluate " & Temp
Debug.Print "Token Value on Eval. " & Token
If Temp = True Then
    If Sq_Left Mod 2 = 0 Then 'Makes win or lose calls Turn checking is made later
        Player_Wins 'call player wins routine
    Else
        Computer_Wins 'calls computer rountine
    End If
End If
Temp = False
If Sq_Left <= 0 Then
    Cats_Game
    Begin = False 'Turns off mark routine
        If multiplayermode = True And usermode = "host" Then 'sets turn to true
            MyTurn = True
            Debug.Print "Set myturn to true on win"
        End If
End If
first_turn = 1
End Sub
Private Sub Computer_Wins()
Dim s As Integer
For s = 0 To 8
    Layer_A(s).Enabled = False
Next s
Begin = True
If multiplayermode = True And usermode = "host" Then
    If sw = True Then 'Checks for Whos Turn and update Host or client
        Out_Box.Caption = opponentsname & " Won!"
        opponentsscore = opponentsscore + 1
    Else
        Out_Box.Caption = profilename & " Won!"
        profilenamescore = profilenamescore + 1
    End If
End If
If multiplayermode = True And usermode = "client" Then
    If sw = True Then
        Out_Box.Caption = profilename & " Won!"
        profilenamescore = profilenamescore + 1
    Else
        Out_Box.Caption = opponentsname & " Won!"
        opponentsscore = opponentsscore + 1
    End If
End If
If multiplayermode = False Then 'Single Player updating
    If sw = True Then
        Out_Box.Caption = "O Won!!!!"
    Else
        Out_Box.Caption = "X Won!!!!!"
    End If
End If
Game_Over.Caption = "Game Over"
'Shows Resart Option if Host
If multiplayermode = True And usermode = "host" Then
    restart.Visible = True
    restart.Enabled = True
End If
Timer4.Enabled = True 'Sets timer to time mark routine
If sw = True Then 'Checks Whos turn sends string to mark
     Call Mark_Win("O")
Else
     Call Mark_Win("X")
End If
End Sub
Private Sub Player_Wins()
'See computer wins for details
Dim a As Integer
    For a = 0 To 8
        Layer_A(a).Enabled = False
    Next a
Begin = True
If multiplayermode = True And usermode = "host" Then
    If sw = True Then
        profilenamescore = profilenamescore + 1
        Out_Box.Caption = profilename & " Won!"
    Else
        opponentsscore = opponentsscore + 1
        Out_Box.Caption = opponentsname & " Won!"
    End If
End If
If multiplayermode = True And usermode = "client" Then
    If sw = True Then
        opponentsscore = opponentsscore + 1
        Out_Box.Caption = opponentsname & " Won!"
    Else
        profilenamescore = profilenamescore + 1
        Out_Box.Caption = profilename & " Won!"
    End If
End If
If multiplayermode = False Then
    If sw = True Then
        Out_Box.Caption = "X Won!!!!"
    Else
        Out_Box.Caption = "O Won!!!!!"
    End If
End If
Game_Over.Caption = "Game Over"
If multiplayermode = True And usermode = "host" Then
    restart.Visible = True
    restart.Enabled = True
End If
Timer4.Enabled = True
If sw = True Then
    Call Mark_Win("X")
Else
    Call Mark_Win("O")
End If
End Sub
Private Sub Mark_Win(tr As String) 'Marks winning squares
Dim PauseTime, start, Finish, TotalTime
While Begin = True
    PauseTime = 0.3  ' Set duration.
    start = Timer   ' Set start time.
        Do While Timer < start + PauseTime And Begin = True
            For n1 = 0 To 2
                mark = Win(n1)
                Layer_A(mark).Caption = tr
                Layer_A(mark).FontBold = False
            Next n1
            DoEvents    ' Yield to other processes.
        Loop
    start = Timer   ' Set start time.
        Do While Timer < start + PauseTime And Begin = True
            For n1 = 0 To 2
                mark = Win(n1)
                Layer_A(mark).FontBold = True
                Layer_A(mark).Caption = tr
            Next n1
            DoEvents    ' Yield to other processes.
        Loop
Wend
End Sub
Private Sub test() 'Tests conditions for the win
Dim n, k, sample As Integer
    sample = 0
    For n = 0 To 2
        Test_Result(sample) = a(3 * n) + a(3 * n + 1) + a(3 * n + 2)
        If Test_Result(sample) = 3 Then
            Win(0) = 3 * n
            Win(1) = 3 * n + 1
            Win(2) = 3 * n + 2
        End If
        sample = sample + 1
    Next n
    For n = 0 To 2
        Test_Result(sample) = a(n) + a(n + 3) + a(n + 6)
        If Test_Result(sample) = 3 Then
            Win(0) = n
            Win(1) = n + 3
            Win(2) = n + 6
        End If
        sample = sample + 1
    Next n
    Test_Result(sample) = a(0) + a(4) + a(8)
        If Test_Result(sample) = 3 Then
            Win(0) = 0
            Win(1) = 4
            Win(2) = 8
        End If
    sample = sample + 1
    Test_Result(sample) = a(6) + a(4) + a(2)
        If Test_Result(sample) = 3 Then
            Win(0) = 6
            Win(1) = 4
            Win(2) = 2
        End If
    sample = sample + 1
    
End Sub
Private Sub LoadPlayer()
Dim e As Integer
For e = 0 To 8
    a(e) = Player_A(e)
Next e
End Sub
Private Sub LoadComputer()
Dim w As Integer
For w = 0 To 8
    a(w) = Computer_A(w)
Next w
End Sub
Private Sub Cats_Game() 'Cats Game display routine
GameUnderway = False
Dim z As Integer
For z = 0 To 8
    Layer_A(z).Enabled = False
Next z
Out_Box.Caption = "Cat's Game!"
Game_Over.Caption = "Game Over"
If multiplayermode = True And usermode = "host" Then
    restart.Visible = True
    restart.Enabled = True
End If
End Sub
Private Sub mnuchat_Click() 'Menu button for chatbox routine
On Error GoTo NoChat 'error handler in case chat initialization problem.
If mnuchat.Checked = True Then
    Frame1.Visible = False
    chatlabel.Visible = False
    send_chat.Visible = False
    chatbox.Visible = False
    mnuchat.Checked = False
    'Packs and sends DXplay message to switch chat on off
    Dim chaton As DirectPlayMessage
    Set chaton = dxplay.CreateMessage
    Call chaton.WriteLong(MSG_CHAT_ON)
    Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, chaton)
    
Else
    Frame1.Visible = True
    chatlabel.Visible = True
    send_chat.Visible = True
    chatbox.Visible = True
    mnuchat.Checked = True
    chatbox.Visible = True
    chatbox.SetFocus
     'Packs and sends DXplay message to switch chat on off
    Dim chaton2 As DirectPlayMessage
    Set chaton2 = dxplay.CreateMessage
    Call chaton2.WriteLong(MSG_CHAT_ON)
    Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, chaton2)
End If
Exit Sub
NoChat:
    MsgBox "Could Not Start Chat", vbOKOnly, "Oops"
    Exit Sub
End Sub
Public Function chatswitch() 'Menu button for incoming online Chatbox routine
On Error GoTo NoChat
If mnuchat.Checked = True Then
     Frame1.Visible = False
    chatlabel.Visible = False
    send_chat.Visible = False
    chatbox.Visible = False
    mnuchat.Checked = False
Else
   Frame1.Visible = True
    chatlabel.Visible = True
    send_chat.Visible = True
    chatbox.Visible = True
    mnuchat.Checked = True
    chatbox.Visible = True
    chatbox.SetFocus
End If
Exit Function
NoChat:
    MsgBox "Could Not Start Chat", vbOKOnly, "Oops"
    Exit Function
End Function
Private Sub mnudisconnect_Click() 'Disconnects and sends disconnect message
mnudisconnect.Enabled = False
newgame.Enabled = True
hostagame.Enabled = True
joinagame.Enabled = True
multiplayermode = False
usermode = "host"
'Sends player has left message to other players
Dim dpmsg As DirectPlayMessage
Set dpmsg = dxplay.CreateMessage
Call dpmsg.WriteLong(MSG_STOP)
Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpmsg)
Call CloseDownDPlay
Unload Connect
onconnect = False
End Sub
Private Sub newgame_Click() 'starts new game single or multiplayer
On Error GoTo NoGame
If usermode = "client" And multiplayermode = True Then
    MsgBox "Only the host can restart the game.", vbOKOnly, "Tic Tac Oops"
    Exit Sub
End If

If multiplayermode = False Then
    usermode = "host"
    Call Initialize
Else
    Call restart_Click 'call restart routine for multiplayer
End If
Exit Sub
NoGame:
    MsgBox "Could Not Start Game.", vbOKOnly, "Oops"
    Exit Sub
End Sub

Public Sub o_Click() 'sets menu item whos first o
If GameUnderway = True Then
MsgBox "You cannot chang this option while a game is in play", vbOKOnly, "Tic Tac Oops"
Exit Sub
End If

If o.Checked = True Then
    sw = False
    Exit Sub
Else
    o.Checked = True
    x.Checked = False
    sw = False
End If
If multiplayermode = True Then
'Sends who goes first message.
   Dim dpmsg As DirectPlayMessage
Set dpmsg = dxplay.CreateMessage
        Call dpmsg.WriteLong(MSG_XORO)
        Call dpmsg.WriteByte(2)
        Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, _
                dpmsg)
End If
Debug.Print "menu X or O clicked sw is " & sw
End Sub

Public Sub restart_Click() 'Restarts Game and updates scores
GameUnderway = True
multiplayermode = True
If usermode = "host" Then
Dim dpmsg As DirectPlayMessage
Set dpmsg = dxplay.CreateMessage
        Call dpmsg.WriteLong(MSG_RESTART)
        Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, _
                dpmsg)
End If

Call Initialize
If usermode = "host" Then
    If sw = True Then
        MyTurn = True
        StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & profilename & "'s Turn"
        playerdisplaylabel.Caption = profilename & "'s Turn."
    Else
        
        MyTurn = False
        StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & opponentsname & "'s Turn"
        playerdisplaylabel.Caption = opponentsname & "'s Turn."
    End If
End If
If usermode = "client" Then
    If sw = True Then
        MyTurn = False
        StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & opponentsname & "'s Turn"
        playerdisplaylabel.Caption = opponentsname & "'s Turn."
    Else
        MyTurn = True
        StatusBar1.SimpleText = "Game count is " & score & "    " & opponentsname & ":" & opponentsscore & " | " & profilename & ":" & profilenamescore & "           " & profilename & "'s Turn"
        playerdisplaylabel.Caption = profilename & "'s Turn."
    End If
End If
restart.Visible = False

End Sub
Private Sub send_chat_Click()
'handles chat boxes
Const chatlen = 5 + MChatString
Dim msgdata(chatlen) As Byte
Dim x As Integer
'packs and sends chat box information
Dim cmsg As DirectPlayMessage
Set cmsg = dxplay.CreateMessage
Call cmsg.WriteLong(MSG_CHAT)
Call cmsg.WriteString(chatbox.Text)
Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, cmsg)
If chatlabel.Text = "" Then
    chatlabel.Text = profilename & ": " & chatbox.Text
      Else
    chatlabel.Text = chatlabel.Text & vbCrLf & profilename & ": " & chatbox.Text
End If
chatbox.Text = ""
End Sub
Private Sub Timer4_Timer()
GameUnderway = False
'sets begin to false to stop letters from flashing.
'Updates score and status bar.
Begin = False
If usermode = "host" And multiplayermode = True Then
    StatusBar1.SimpleText = "Select Restart Game.  " & "Game #" & score & "      " & profilename & ":" & profilenamescore & " " & opponentsname & ":" & opponentsscore
    MyTurn = True
ElseIf usermode = "client" And multiplayermode = True Then
    StatusBar1.SimpleText = "Waiting on Host To Restart.   " & "Game #" & score & "      " & profilename & ":" & profilenamescore & " " & opponentsname & ":" & opponentsscore
End If
Timer4.Enabled = False
End Sub
Public Sub x_Click() 'handles menu item X whos turn first
If GameUnderway = True Then
MsgBox "You cannot chang this option while a game is in play", vbOKOnly, "Tic Tac Oops"
Exit Sub
End If
If x.Checked = True Then
    sw = True
    Exit Sub
Else
    x.Checked = True
    o.Checked = False
    sw = True
End If
If multiplayermode = True Then
'Sends who goes first message.
  Dim dpmsg As DirectPlayMessage
Set dpmsg = dxplay.CreateMessage
        Call dpmsg.WriteLong(MSG_XORO)
        Call dpmsg.WriteByte(1)
        Call dxplay.Send(MyPlayer, DPID_ALLPLAYERS, DPSEND_GUARANTEED, _
                dpmsg)
End If
Debug.Print "menu X or O clicked sw is " & sw
End Sub

