VERSION 5.00
Object = "{469F500F-1932-11D5-A4E9-444553540000}#1.0#0"; "AOL6.ocx"
Begin VB.Form frmaol6example 
   Caption         =   "AOL 6.0 Example Form"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "run menu goto web"
      Height          =   495
      Index           =   4
      Left            =   2400
      TabIndex        =   15
      ToolTipText     =   "this will show how to run a menu"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "run menu edit hotkeys"
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   14
      ToolTipText     =   "this will show how to run a menu"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "run menu my profile"
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "this will show how to run a menu"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "run menu im"
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   12
      ToolTipText     =   "this will show how to run a menu"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "run menu new mail"
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "this will show how to run a menu"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "get user name"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "get AOL screen name"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "uningnore im from..."
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ignore IM from..."
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "if off"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "your ims..."
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "im on"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "your ims.."
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Send Im"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      ToolTipText     =   "send an im to someone"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send Chat Text"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "send text to the chat room"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KeyWord"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "open a keyword"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find Chat Room"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "find the handle of the chat room"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Chat Room"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "shows how to open a chat room"
      Top             =   0
      Width           =   1215
   End
   Begin AOL6_OCX.AOL AOL1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "results"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   3615
   End
End
Attribute VB_Name = "frmaol6example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  'this will make sure every thing is declared properly..

Private Sub Command1_Click()
    AOL1.ChatRoom "AOL 6 Chat Room Test", PrivateRoom
    'this willl open a chat room, you can also open a
    'member room...
End Sub

Private Sub Command10_Click()
    'get the users screen name...
    Label1.Caption = AOL1.GetUser
End Sub

Private Sub Command11_Click(Index As Integer)
'theser are just some of the things you can do
'i am just showing you how to do each thing
    Select Case Index
        Case 0
            AOL1.RunMenu Mail, New_Mail
        Case 1
            AOL1.RunMenu People, 0, IM
        Case 2
            AOL1.RunMenu Settings, 0, 0, 0, My_Profile
        Case 3
            AOL1.RunMenu Favorites, 0, 0, 0, 0, Edit_HotKeys
        Case 4
            AOL1.RunMenu AOL_Services, 0, 0, GoTo_Web
    End Select
End Sub

Private Sub Command2_Click()
    'this will find the handle of a chat room
    'it's useful for ignorer's or things of that sort..
    Label1.Caption = AOL1.FindChatroom
End Sub

Private Sub Command3_Click()
    'duhh  opens a keyword of your choice...
    AOL1.Keyword "VB"
End Sub

Private Sub Command4_Click()
    'uses FindChatRoom to find the chat room..
    'then searches for the edit window and sends some
    'text...
    AOL1.ChatSendText "testing..."
End Sub

Private Sub Command5_Click()
    Dim Sn              As String   'the screen name to send to..
    Dim Msg             As String   'the message to send..
    Dim Recieved        As Boolean  'this will tell you whether or not they got the IM
    'declare the variables
    'you can see that i have tabbed the "As String"
    'this makes it easier on your eyes, never do this...
    'Dim X, Y, Z As Long
    'above you have made X and Y variants
    'and Z long, also don't do this...
    'Dim X As Long, Y As Long, Z As Long
    'that gets confusing, especially if someone
    'else is going to be reading your code...
    Sn = InputBox("Please enter the Screen Name of the Reciever..", "Send IM", AOL1.GetUser())
    'figure out the screen name to send to
    'default is the user name...
    If Sn = "" Then Exit Sub
    'uh oh they didn't type anything... exit the sub
    Msg = InputBox("Please enter the message to send to " & Sn & ".", "Send IM", "Testing...")
    'figure out the message to send
    'default is testing...
    If Msg = "" Then Exit Sub
    'didn't type any message, that won't work...
    Call AOL1.IMSend(Sn, Msg, Recieved)
    'call the send IM function, recieved is optional...
    'if you want to find out if they got the im or not
    'the do it like this...
    Select Case Recieved
        Case True
            Label1.Caption = Sn & " Has recived your message..."
            'they got it...
        Case False
            Label1.Caption = Sn & " Is not online or has there messages off..."
            'nope no luck
    End Select
End Sub

Private Sub Command6_Click()
    AOL1.IMSend "", "", False, True
    'you can see what we do here.
    'not to hard
End Sub

Private Sub Command7_Click()
    AOL1.IMSend "", "", False, False, True
    'this also isn't hard but if you can't figure it out
    'then ummm oh well!!
End Sub

Private Sub Command8_Click()
    Dim Sn          As String
    
    Sn = InputBox("Please enter the screen name to block..", "IM Block..")
    If Sn = "" Then Exit Sub
    AOL1.IMSend "$IM_OFF " & Sn, "Blocking Your IM's.."
    'ignore the im's of this person...
    
End Sub

Private Sub Command9_Click()
    Dim Sn          As String
    
    Sn = InputBox("Please enter the screen name to block..", "IM Block..")
    If Sn = "" Then Exit Sub
    AOL1.IMSend "$IM_ON " & Sn, "No Longer Blocking your im's.."
    'unignore the ims...
End Sub
