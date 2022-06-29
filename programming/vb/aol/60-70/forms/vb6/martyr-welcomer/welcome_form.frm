VERSION 5.00
Begin VB.Form FRMwelcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "welcome bot"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "s$ - person to welcome"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "n$ - room name"
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Welcome to <u>n$</u> <b>s$</b>"
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1800
      Top             =   1920
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "welcome bot options"
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "FRMwelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'last time i named a variable wrong and
'was confused as hell, so i added this to
'make sure i name my variables correctly.



'ever wanted to have a welcome bot without having to change
'the users options? then this is the example for you!
'enough of selling this example, read on and learn
'IM: stupid limits
'e-mail: stupid_limits@hotmail.com
'handle: martyr



'*update* 03-24-2002
'i was playing with this bot and figured
'more options were needed. so i added
'the TrimRoom function to make the welcome
'not so long.
'i was going to add a feature to disable
'the TrimRoom function, but I'd never
'disable it myself, so what's the point?

'-resolved issues-
'when going to a new room, the bot had no idea
'what the hell was going on, so it'd welcome
'everyone in the room. well i went through and
'added the various coding so it would know when
'it was in a new room.


Dim RoomName As String
'this variable is outside a function/sub so it
'will retain it's value through each function/sub.


Private Function TrimRoom() As String
'very helpful little function
    
    Dim RoomCaption As String
        RoomCaption = GetCaption(FindRoom)
            'get the caption of the current room

            If InStr(RoomCaption, "-") <> 0& Then
                'if we can find a hyphen, we'll
                'get the text to the right of it.
                RoomCaption = Mid(RoomCaption, InStr(RoomCaption, "-") + 2)
            End If
        
        TrimRoom = RoomCaption
        'setting the function the the final value of
        'RoomCaption
End Function

Private Sub Command1_Click()
    If Command1.Caption = "start" Then
        'ChatSend "<b>-[</b><u>Welcome Bot Activated</u><b>]-</b>"
        'i removed this because i was playing with the bot and found it to be very annoying.

        '*note* the bot is also very annoying.

        AddRoomToListbox List1, False
        RoomName = TrimRoom
        Me.Caption = "welcome to: " & RoomName
        Timer1.Enabled = True
        Command1.Caption = "stop"
    ElseIf Command1.Caption = "stop" Then
        'ChatSend "<b>-[</b><u>Welcome Bot Deactivated</u><b>]-</b>"
        List1.Clear
        List2.Clear
        Me.Caption = "welcome bot"
        Timer1.Enabled = False
        Command1.Caption = "start"
    End If
End Sub

Private Sub Timer1_Timer()
'this is where the magic happens
    Dim Checkin As Boolean
    Dim L1 As Long
    Dim L2 As Long
    Dim Tmpstr As String
    Dim Welcome As String

    If Not TrimRoom Like RoomName Then
        'if the trimmed down name doesn't match the value of the variable RoomName...
        
        RoomName = TrimRoom
        'change the variable's value to the new value
        AddRoomToListbox List1, False
        'get a new list of people to list1
        
        Me.Caption = "welcome to: " & RoomName
    End If

        AddRoomToListbox List2, False
        'get a new copy of the chat rooms list
        Checkin = False
        'set checkin to false

        For L2 = 0 To List2.ListCount - 1
            Tmpstr = List2.List(L2)
            'get the first name off our new list

            Welcome = ReplaceString(Text1.Text, "n$", TrimRoom)
            Welcome = ReplaceString(Welcome, "s$", Tmpstr)

                For L1 = 0 To List1.ListCount - 1
                    'look for the name in our old list
                    If List1.List(L1) Like Tmpstr Then Checkin = True: Exit For
                    'if the name is found then set checkin to true and get out of this for statement
                Next L1

            If Checkin = False Then ChatSend Welcome
            'if we couldn't find the name, then welcome them to the room
            Checkin = False
            'reset the variable to false incase the person was in the old list
        Next L2
        
        List1.Clear
        List2.Clear
        'clear the lists so they don't build up to be godlike
        
        AddRoomToListbox List1, False
        'update the list for comparison
End Sub
