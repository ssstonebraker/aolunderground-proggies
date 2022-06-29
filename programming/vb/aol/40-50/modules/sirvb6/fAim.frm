VERSION 5.00
Begin VB.Form fAim 
   Caption         =   "AIM Examples"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   Icon            =   "fAim.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   3210
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   1830
   End
   Begin VB.TextBox Text1 
      Height          =   2745
      Left            =   3210
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Text            =   "fAim.frx":08CA
      Top             =   0
      Width           =   3015
   End
   Begin VB.CommandButton Command14 
      Caption         =   "AIM_SendRoom"
      Height          =   375
      Left            =   1590
      TabIndex        =   14
      Top             =   1440
      Width           =   1605
   End
   Begin VB.CommandButton Command13 
      Caption         =   "AIM_SendIM"
      Height          =   375
      Left            =   1590
      TabIndex        =   13
      Top             =   1080
      Width           =   1605
   End
   Begin VB.CommandButton Command12 
      Caption         =   "AIM_RoomLink"
      Height          =   375
      Left            =   1590
      TabIndex        =   12
      Top             =   720
      Width           =   1605
   End
   Begin VB.CommandButton Command11 
      Caption         =   "AIM_RoomEnter"
      Height          =   375
      Left            =   1590
      TabIndex        =   11
      Top             =   360
      Width           =   1605
   End
   Begin VB.CommandButton Command10 
      Caption         =   "AIM_KillAd"
      Height          =   375
      Left            =   1590
      TabIndex        =   10
      Top             =   0
      Width           =   1605
   End
   Begin VB.CommandButton Command9 
      Caption         =   "AIM_Ignore"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2850
      Width           =   1605
   End
   Begin VB.CommandButton Command8 
      Caption         =   "AIM_GetUser"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2490
      Width           =   1605
   End
   Begin VB.CommandButton Command7 
      Caption         =   "AIM_GetIMtext"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2130
      Width           =   1605
   End
   Begin VB.CommandButton Command6 
      Caption         =   "AIM_GetIMsn"
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   1605
   End
   Begin VB.CommandButton Command5 
      Caption         =   "AIM_GetChat"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   1605
   End
   Begin VB.CommandButton Command4 
      Caption         =   "AIM_ClearChat"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1605
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AIM_AntiPunt Off"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AIM_AddRoom"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1605
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1590
      TabIndex        =   1
      Top             =   1800
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AIM_AddBuddyList"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1605
   End
End
Attribute VB_Name = "fAim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'//  adds your aim buddy list into list1
AIM_AddBuddyList List1
End Sub

Private Sub Command10_Click()
'//  Kills the aim ad on your buddy list
AIM_KillAd
End Sub

Private Sub Command11_Click()
'//  Enters an AIM room, notice i also used AIM_GetUser, that is so it will
'//  Send YOU the invite, if you wanted to invite someone ELSE then you
'//  Could add a text box and let the user choose who they would like to
'//  Invite into the room.
AIM_RoomEnter AIM_GetUser, "invite test", "SiRvb6"
End Sub

Private Sub Command12_Click()
'//  Sends an AIM room link into the room, others simply click on the link
'//  And it takes them into the room SiRvb6
AIM_roomLink "SiRvb6"
End Sub

Private Sub Command13_Click()
'//  This will send an instant message to yourself, again if you would like to
'//  Send this to a different person you can use a text box with the persons SN
AIM_SendIM AIM_GetUser, "testing SiRvb6"
End Sub

Private Sub Command14_Click()
'//  Sends chat text into the room, this also shows you how to send the
'//  User's SN into a room so you can say who loaded the program
AIM_SendRoom "-=[ my test program"
AIM_SendRoom "-=[ loaded by: " & AIM_GetUser
End Sub

Private Sub Command2_Click()
'//  Adds all the people in an AIM chat room into list1
AIM_Addroom List1
End Sub

Private Sub Command3_Click()
If Command3.Caption = "AIM_AntiPunt Off" Then
    Timer1.Enabled = True
    Command3.Caption = "AIM_AntiPunt On"
    MsgBox "Anti Punter Is Now On"
Else
    Command3.Caption = "AIM_AntiPunt Off"
    Timer1.Enabled = False
    MsgBox "Anti Punter Is Now Off"
End If
End Sub

Private Sub Command4_Click()
'//  This will clear all the chat text in an AIM chat room
AIM_ClearChat
End Sub

Private Sub Command5_Click()
'//  This puts all the AIM chat text into text1, if you look at the AIM_GetChat
'//  Sub you will notice the StripHTML in it, well that is because AIM uses html
'//  In their chat rooms, its a pain in the ass if you don't have a good StripHTML
'//  Sub or function such as in SiRvb6.bas
Text1 = AIM_GetChat
End Sub

Private Sub Command6_Click()
'//  Returns the SN of the person who sent you the IM into text box 2
Text2 = AIM_GetIMsn
End Sub

Private Sub Command7_Click()
'//  Adds the AIM instant message text from the IM, and removes the leading
'//  Screen Name
Text1 = AIM_GetIMtext
End Sub

Private Sub Command8_Click()
'//  Gets the aim user's Screen Name from the aim buddy list caption
Text2 = AIM_GetUser
End Sub

Private Sub Command9_Click()
'//  This will ignore anyone by their Screen Name, just put the person who you
'//  Want to ignore into Text2 then push this button. (You cannot ignore yourself)
AIM_Ignore Text2
End Sub

Private Sub Form_Load()
'//  This sets this form as the top most window when it loads
StayOnTop fAim
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'//  cleans this form from the memory
Set fAim = Nothing
End Sub

Private Sub Form_Resize()
'//  Sometimes when you minimize a form, then restore it to its regular size
'//  It will remove it as the top most window, This will make sure that does
'//  Not happen
StayOnTop fAim
End Sub

Private Sub Form_Unload(Cancel As Integer)
'//  Unload the form
Unload fAim
End Sub

Private Sub List1_DblClick()
'//  When put in List1_DblClick, you just have to double click on the list item
'//  That you want to remove from the listbox
List_Remove List1
End Sub

Private Sub Timer1_Timer()
'//  This looks for punt or error strings in the chat room, do not set the timer too
'//  Fast or it will not have time to scan the room for the error strings, plus it may
'//  Hog alot of memory if its too fast to complete
AIM_antiPunt
End Sub
