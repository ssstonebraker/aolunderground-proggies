VERSION 5.00
Begin VB.Form fAol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aol Example"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "fAol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command29 
      Caption         =   "AOLwwwLink"
      Height          =   345
      Left            =   0
      TabIndex        =   32
      Top             =   4770
      Width           =   4185
   End
   Begin VB.CommandButton Command28 
      Caption         =   "AolChildByTitle"
      Height          =   345
      Left            =   2190
      TabIndex        =   31
      Top             =   4440
      Width           =   1995
   End
   Begin VB.CommandButton Command27 
      Caption         =   "AOL4KW"
      Height          =   345
      Left            =   0
      TabIndex        =   30
      Top             =   4440
      Width           =   2205
   End
   Begin VB.CommandButton Command26 
      Caption         =   "AOL_Wait4Ok"
      Height          =   345
      Left            =   2190
      TabIndex        =   29
      Top             =   4110
      Width           =   1995
   End
   Begin VB.CommandButton Command25 
      Caption         =   "AOL_Version4"
      Height          =   345
      Left            =   0
      TabIndex        =   28
      Top             =   4110
      Width           =   2205
   End
   Begin VB.CommandButton Command24 
      Caption         =   "AOL_UpChatOn"
      Height          =   375
      Left            =   2190
      TabIndex        =   27
      Top             =   3750
      Width           =   1995
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   4200
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   3540
      Width           =   2895
   End
   Begin VB.CommandButton Command23 
      Caption         =   "AOL_SignOnAs"
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   3750
      Width           =   2205
   End
   Begin VB.CommandButton Command22 
      Caption         =   "AOL_SendRoom"
      Height          =   345
      Left            =   2190
      TabIndex        =   24
      Top             =   3420
      Width           =   1995
   End
   Begin VB.CommandButton Command21 
      Caption         =   "AOL_SendMail"
      Height          =   345
      Left            =   0
      TabIndex        =   23
      Top             =   3420
      Width           =   2205
   End
   Begin VB.CommandButton Command20 
      Caption         =   "AOL_SendIM"
      Height          =   375
      Left            =   2190
      TabIndex        =   22
      Top             =   3060
      Width           =   1995
   End
   Begin VB.CommandButton Command19 
      Caption         =   "AOL_RoomBust"
      Height          =   345
      Left            =   2190
      TabIndex        =   21
      Top             =   2730
      Width           =   1995
   End
   Begin VB.CommandButton Command18 
      Caption         =   "AOL_MailCount"
      Height          =   345
      Left            =   2190
      TabIndex        =   20
      Top             =   2400
      Width           =   1995
   End
   Begin VB.CommandButton Command17 
      Caption         =   "AOL_MailByIcon"
      Height          =   375
      Left            =   2190
      TabIndex        =   19
      Top             =   2040
      Width           =   1995
   End
   Begin VB.CommandButton Command16 
      Caption         =   "AOL_LocateSN"
      Height          =   375
      Left            =   2190
      TabIndex        =   18
      Top             =   1680
      Width           =   1995
   End
   Begin VB.CommandButton Command15 
      Caption         =   "AOL_KillWait"
      Height          =   375
      Left            =   2190
      TabIndex        =   17
      Top             =   1320
      Width           =   1995
   End
   Begin VB.CommandButton Command14 
      Caption         =   "AOL_IMsOn"
      Height          =   345
      Left            =   2190
      TabIndex        =   16
      Top             =   990
      Width           =   1995
   End
   Begin VB.CommandButton Command13 
      Caption         =   "AOL_IMsOff"
      Height          =   345
      Left            =   2190
      TabIndex        =   15
      Top             =   660
      Width           =   1995
   End
   Begin VB.CommandButton Command12 
      Caption         =   "AOL_im_txt"
      Height          =   345
      Left            =   2190
      TabIndex        =   14
      Top             =   330
      Width           =   1995
   End
   Begin VB.CommandButton Command11 
      Caption         =   "AOL_im_sn"
      Height          =   345
      Left            =   2190
      TabIndex        =   13
      Top             =   0
      Width           =   1995
   End
   Begin VB.CommandButton Command10 
      Caption         =   "AOL_Ignore"
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3060
      Width           =   2205
   End
   Begin VB.CommandButton Command9 
      Caption         =   "AOL_GetUser"
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   2730
      Width           =   2205
   End
   Begin VB.CommandButton Command8 
      Caption         =   "AOL_FindRoom"
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   2205
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   5670
      TabIndex        =   9
      Top             =   4080
      Width           =   1425
   End
   Begin VB.CommandButton Command7 
      Caption         =   "AOL_FindBuddyPR"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   2205
   End
   Begin VB.CommandButton Command6 
      Caption         =   "AOL_Find_a_Chat"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   2205
   End
   Begin VB.TextBox Text1 
      Height          =   3525
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "fAol.frx":08CA
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "AOL_ChatView"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   2205
   End
   Begin VB.CommandButton Command4 
      Caption         =   "AOL_AddRoom"
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   990
      Width           =   2205
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AOL_AddMemberDirectory"
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   660
      Width           =   2205
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AOL_AddMailbox"
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   2205
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   4200
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AOL_AddLB"
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "fAOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aolicon As Long
Dim Setup As Long
Dim EditBuddy As Long
Call AOL_BuddySetup  '//  clicks the Setup icon on your aol buddylist
Pause 2  '//  ensure enough time to wait, just to be safe
Setup& = AOLChildByTitle("'s Buddy Lists")  '//  Finds the child by its title
    aolicon& = FindWindowEx(Setup&, 0&, "_aol_icon", vbNullString) '//  Create Icon
    aolicon& = FindWindowEx(Setup&, aolicon&, "_aol_icon", vbNullString)  '//  Edit Icon
    ClickIt aolicon&  '//  Clicks the edit icon
Pause 2  '//  ensure enough time for the edit list form to show up
    AOL_AddLB "Edit List", List1  '//  Find the aol window by its title, then add it to list1
End Sub

Private Sub Command10_Click()
'//  You musst be in a chat room to use this
'//  In this example i added the room into list1, then i am having Person$
'//  be the 4th person in list1 (so you must be in a room with more than 4 people :P)
'//  to use this in your program, you can just have the person click a list item
'//  or use chat commands to ignore the person, This ignores the person in Text1
Dim Person As String
AOL_AddRoom List1
Person$ = List1.List(3)
Text1 = Person$
AOL_ignore Text1
End Sub

Private Sub Command11_Click()
'//  gets the SN from the instant message
Text1 = AOL_im_sn
End Sub

Private Sub Command12_Click()
'//  Gets the instant message text message without the screen name
Text1 = AOL_im_txt
End Sub

Private Sub Command13_Click()
'//  turns your aol instant messages off
'//  you can also do this by using AOL_SendIM this is just a shorter way
AOL_IMsOFF
End Sub

Private Sub Command14_Click()
'//  turns your aol instant messages off
'//  you can also do this by using AOL_SendIM this is just a shorter way
AOL_IMsON
End Sub

Private Sub Command15_Click()
'//  Kill that god awful hourglass icon that sometimes freezes on aol
AOL_KillWait
End Sub

Private Sub Command16_Click()
Dim x As Long
List1.AddItem "tosadvisor"
List1.AddItem "vib6"
List1.AddItem "IHI IE IHI"
Text1 = ""
AOL_LocateSN List1, List2
For x = 0 To List2.ListCount - 1
    Text1 = Text1 & List2.List(x) & vbCrLf
Next x
End Sub

Private Sub Command17_Click()
'//  opens up your mailbox, change the 0 to a 1 and it will open up a blank email
AOL_MailByIcon 0&
End Sub

Private Sub Command18_Click()
Text1 = AOL_MailCount
End Sub

Private Sub Command19_Click()
AOL_RoomBust "coderz"
End Sub

Private Sub Command2_Click()
AOL_MailByIcon (0)  '//  clicks the mailbox icon
Pause 3.2   '//  let the aol listbox load all the contents
AOL_AddMailBox List1  '//  add mail into lists1
End Sub

Private Sub Command20_Click()
'//  For testing purposes it sends the IM to yourself, change aol_getuser to any
'//  Text box or string to send an IM to
AOL_SendIM AOL_GetUser, ":) works"
End Sub

Private Sub Command21_Click()
'//  Send email to anyone, just replace aol_getuser with whatever name you want
'//  To send email to.
AOL_SendMail AOL_GetUser, "bas test", "just testing"
End Sub

Private Sub Command22_Click()
'//  Simulates another user opening your program, this is just a test of the
'//  AOL_SendRoom
AOL_SendRoom "-=[ loaded by: " & AOL_GetUser
End Sub

Private Sub Command23_Click()
'//  For this sign on another one of your screen names, put the screen name
'//  To sign on as into text1, then the screen name's password into text2
'//  This is NOT a password stealer so don't worry about the security of your PW, it is safe
AOL_SignOnAs Text1, Text2
End Sub

Private Sub Command24_Click()
If Command24.Caption = "AOL_UpChatOn" Then
    AOL_UpChatOn  '//  minimizes aol's upload screen so you can chat as normal
    Command24.Caption = "AOL_UpChatOff"  '//  changes the caption of the command
Else
    Command24.Caption = "AOL_UpChatOn"  '//  changes back the caption of the command
    AOL_UpChatOff     '//  restores the modal to its normal size
End If
End Sub

Private Sub Command25_Click()
'//  Checks to see which aol version the person is on, if it is aol 4.0 it will return
'//  True, if it is anything OTHER than aol4 it will return False
MsgBox "True or False, You are on aol4: " & vbCrLf & vbCrLf & "Answer: [ " & AOL_Version4 & " ]"
End Sub

Private Sub Command26_Click()
AOL4KW "aol://1391:Testing Purposes" '//  Just a fake KW to show you how to use aol_wait4ok
AOL_Wait4OK  '//  waits for an aol message box then kills it
End Sub

Private Sub Command27_Click()
'//  Just enter the keyword or web page you want to go to
AOL4KW "aol://5863:126/mBLA:283750"
End Sub

Private Sub Command28_Click()
Dim result As Long
result& = AOLChildByTitle("Buddy List")
If result& <> 0 Then
    MsgBox "Buddy List was found, the result is: " & result&
Else
    MsgBox "Buddy List was not found"
End If
End Sub

Private Sub Command29_Click()
'//  must be in a room to use this, it will send a HTML Hyperlink into aol4 chat
AOLwwwLink "http://www.escrambler.com", "download the best aol4 scrambler"
End Sub

Private Sub Command3_Click()
Dim memdir As Long
AOL4KW "aol://4950:0000010000|all:VB"  '//  Search the MemDir for profiles that have VB in it
    Do
      DoEvents
      memdir& = AOLChildByTitle("Member Directory Search Results")
    Loop Until memdir& <> 0   '//  loops until the search results are found
  Pause 1.9  '//  allow the listbox to fill up
AOL_AddMemberDirectory List1  '//  add the search results into list1 NOTE: this only adds the screen name, it strips the other garbage from the listbox
End Sub

Private Sub Command4_Click()
'//  Add aol's room listbox into a list on your own form
AOL_AddRoom List1
End Sub

Private Sub Command5_Click()
Dim sTemp As String
  sTemp$ = GetAPIText(AOL_ChatView())  '// uses Getapitext function and aol_chatview to get chat text on aol4
Text1 = Replace(sTemp$, vbCr, vbCrLf) '//  replaces the carriage return to carriage return & linefeed, so it looks like normal chat
End Sub

Private Sub Command6_Click()
'//  must be in a chat room to use this
If AOL_FindRoom = 0 Then  '//  makes sure the person is in a room
    MsgBox "Please Enter A Chat Room And Try Again"
Else
    AOL_Find_a_Chat "scrambler"   '//  searches aol's chatrooms for a scrambler room
End If
End Sub

Private Sub Command7_Click()
List1.AddItem "scrambler2"  '//  the list needs rooms in it so i just put
List1.AddItem "vb5"             '//  these in it just for this tutorial
AOL_FindBuddyPR List1, List2, AOL_GetUser  '//  This will find yourself in the chat room, replace aol_getuser with someone you would like to locate
List1.Clear  '//  you don't really need these, i just cleared them for this tutorial
List2.Clear  '//  you don't really need these, i just cleared them for this tutorial
End Sub

Private Sub Command8_Click()
'//  finds the aol chatroom, even if its not the top window
Text1 = AOL_FindRoom
End Sub

Private Sub Command9_Click()
'//  puts your screen name into text1
Text1 = AOL_GetUser
End Sub

Private Sub Form_Load()
'//  make fAOL the top window
StayOnTop fAOL
End Sub

