VERSION 5.00
Object = "{DE8D4E3E-DD62-11D2-821F-444553540001}#1.0#0"; "CHATSCAN³.OCX"
Begin VB.Form Form1 
   Caption         =   "inept's ccom example"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin chatscan³.Chat Chat1 
      Left            =   120
      Top             =   120
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
' mmm dos' chat scanner. it's great.
' ok, we're going to start off by making sure
' that the person entering the commands is the
' ccom user. if it is anyone else, we exit the
' chat1_chtmsg dec.

' the string 'screen_name' pulls the last screen name
' from the chatroom. 'what_said' pulls what was last
' said.

If Screen_Name = GetUser = False Then Exit Sub
' if the last screen name wasn't the ccom user
' then everything below is ignored.

' now the more important stuff. you'll have to pay
' close attention to this. ill try to make it as
' simple as possible. this is my first example ever
' so please, hang in there with me.

If LCase(What_Said) = ".adv" Then
' lower case what is said so that the ccom isnt
' case sensitive. this way the user can type
' .AdV or .Adv or .adv or .ADV

' now, what this is doing is pulling what was last
' said and checking to see if it is an actual
' command. in this case, if what was pulled is
' ".adv" then it will scroll an advertise for you
' chat command. forgive me if you were looking for
' a more complicated option. im trying to start
' out simple.

chatsend ". inept ccom example"
pause 0.6
chatsend ". for the ryze crew"
End If
' the end if. any time... ANY TIME you use if, you
' must close the option using end if. i hope you can
' see how ive done that.

' option number two would go right below the end if
If LCase(What_Said) = ".imson" Then
chatsend ". im enabled for recieval"
Call IMsOn
End If
' that seems simple enough. now, lets look at how
' to pull extra information. example: '.sup joebob'

If LCase(What_Said) Like ".sup *" Then
' important: you must change the = to like.
' this checks for anything LIKE .sup instead of
' anything equal to .sup. now, the star must be
' included so that the program knows where to stop
' the like function.

' now this gets complicated. i hope you guys are
' familiar with left, right, mid, and instr.

prsn$ = Mid((What_Said), 6)
' ok, prsn$ is going to equal the string following
' '.sup '. basically, replacing *. how this works,
' it finds the string (what_said) starts at the
' beginning of the string and counts over six.
' why 6? .(1)s(2)u(3)p(4) (5)*(6)
' i hope this explains it well enough. thats
' the best i can do.

chatsend "hey " & (prsn$) & " whats up?"
End If
' closed if.
End Sub

Private Sub Form_Load()
' im using modcrue .bas for this example because
' it seems to have all of the basic needed subs for
' any.. well.. basic program.

' start by enabling the chat scanner.
' in this example, ill use dos' aol chat scan
' for visual basic six. this .ocx file can also
' be used in vb5.

' enabling the chatscan is simple. simply
' type [scanner name].scanon. in this example,
' im going to use chat1 since that is the default
' name dos has decided to use. chat1.scanon

Chat1.scanon

chatsend ". inepts ccom example"
pause 0.6 'so the chatsend wont clutter
chatsend ". example for the ryze crew"
End Sub
