VERSION 5.00
Object = "{DE8D4E3E-DD62-11D2-821F-444553540001}#1.0#0"; "CHATSCAN³.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   75
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1020
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   60
      Width           =   2595
   End
   Begin chatscan³.Chat Chat1 
      Left            =   2670
      Top             =   1935
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2730
      TabIndex        =   3
      ToolTipText     =   "visit www.george.cx!"
      Top             =   600
      Width           =   195
   End
   Begin VB.Label lAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "about ccom"
      Top             =   360
      Width           =   135
   End
   Begin VB.Label lQuit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "exit ccom"
      Top             =   135
      Width           =   135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ccom engine example by naïve (george huger)
'visit http://www.george.cx for more examples
'aim: lazy naive / cannibishop
'questions? email george@george.cx. make em good, dont waste my time!
'visit pr unity! thats my group!

'about:
'this is the shell (engine) for a ccom
'you can add you own cmds, i put some basic ones in so you see how it works.
'it takes care of the trigger and all that also
'enjoy! :)

'note: most ccoms have no iface. so in the form_load event you may want to put
'Me.hide
'personally i leave some iface so you can close the prog w/o aol being open

'please give credit to naïve if you build on this engine.


'Please visit my website! im going to make it cool soon, :)
'also if you are an artist and want to do some art for my site, i need it a lot, email me at george@george.cx

Dim ChatArgs() As String
Public ChatTrigger As String
Public MinsIdle As Integer

Public IdleRsn As String



Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
On Error GoTo errhandler
Dim TestLen As Integer
TestLen = Len(ChatTrigger)
If TestLen = 0 Then TestLen = 1
If Screen_Name <> GetUser Then Exit Sub 'make sure the user is giving the cmd
If Left(What_Said, TestLen) <> ChatTrigger And ChatTrigger <> "" Then Exit Sub 'check for the trigger



Dim ChatCommand As String, ArgStr As String
Dim Spot As Integer, NumArgs As Integer

Spot = InStr(1, What_Said, " ") ' find the end of the cmd/start of args

If Spot = 0 Then 'no args
ChatCommand = What_Said
ArgStr = ""
GoTo StartCCom
End If



ChatCommand = Left(What_Said, Spot - 1) 'put cmd in str
ArgStr = Mid(What_Said, Spot + 1) 'put args in str to split

'cut out spaces between args
ArgStr = Replace(ArgStr, ", ", ",")
ArgStr = Replace(ArgStr, " ,", ",")

'split the args into an array, show how many there are in numargs
ChatArgs = Split(ArgStr, ",")
NumArgs = UBound(ChatArgs)






StartCCom:


'okay now here is what you need to understand
'all of the arguments are now in an array called ChatArgs, and the actual command,w/o the trigger is in ChatCommand
'it is zero-based, so the first arg is stored in ChatArgs(0), the second in ChatArgs(1)
'if it looks for an arg that isnt there, an error will occur and the user will be told to check the arguments he passed.


If ChatTrigger <> "" Then ChatCommand = Mid(ChatCommand, Len(ChatTrigger) + 1) 'remove the trigger from the cmd

Select Case ChatCommand
'cmd code here

    'if you dont know, this is a simple case select format
    'i think its obvious how this works
    
    Case "im"
        Call InstantMessage(ChatArgs(0), ChatArgs(1))
                
    Case "email"
        Call SendMail(ChatArgs(0), ChatArgs(1), ChatArgs(2))
        
    Case "closechat"
        Call ChatSend("see you all later!")
        Call CloseWindow(FindRoom)
    
    Case "trig"
        ChatTrigger = ChatArgs(0)
        If ChatArgs(0) = "none" Then
        ChatTrigger = ""
        Call ChatSend("trigger removed")
        Else
        Call ChatSend("trigger set to: " & ChatArgs(0))
        End If
        
    Case "pr"
        PrivateRoom (ChatArgs(0))
        WaitForOKOrRoom (ChatArgs(0))
        E = GetCaption(FindRoom)
        Pause (1)
        If GetCaption(FindRoom) <> ChatArgs(0) Then
        ChatSend "the room: " & ChatArgs(0) & " is full."
        Else
        ChatSend "ccom engine ex by naive entered " & ChatArgs(0)
        End If
        
    Case "adv"
        ChatSend "ccom engine example by naive"
        ChatSend "http://www.george.cx"
        
 End Select

Exit Sub

errhandler:
ChatSend "ccom engine ex by naive - there was a general error. check your args."


End Sub

Private Sub Form_Load()
Chat1.ScanOn
FormOnTop Me
ChatSend "·!¡×¡!·  ccom ex by naive - loaded ·!¡×¡!·"
ChatSend "·!¡×¡!·  http://www.george.cx ·!¡×¡!·"

End Sub


Private Sub Form_Unload(Cancel As Integer)
ChatSend "·!¡×¡!·  ccom ex by naive - unloaded ·!¡×¡!·"
ChatSend "·!¡×¡!·  http://www.george.cx ·!¡×¡!·"
End

End Sub


Private Sub Picture1_Click()

End Sub


Private Sub timidle_Timer()
MinsIdle = MinsIdle + 1
ChatSend "fcc - idle: " & MinsIdle & " rsn: " & IdleRsn

End Sub


Private Sub Label1_Click()
On Error Resume Next
Shell ("c:\windows\command\start.exe http://www.george.cx")
End Sub


Private Sub lAbout_Click()
Dim msg As String, nl As String, msg2 As String
nl = Chr(13) & Chr(10)
msg = "com engine example by naïve" & nl & "coded friday july 28, 2000" & nl & "visit www.george.cx for more!" & nl & "peace to unity!"
msg2 = "about the ccom engine ex:" & nl & "i wrote this ex because i cracked and began writing a ccom. id typed mid, left, right, instr, etc too much and then i realized i could split and case the ccom. as most peeps dont know how to use split and many dont know how to use case, i wrote this ex. all you have to do now is fill in options, and as kid says, thats the part that sucks! :) good luck, send me copies of your finished progs and give me credit! also im writing flight ccom on this engine, so look out! - peace - naïve of unity"
Call MsgBox(msg & nl & nl & msg2 & nl & nl & "-this example uses dos32.bas and dos's chatscan control. - thanks! (www.dosfx.com)", , "about this ex.")

End Sub


Private Sub lQuit_Click()
Dim r
r = MsgBox("are you sure you want to exit?", vbYesNo, "exit?")
If r = vbYes Then End

End Sub

Private Sub pLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me

End Sub


