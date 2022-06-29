VERSION 5.00
Object = "{DE8D4E3E-DD62-11D2-821F-444553540001}#1.0#0"; "CHATSC~1.OCX"
Begin VB.Form main 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin chatscan³.Chat Chat1 
      Left            =   600
      Top             =   360
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin VB.ListBox comz 
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   1620
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'you may edit this all you wan't just put me
'in on the credits as a teacher but you can say you made all this
'this is made so that people new to visual basic can learn how to make a
'c-com for aol chat rooms






Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
'this tells it to watch for the word .exit
If LCase(What_Said) = ".exit" Then
'this tell it to only listen to you
'if you take this line out it will listen to
'anyone that types in .exit
'You can also make it listen to only your s/n
'if some one else has this program and you sent
'it as a update just type it like this
'if screen_name="" then
'you have to have your s/n in the "" right or
'an easy way is to highlight your s/n in a
'chat room and then copy then paste it as a
'screen name
If Screen_Name = GetUser Then
'type in the "" what you want it to say when it ends
dos32.ChatSend "aol c-com v0.1 beta example by bdp"
'end makes the progream exit
End
'end if has to be here so that the program
'will work cause if has if's for the command
'they all have to have end if on each if  in
'the command you add-in
End If
End If

If LCase(What_Said) = ".cmds" Then
If Screen_Name = GetUser Then
'this will make the form pop up
main.Show
End If
End If
If LCase(What_Said) = ".ver" Then
If Screen_Name = GetUser Then
'this will type in the chat room the current version of aol you are using
dos32.ChatSend (dos32.GetUser) & " is using " & (dos32.aolversion)
End If
End If
If LCase(What_Said) = ".imon" Then
If Screen_Name = GetUser Then
dos32.IMsOn
End If
End If
If LCase(What_Said) = ".imoff" Then
If Screen_Name = GetUser Then
dos32.IMsOff
End If
End If
End Sub

Private Sub Form_DblClick()
'This makes the form hide
Me.Hide
End Sub

Private Sub Form_Load()
'makes prog stay on top of all other windows
dos32.FormTop Me
'This tells the chat room scaner
'(the ocx file that tells it how to read chat)
'to turn on
Chat1.ScanOn
'this adds the commands to the list box so
'that people can see the commands to work
comz.AddItem ".exit - exits program"
comz.AddItem ".cmds - views commands"
comz.AddItem ".ver - tells ver of aol"
comz.AddItem ".imon -  ims on"
comz.AddItem ".imoff - ims off"
'type in the "" below what you wan't it to say in the chat room when it loads
dos32.ChatSend "Aol c-com v0.1 beta example by bdp"
End Sub
