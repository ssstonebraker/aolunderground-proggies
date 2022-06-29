VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Chat Command Example By Vendaz"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB5Chat2.Chat Chat1 
      Left            =   60
      Top             =   765
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.Label Label2 
      Caption         =   "For help email: vendaz@graffiti.net"
      Height          =   285
      Left            =   810
      TabIndex        =   1
      Top             =   510
      Width           =   2520
   End
   Begin VB.Label Label1 
      Caption         =   "Simple chat example by Vendaz"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   150
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
'Thanks for getting this file I hope it helps you
'learn how to make a ccom.
'It uses dos32.bas and vb5chat2.ocx, both made
'by dos, you can get them at http://www.hider.com/dos
'I find these are good to learn and very well made.
'Yes I can make my own modules (bas files) but
'I use UK aol (cause Im in London) and so
'so it tends not to work for US aol ;/
    If LCase(Screen_Name) = LCase(GetUser) And InStr(What_Said, ".") Then
    'This checks to see if the last person to talk
    'was the user of the program
        SpacePos& = InStr(What_Said$, " ")
        'This finds the position of the space in the
        'chat text
        If SpacePos& = 0 Then
            'If there is no space, just get the chat
            'text
            ChatCommand$ = What_Said$
        Else
            'if there is a space then just get the
            'text before the the space
            ChatCommand$ = Left(What_Said$, SpacePos& - 1&)
        End If
        Select Case ChatCommand$
        'Takes the case of what is said
            Case ".adv"
            'if the chat message is .adv then
                dos32.ChatSend "          <b>c</b>hat <b>c</b>ommand <b>e</b>xample"
                dos32.Pause (0.4)
                dos32.ChatSend "          <b>b</b>y <b>v</b>endaz - <b>u</b>ser " & LCase(GetUser)
            Case ".lol"
            'if the chat message is .lol then
                dos32.ChatSend GetUser & " <b>l</b>aughs <b>o</b>ut <b>l</b>oud"
            Case ".kick"
            'if the chat message is .kick then
                If Len(What_Said) > 6 Then
                'If the chat text is longer than 6
                'letters long then...
                    Person$ = Mid(What_Said, 7)
                    'gets the text after the .kick
                    dos32.ChatSend GetUser & " <b>k</b>icks " & Person$
                    'sends the users name then the
                    'person they kicked
                End If
            Case ".imsoff"
            'if chat message is .imsoff
                dos32.ChatSend GetUser & " <b>t</b>urns <b>i</b>ms <b>o</b>ff"
                'sends chat text
                dos32.IMsOff
                'turns ims off using dos's bas
            Case ".imson"
            'if chat message is .imson
                dos32.ChatSend GetUser & " <b>t</b>urns <b>i</b>ms <b>o</b>n"
                'sends chat text
                dos32.IMsOn
                'turns ims on using dos's bas
            Case ".end"
            'if chat message is .end
            Unload Form1
            'closes program
        End Select
        'ends the select
    End If
    'ends the first if statement
End Sub

Private Sub Form_Load()
    'when the form loads this is what is done
    Call MsgBox("This is a chat command example made by vendaz, I hope that this helps you to get to grips with the making of a ccom and please if you use this file put me in your greets.", vbInformation, "Information")
    'makes a message box
    dos32.ChatSend "          <b>c</b>hat <b>c</b>ommand <b>e</b>xample"
    dos32.Pause (0.4)
    dos32.ChatSend "          <b>b</b>y <b>v</b>endaz - <b>l</b>oaded <b>b</b>y " & LCase(GetUser)
    'advertize in the chat room
    Chat1.ScanOn
    'turns chat scan on
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'when the closes loads this is what is done
    Call MsgBox("Remember put me in the greets and dont just steal the code, whats the point?  If you need help Mail me at:  vendaz@graffiti.net", vbInformation, "Information")
    dos32.ChatSend "          <b>c</b>hat <b>c</b>ommand <b>e</b>xample"
    dos32.Pause (0.4)
    dos32.ChatSend "          <b>b</b>y <b>v</b>endaz - <b>u</b>nloaded <b>b</b>y " & LCase(GetUser)
    'advertize in the chat room
    Chat1.ScanOff
    'Turns chat scan off
    End
    'ends program
End Sub
