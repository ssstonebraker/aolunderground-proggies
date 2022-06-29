VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test"
   ClientHeight    =   3660
   ClientLeft      =   1200
   ClientTop       =   1515
   ClientWidth     =   6030
   Height          =   4065
   Left            =   1140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   6150
   Begin VB.CheckBox Check1 
      BackColor       =   &H000000FF&
      Caption         =   "add TESTing to IM"
      Height          =   615
      Left            =   3600
      TabIndex        =   21
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Send"
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   1095
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "Form1.frx":0000
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4680
      TabIndex        =   17
      Text            =   "GividenB"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Send"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Text            =   "3"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "Form1.frx":0008
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "TESTING"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "GividenB"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Advertise"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "i am a dood"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "http://w3.to/wer420inc  http://w3.to/wer420inc  http://w3.to/wer420inc  "
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   5775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TEST was made by eDiT to show people how to use a .bas file and to show how to use vb. Hope you found this useful!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   22
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   3480
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Text to Send"
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Person TO IM"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IM Stuff"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   240
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   3480
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject of Mail:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAIL STUFF"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CHAT STUFF"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Person To Mail:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   120
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text To Send:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
ChatSend Text1.Text
'This sends the text in text1.text to a chat room.
'Check out the advertise features button to get more sending options.
End Sub


Private Sub Command2_Click()
ChatSend "  <-----Test Toolz Loaded----->"
Pause 0.5
ChatSend "<-----User:" + GetUser + "----->"
Pause 0.5
ChatSend "  <-----Made As a Tester----->"

'OK, This shows a lot of things...
'Notice the "quotations around the stuff to send?
'That is becuase that is stuff u want to send
'When u send a text boxes text u dont use quotes
'You could also send a comboboxs text ex:
'ChatSend ComboBox1.text
'You can also send a labels caption...'
'The Pause 0.5 stalls the sending of the text
'You must do this if u want to send multiple lines of
'chat to a room, other wise a person gets logged off for scrolling
'
'In the second ChatSend it has "User:" + GetUser
'the GetUser is a function in dos32.bas
'So it sends to chat "User edit420" or whatever the users sn is
'You must have the quotes around the actual text, ex: "User:"
'
End Sub


Private Sub Command3_Click()

Dim X As Integer
X = 0
Do: DoEvents
X = X + 1
Call SendMail(Text2.Text, Text3.Text, Text4.Text)
Loop Until X = (Text5.Text)
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
'Sendmail(person, subject, msg)
'the loop until x is for how many times to send
'with this code u have a mail bomber
'
'now with the mail code u could also have it mail you
'Just SendMail("lolboblol@hehe.com", "TESTING", text3.text)
'That would email me with the subject "testing" and the msg text3
'Notice the quotes around my email and the subject
'This is just like the chat send
'Objects dont get qoutes and text the program sends in mails, ims, chat do

End Sub

Private Sub Command4_Click()
If Check1.Value = "1" Then
InstantMessage (Text7.Text), (text8.Text + "I am TESTing")
Else
InstantMessage (Text7.Text), (text8.Text)
End If
'OK thre if check1.value thing checks if it is checked.
'if it is then it sends what the person wants to send plus "TESTing"
'the else respresnces what happens if it aint checked
'and it just sends the im to the person with out "TESTing
'you cood make a punter with this
'EX:
'InstantMessage (Text7.Text), (text8.Text + "<font size=""9999999999999999999999999999999999999999999999999999999"">")
End Sub

Private Sub Form_Load()
FormOnTop Me
End Sub


