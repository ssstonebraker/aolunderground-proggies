VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form idler 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "a0de's idler example"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   2520
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   120
      Top             =   1320
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.CommandButton cstop 
      Caption         =   "stop"
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cstart 
      Caption         =   "start"
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox lmsgs 
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox tres 
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Text            =   "reason here"
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox tuser 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "handle here"
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "idler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hours As Integer
Private mins As Integer
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
'thanks for downloadin this afk bot example
'i made this because i've noticed alot of people
'are beginning to make afk bots/idlers
'and i see them get scrolled off everyday for it
'so, a couple of my friends asked me how to set a limit
'and this is how:
'                         a0de
Dim MSG As String, Space As Long
Space& = InStr(What_Said$, " ") 'sets up the lenght to begin the taken msg
ws$ = LCase(What_Said$) 'so u wont have to type what_said
sn$ = LCase(Screen_Name$) 'so you wont have to type screen_name

    If ws$ Like "-" & tuser.text & " *" Then
    Dim Hmany As Integer
    Hmany = 1
    If f2.lnames.ListCount = 0 Then GoTo Accept: 'prevents error
    i = f2.lnames.ListCount - 1
    f2.lnames.ListIndex = 0 'makes the selected index start at the top
    For X = 0 To i 'counts up to the list's listcount
    f2.lnames.ListIndex = X 'makes the selected index go further down
    If f2.lnames = sn$ Then Hmany = Hmany + 1
    Next X
    If Hmany = 4 Then Exit Sub

Accept:
    lmsgs.AddItem sn$ & ":  " & ws$ 'adds the name and msg to list
    f2.lnames.AddItem sn$ 'adds the name to another list for later use
    ChatSend sn$ & " thanks for leaving a msg " & Hmany & "/3"
    Pause 0.7
    End If
End Sub

Private Sub cstart_Click()
mins = 0
hours = 0
Timer1.Enabled = True
Chat1.ScanOn
ChatSend tuser.text & "is going idler. rsn: " & tres.text
Pause 0.5
ChatSend "type '-" + tuser.text + "' and a message"
End Sub

Private Sub cstop_Click()
Timer1.Enabled = False
Chat1.ScanOff
ChatSend tuser.text & " is back, " & lmsgs.ListCount & "/msg's"
End Sub

Private Sub Form_Load()
FormOnTop Me
Call MsgBox("Hey, i hope this example helpz u out alot and stops u from gettin logged off...i just hope u take the time to READ the code and TRY to UNDERSTAND IT before copying and pasting it into a new project ;\. As you can see all i made was a idler example because i dont feel comforatable givin out usefull stuff cuz thats where the copying comes in...but imma make a couple more examples for KnK, because i too was a beginner and KnK was in my favorites list ;P" & Chr$(13) & Chr$(13) & " Copyright© a0de 2000 - a0de.cjb.net", _
vbOKOnly, "Read Now..")
ChatSend "-› afk bot example¹ · by a0de"
End Sub

Private Sub Timer1_Timer()
mins = mins + 1
If mins = 60 Then
hours = hours + 1
mins = 0
End If
ChatSend tuser.text & " afk for: " & hours & ":" & mins & ""
Pause 0.5
ChatSend "leave a msg '-" + tuser.text + "' and a msg"
End Sub
