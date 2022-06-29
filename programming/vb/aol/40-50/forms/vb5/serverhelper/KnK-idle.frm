VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Idle Bot"
   ClientHeight    =   645
   ClientLeft      =   2040
   ClientTop       =   2070
   ClientWidth     =   915
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "KnK-idle.frx":0000
   ScaleHeight     =   645
   ScaleWidth      =   915
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'AOL4.o command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol4" Then
On Local Error Resume Next
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders Idle Bot  °´¯`×-»")
Timer1.Enabled = True
End If
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders Idle Bot  °´¯`×-»")
Timer1.Enabled = True
End If
End If
'AOL95 command
If aversion$ = "aol95" Then
AOLChatSend ("«-×´¯`°  KnK Founders Idle Bot  °´¯`×-»")
Timer1.Enabled = True
End If
End Sub


Private Sub Command2_Click()
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol4" Then
Timer1.Enabled = False
SendChat BlackGreenBlack("«-×´¯`°  I'm Back!!  °´¯`×-»")
End If
If aversion$ = "aol95" Then
Timer2.Enabled = False
AOLChatSend ("«-×´¯`°  I'm Back!!  °´¯`×-»")
End If
End Sub


Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Timer1_Timer()
DoEvents
If UserAOL = "AOL4" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders Idle Bot  °´¯`×-»")
TimeOut (60)
TimeOut (60)
End If
If UserAOL = "AOL95" Then
AOLChatSend ("«-×´¯`°  KnK Founders Idle Bot  °´¯`×-»")
TimeOut (60)
TimeOut (60)
End If
End Sub


Private Sub Timer2_Timer()


End Sub
