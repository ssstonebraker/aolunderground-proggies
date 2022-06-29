VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H80000007&
   Caption         =   "Guide Bot"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   ScaleHeight     =   855
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Scare People Even More"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "Scare people"
      Height          =   255
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If FindChatRoom() = "" Then
Kazoo = MsgBox("You must be in a chat room to use this function", vbCritical, "HyPO")
Exit Sub
End If

SendChat ("CATWATCH01 Has Entered The Room")
SendChat ("GuideBOB Has Entered The Room")
TimeOut 3
SendChat ("GuideMax Has Entered The Room")
SendChat ("SteveCase Has Entered The Room")
TimeOut 3
SendChat ("TosGeneral Has Entered The Room")
SendChat ("Catwatch02 Has Entered The Room")
TimeOut 3
SendChat ("GuideZag Has Entered The Room")
SendChat ("RNGBOB Has Entered The Room")
TimeOut 3
SendChat ("RNGJack Has Entered The Room")
SendChat ("Catwatch01 Has Entered The Room")
TimeOut 3
SendChat ("GuideOJG Has Entered The Room")
SendChat ("GuideJOG Has Entered The Room")
TimeOut 3
SendChat ("GuideUKby Has Entered The Room")
SendChat ("GuideHi Has Entered The Room")



End Sub

Private Sub TheView_Change()

End Sub

Private Sub Command2_Click()
Unload Form10

Form10.Hide

End Sub

Private Sub Command3_Click()
If FindChatRoom() = "" Then
Kazoo = MsgBox("You must be in a chat room to use this function", vbCritical, "HyPO")
Exit Sub
End If

SendChat ("CATWATCH01 Has Entered The Room")
TimeOut 2
SendChat ("CATWATCH01 Has left The Room")

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form10
 
End Sub
