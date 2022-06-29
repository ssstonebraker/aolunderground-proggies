VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "welcome bot exanple    by scope"
   ClientHeight    =   240
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4395
   Height          =   645
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleWidth      =   4395
   Top             =   1170
   Width           =   4515
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   2640
      Top             =   0
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
    If Screen_Name = "OnlineHost" Then
        blah = Len(What_Said$)
        blah2 = blah - 22
        sn$ = Left$(What_Said$, blah2)
        Chat1.ChatSend ("Welcome " + sn$ + "Welcome")
    End If
End Sub


Private Sub Command1_Click()
Chat1.ScanOn
End Sub


Private Sub Command2_Click()
Chat1.ScanOff
End Sub


