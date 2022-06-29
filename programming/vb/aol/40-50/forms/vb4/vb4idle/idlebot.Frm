VERSION 4.00
Begin VB.Form Form2 
   Caption         =   "HoBBes's Idle Bot"
   ClientHeight    =   2175
   ClientLeft      =   3045
   ClientTop       =   3405
   ClientWidth     =   4710
   Height          =   2640
   Left            =   2985
   LinkTopic       =   "Form2"
   ScaleHeight     =   2175
   ScaleWidth      =   4710
   Top             =   3000
   Width           =   4830
   Begin VB.CommandButton Command1 
      Caption         =   "Idle Bot On"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Idle Bot Off"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "0"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Anything Below Here Should Be Invisible While the Bot Is Running"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Reason:"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = "0"
Text1.Text = "1"
ChatSend ("HoBBes's Idle Bot")
Call pause(".6")
ChatSend ("" + GetUser$ + " Is Now Idle")
Call pause(".6")
If Text3.Text = "" Then
ChatSend ("Reason:</font></font> N/A")
Else
ChatSend ("Reason:</font></font> " + Text3.Text + "")
End If
Text4.Text = Text3.Text
Call pause("1")
Text3.Text = ""
Call pause("59")
Do
ChatSend ("HoBBes's Idle Bot")
Call pause(".6")
ChatSend ("" + GetUser$ + " Is Now Idle For [" + Text1.Text + "] Mins")
Call pause(".6")
If Text4.Text = "" Then
ChatSend ("Reason:</font></font> N/A")
Else
ChatSend ("Reason:</font></font> " + Text4.Text + "")
End If
Text1.Text = Text1.Text + 1
Call pause("60")
Loop Until Text2.Text = "3"
End Sub


Private Sub Command2_Click()
Text2.Text = "3"
ChatSend ("HoBBes's Idle Bot")
Call pause(".6")
ChatSend ("Idle Bot Off")
Call pause(".6")
ChatSend ("" + GetUser$ + " Is No Longer Idle")
End Sub


