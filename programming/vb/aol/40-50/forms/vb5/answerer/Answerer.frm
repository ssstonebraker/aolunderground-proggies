VERSION 5.00
Begin VB.Form Answerer 
   Caption         =   "IM Answer"
   ClientHeight    =   3015
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   2490
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Answerer.frx":0000
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuRight 
      Caption         =   "Right"
   End
   Begin VB.Menu mnuLeft 
      Caption         =   "Left"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuStart 
      Caption         =   "Start"
   End
   Begin VB.Menu mnuStop 
      Caption         =   "Stop"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Answerer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuLeft_Click()
mnuLeft.Visible = False
mnuRight.Visible = True
Answerer.Width = 2610
End Sub

Private Sub mnuRight_Click()
mnuLeft.Visible = True
mnuRight.Visible = False
Answerer.Width = 5760
End Sub

Private Sub mnuStart_Click()
mnuStart.Visible = False
mnuStop.Visible = True
Timer1.Enabled = True
End Sub

Private Sub mnuStop_Click()
mnuStart.Visible = True
mnuStop.Visible = False
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
IM% = FindChildByTitle(FindAOLsMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(FindAOLsMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
List1.AddItem = ("snfromim")
List2.AddItem = ("messagefromim")
IMKeyword (SNfromIM), " " + Text1.Text
ClosePopUpIM ' if u don't close this IM window sumhow......it'll jus keep send'n um a IM
' make sure u make a ClosePopUpIM sub
End Sub
