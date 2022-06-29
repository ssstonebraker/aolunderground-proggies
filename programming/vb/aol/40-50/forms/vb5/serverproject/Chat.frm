VERSION 5.00
Begin VB.Form frmChat 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3390
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5670
   ControlBox      =   0   'False
   HelpContextID   =   70
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5280
      TabIndex        =   3
      Top             =   2760
      WhatsThisHelpID =   70
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   2
      Top             =   2760
      WhatsThisHelpID =   70
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      WhatsThisHelpID =   70
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      WhatsThisHelpID =   70
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      X1              =   120
      X2              =   5520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Report Any Bugs Thank You "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      WhatsThisHelpID =   70
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "- •´)•–  X-Treme Server '99 By M Chat Serve •´)•–"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      WhatsThisHelpID =   70
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      BorderWidth     =   3
      Height          =   2175
      Left            =   120
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text2.SetFocus
If Text2 = "" Then Exit Sub
If Left(Text2, 1) = Chr$(13) Or Left(Text2, 1) = Chr$(10) Then Text2 = Mid$(Text2, 2)
If Left(Text2, 1) = Chr$(13) Or Left(Text2, 1) = Chr$(10) Then Text2 = Mid$(Text2, 2)
If Right(Text2, 1) = Chr$(13) Or Right(Text2, 1) = Chr$(10) Then Text2 = Left(Text2, Len(Text2) - 1)
If Right(Text2, 1) = Chr$(13) Or Right(Text2, 1) = Chr$(10) Then Text2 = Left(Text2, Len(Text2) - 1)
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & Text2)
X = SetFocusAPI(Me.hwnd)
Text2 = ""
End Sub

Private Sub Command2_Click()
MChatBot = False
Unload Me
End Sub

Private Sub Form_Load()

StayOnTop Me
CenterForm Me
MChatBot = True
'Server.Timer6.Enabled = True
Server.Chat1.ScanOn
Text1 = "                      –•´)•–" & " X-Treme Server '99" & " Chat & Serve –•(`•–" & Chr$(13) & Chr$(10)
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•X-Treme Server '99 M-Chat Enabled•–")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•X-Treme Server '99 M-Chat Off•–")
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1)
End Sub


