VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Link Sender Example ³·º · By EccO"
   ClientHeight    =   1170
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu Links 
      Caption         =   "Links"
      Visible         =   0   'False
      Begin VB.Menu regular 
         Caption         =   "Regular"
      End
      Begin VB.Menu nonunderlined 
         Caption         =   "Non-Underlined"
      End
      Begin VB.Menu faded 
         Caption         =   "Faded"
         Begin VB.Menu blabkblue 
            Caption         =   "Black - Blue"
         End
         Begin VB.Menu blackred 
            Caption         =   "Black - Red"
         End
         Begin VB.Menu break2 
            Caption         =   "-"
         End
         Begin VB.Menu randomfade 
            Caption         =   "Random Fade"
         End
      End
      Begin VB.Menu break3 
         Caption         =   "-"
      End
      Begin VB.Menu imlink 
         Caption         =   "IM Link Sender"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Link Sender Example³·º  By EccO
'xeccox@mailcity.com

Private Sub blabkblue_Click()
ChatSend "< a href=" & Text1.Text & ">" & BlackBlue("" & Text2.Text & "") & "</a>"
End Sub

Private Sub blackred_Click()
ChatSend "< a href=" & Text1.Text & ">" & blackred("" & Text2.Text & "") & "</a>"
End Sub

Private Sub Command1_Click()
Form1.PopupMenu Form1.Links
End Sub

Private Sub Command2_Click()
Dim Ent As String * 2
Dim TmpStr As String

Ent = Chr(10) & Chr(13)
TmpStr = "Link Sender Example ³·º · By: EccO" & Ent & Ent
TmpStr = TmpStr & "Sup?  Welcome to Link Sender Example 3.0.  Welcome to my newiest version of Link Sender Exampl!  In this version I added two more things more cool feature, faded links, and random faded links!  I decided not to add them for IM's because you autta do some programming, right?  Enjoy!" & Ent & Ent
TmpStr = TmpStr & "~EccO" & Ent & Ent
TmpStr = TmpStr & "E-Mail:" & Ent & "xeccox@mailcity.com" & Ent & Ent

MsgBox TmpStr, vbInformation, "About This Program..."
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
FormOnTop Me
End Sub

Private Sub imlink_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub nonunderlined_Click()
ChatSend "< a href=" & Text1.Text & "></u>" & Text2.Text & "" & "</a>"
End Sub

Private Sub randomfade_Click()
Dim X As Integer
    Dim X2
    Dim X3
    Dim X4
    Dim X5
    Dim X6
    Dim X7

X2 = BlackBlueBlack("" & Text2.Text & "")
X3 = BlackRedBlack("" & Text2.Text & "")
X4 = BlackGreenBlack("" & Text2.Text & "")
X5 = BlackPurpleBlack("" & Text2.Text & "")
X6 = BlackYellowBlack("" & Text2.Text & "")

        X = Int((Val(6) * Rnd) + 1)
        Select Case X

Case 1
ChatSend "< a href=" & Text1.Text & ">" & X2 & "</a>"
Case 2
ChatSend "< a href=" & Text1.Text & ">" & X3 & "</a>"
Case 3
ChatSend "< a href=" & Text1.Text & ">" & X4 & "</a>"
Case 4
ChatSend "< a href=" & Text1.Text & ">" & X5 & "</a>"
Case 5
ChatSend "< a href=" & Text1.Text & ">" & X6 & "</a>"

    End Select
End Sub

Private Sub regular_Click()
ChatSend "< a href=" & Text1.Text & ">" & Text2.Text & "" & "</a>"
End Sub
