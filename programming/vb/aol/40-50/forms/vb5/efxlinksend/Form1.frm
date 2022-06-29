VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EFX Link Example"
   ClientHeight    =   1200
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4590
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
   ScaleHeight     =   1200
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL Description:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL Link:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   780
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
      Begin VB.Menu break2 
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
'LINK SENDER EXAMPLE BY EFX
'Go to www.embersonfx.cjb.net for more stuff


Private Sub Command1_Click()
Form1.PopupMenu Form1.Links
End Sub

Private Sub Command2_Click()
Dim Ent As String * 2
Dim TmpStr As String

Ent = Chr(10) & Chr(13)
TmpStr = "Link Example By: Emberson Fx Program Design" & Ent & Ent
TmpStr = TmpStr & "Thank you for download Emberson Fx's Link Sender Example." & Ent & Ent
TmpStr = TmpStr & "FOR MORE VISUAL BASICS EXAMPLES AND HELP GO TO www.embersonfx.cjb.net" & Ent & Ent
TmpStr = TmpStr & "OR EMAIL US AT DrushDman@aol.com" & Ent & Ent
TmpStr = TmpStr & "The Module was not made by Emberson Fx" & Ent & Ent

MsgBox TmpStr, vbInformation, "About EFX Link Example"
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

Private Sub regular_Click()
ChatSend "< a href=" & Text1.Text & ">" & Text2.Text & "" & "</a>"
End Sub
