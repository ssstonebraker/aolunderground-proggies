VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Link Sender Example ²·º · By EcCo"
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
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
'Link Sender Example²·º  By EcCo
'X EcCo X@hotmail.com

Private Sub Command1_Click()
Form1.PopupMenu Form1.Links
End Sub

Private Sub Command2_Click()
Dim Ent As String * 2
Dim TmpStr As String

Ent = Chr(10) & Chr(13)
TmpStr = "Link Sender Example ²·º · By: EcCo" & Ent & Ent
TmpStr = TmpStr & "Sup?  Welcome to Link Sender Example 2.0.  In this version I added two more features, Non-Underlined and IM Links!  Enjoy!" & Ent & Ent
TmpStr = TmpStr & "~EcCo" & Ent & Ent
TmpStr = TmpStr & "E-Mail:" & Ent & "X EcCo X@hotmail.com" & Ent & Ent

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

Private Sub regular_Click()
ChatSend "< a href=" & Text1.Text & ">" & Text2.Text & "" & "</a>"
End Sub
