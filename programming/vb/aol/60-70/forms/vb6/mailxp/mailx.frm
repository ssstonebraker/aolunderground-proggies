VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MAIL SENDER X-AMPLE BY FuSe"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SEND"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "mailx.frx":0000
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "SUBJECT"
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "TO WHO"
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "fusemailz@aol.com"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "mail sender example by FuSe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call AOL6_MailSend("" + Text1 + "", "" + Text2.text + "", "" + Text3 + "")
End Sub
