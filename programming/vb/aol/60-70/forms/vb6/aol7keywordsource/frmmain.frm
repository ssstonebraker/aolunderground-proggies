VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Aol 7.0 Keyword (website) Example "
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "From Label Caption"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "www.aol.com"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "From Textbox"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fixed Text "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.aol.com"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Aol 7.0 Keyword Example By Source
'Visits a website via toolbar icon, keyword window textbox, then
'keyword window icon. All this is fully commented
'released : October 23rd, 2001 - 12:01 am. eastern time
'contact on aim: ciasource

Private Sub Command1_Click()
Call Keyword("www.aol.com")
End Sub

Private Sub Command2_Click()
Call Keyword("" + Text1.text + "")
End Sub

Private Sub Command3_Click()
Call Keyword("" + Label1.Caption + "")
End Sub
