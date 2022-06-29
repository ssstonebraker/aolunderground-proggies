VERSION 5.00
Object = "{2FD385ED-D668-11D2-AAC2-44455354616F}#1.0#0"; "STEALHTML.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Http://www.Hider.com/"
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Steal Html"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin StealHtml.HtmlSteal HtmlSteal1 
      Left            =   0
      Top             =   0
      _ExtentX        =   4471
      _ExtentY        =   1508
   End
   Begin VB.Label Label1 
      Caption         =   $"OcxTestfrm.frx":0000
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   4320
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = HtmlSteal1.StealHtml(Text2.Text)
End Sub
