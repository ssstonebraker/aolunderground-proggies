VERSION 5.00
Begin VB.Form frmWeb_Launch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Launcher by NightShade"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmWeb_Launch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblURL 
      Caption         =   "Click Here or the picture to launch web browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   345
      MouseIcon       =   "frmWeb_Launch.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Image imgURL 
      Height          =   1440
      Left            =   232
      MouseIcon       =   "frmWeb_Launch.frx":0614
      MousePointer    =   99  'Custom
      Picture         =   "frmWeb_Launch.frx":091E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5160
   End
End
Attribute VB_Name = "frmWeb_Launch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgURL_Click()

    Call WebLaunch("http://www.knk2000.com/knk/index2.html")

End Sub

Private Sub lblURL_Click()

    Call WebLaunch("http://www.knk2000.com/knk/index2.html")
    
End Sub
