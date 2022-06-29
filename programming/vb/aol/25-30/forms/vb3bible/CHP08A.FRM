VERSION 2.00
Begin Form formGraphics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grapics Project"
   ClientHeight    =   3192
   ClientLeft      =   1140
   ClientTop       =   2472
   ClientWidth     =   4020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   3636
   Icon            =   CHP08A.FRX:0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   4020
   Top             =   2088
   Width           =   4140
   Begin CommandButton cmndQuit 
      Caption         =   "E&xit"
      Height          =   492
      Left            =   2640
      TabIndex        =   7
      Top             =   2520
      Width           =   1212
   End
   Begin CommandButton cmndActivate 
      Caption         =   "&Acitvate"
      Height          =   852
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1212
   End
   Begin Frame framBlanker 
      Caption         =   "Blanker Options"
      Height          =   2892
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2424
      Begin OptionButton optnBlanker 
         Caption         =   "Lines"
         Height          =   192
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   1932
      End
      Begin OptionButton optnBlanker 
         Caption         =   "Squares"
         Height          =   192
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2052
      End
      Begin OptionButton optnBlanker 
         Caption         =   "Circles"
         Height          =   192
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   2052
      End
      Begin OptionButton optnBlanker 
         Caption         =   "Color Spots"
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1932
      End
      Begin OptionButton optnBlanker 
         Caption         =   "Blank Screen"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1932
      End
   End
   Begin Image imagIcon 
      Height          =   1212
      Index           =   0
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1212
   End
End

Sub cmndActivate_Click ()
    formGraphics.Hide
    formScreen.Show
    formScreen!timeBlanker.Enabled = True
End Sub

Sub cmndQuit_Click ()
     End
End Sub

Sub Form_Load ()
     For i = 1 To 4            'create a control array of image controls
          Load imagIcon(i)
          imagIcon(i).Visible = True
     Next i
     'the next lines will not work if you've changed the default
     'directory setup in visual basic.
     On Error Resume Next
     imagIcon(0).Picture = LoadPicture("\VB\ICONS\ELEMENTS\MOON01.ICO")
     imagIcon(1).Picture = LoadPicture("\VB\ICONS\ELEMENTS\MOON05.ICO")
     imagIcon(2).Picture = LoadPicture("\VB\ICONS\MISC\MISC38.ICO")
     imagIcon(3).Picture = LoadPicture("\VB\ICONS\MISC\MISC36.ICO")
     imagIcon(4).Picture = LoadPicture("\VB\ICONS\MISC\MISC22.ICO")
End Sub

Sub optnBlanker_GotFocus (Index As Integer)
    imagIcon(Index).ZOrder BRINGTOFRONT
    Icon = imagIcon(Index).Picture
    BlankerType = Index
End Sub

