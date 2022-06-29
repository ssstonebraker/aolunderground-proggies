VERSION 5.00
Begin VB.Form frmPics 
   Caption         =   "Form2"
   ClientHeight    =   2415
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Pics.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   Begin VB.PictureBox Next 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   7
      Left            =   3000
      Picture         =   "Pics.frx":030A
      ScaleHeight     =   960
      ScaleWidth      =   1440
      TabIndex        =   14
      Top             =   1440
      Width           =   1440
   End
   Begin VB.PictureBox Next 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   6
      Left            =   1560
      Picture         =   "Pics.frx":0B58
      ScaleHeight     =   960
      ScaleWidth      =   1440
      TabIndex        =   13
      Top             =   1440
      Width           =   1440
   End
   Begin VB.PictureBox Next 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   5
      Left            =   120
      Picture         =   "Pics.frx":13D2
      ScaleHeight     =   960
      ScaleWidth      =   1440
      TabIndex        =   12
      Top             =   1440
      Width           =   1440
   End
   Begin VB.PictureBox Next 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   4
      Left            =   4440
      Picture         =   "Pics.frx":1C3C
      ScaleHeight     =   960
      ScaleWidth      =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   1440
   End
   Begin VB.PictureBox Next 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   3
      Left            =   3000
      Picture         =   "Pics.frx":248A
      ScaleHeight     =   960
      ScaleWidth      =   1440
      TabIndex        =   10
      Top             =   480
      Width           =   1440
   End
   Begin VB.PictureBox Next 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   2
      Left            =   1560
      Picture         =   "Pics.frx":2CF4
      ScaleHeight     =   960
      ScaleWidth      =   1440
      TabIndex        =   9
      Top             =   480
      Width           =   1440
   End
   Begin VB.PictureBox Next 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   120
      Picture         =   "Pics.frx":356E
      ScaleHeight     =   960
      ScaleWidth      =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   1440
   End
   Begin VB.PictureBox Tetris0 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1800
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Tetris7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1560
      Picture         =   "Pics.frx":3DE8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Tetris6 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1320
      Picture         =   "Pics.frx":42FC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Tetris5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1080
      Picture         =   "Pics.frx":4828
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Tetris4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   840
      Picture         =   "Pics.frx":4D4C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Tetris3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   600
      Picture         =   "Pics.frx":5260
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Tetris2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      Picture         =   "Pics.frx":5784
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Tetris1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Picture         =   "Pics.frx":5CB0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'This form contains all the images used in the program
'-------------------------------------------------------

