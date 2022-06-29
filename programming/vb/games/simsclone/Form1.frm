VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CitySimulator"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Blank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   129
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer BirdTimer 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer TILEDATAUPDATE 
      Interval        =   5000
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer MainDrawTimer 
      Interval        =   50
      Left            =   840
      Top             =   120
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   11
      Left            =   5280
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   127
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   10
      Left            =   5280
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   126
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   9
      Left            =   5280
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   125
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   8
      Left            =   5280
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   124
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   7
      Left            =   5280
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   123
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   6
      Left            =   5280
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   122
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   118
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   117
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox LBLPPOP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   210
      ScaleWidth      =   2115
      TabIndex        =   110
      Top             =   2520
      Width           =   2145
   End
   Begin VB.PictureBox LBLLYI 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5880
      Picture         =   "Form1.frx":00BA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   109
      Top             =   3120
      Width           =   2145
      Begin VB.Line Line1 
         X1              =   0
         X2              =   144
         Y1              =   28
         Y2              =   28
      End
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   2760
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   45
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   2880
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   41
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   2880
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   37
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   2880
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   33
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2520
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   46
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2640
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   42
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2640
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   38
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2640
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   34
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   10
      Left            =   2280
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   47
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   10
      Left            =   2400
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   43
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   10
      Left            =   2400
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   39
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   10
      Left            =   2400
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   35
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   11
      Left            =   2040
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   48
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   11
      Left            =   2160
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   44
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   11
      Left            =   2160
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   40
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   11
      Left            =   2160
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   36
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   1800
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   32
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   1560
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   31
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   1320
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   30
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   1080
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   29
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   840
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   28
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   600
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   27
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   26
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFFall 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   25
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   1920
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   24
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   1680
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   23
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   1440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   22
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   1200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   21
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   960
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   20
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   19
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   18
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSummer 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   240
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   17
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   1920
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   16
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   1680
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   15
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   1440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   14
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   1200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   960
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   12
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   11
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   10
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFSpring 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   240
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   9
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   1920
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   8
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   1680
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   1440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   1200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   5
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   960
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox TURFWinter 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   240
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Caption         =   "Status"
      Height          =   2415
      Left            =   5880
      TabIndex        =   91
      Top             =   30
      Width           =   2175
      Begin VB.Label lblZNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Overlook Appartments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Zone Title"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   98
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblLV 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$100"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ Land Value $"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   96
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblGrowth 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   101
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Growth"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   93
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblSPOP 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   99
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pop."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Residential"
         Height          =   255
         Left            =   960
         TabIndex        =   94
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Zone"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox BlackBird 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   5280
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   111
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   112
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   113
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   114
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BirdMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   115
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame5 
      Caption         =   "FreeBes"
      Height          =   960
      Left            =   5880
      TabIndex        =   104
      Top             =   4090
      Width           =   2175
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   36
         Left            =   840
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   108
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   35
         Left            =   600
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   107
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   34
         Left            =   360
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   106
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   33
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   105
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Road $2000"
      Height          =   1335
      Left            =   120
      TabIndex        =   77
      Top             =   5060
      Width           =   1095
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   20
         Left            =   720
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   102
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   19
         Left            =   720
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   78
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   14
         Left            =   720
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   79
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   18
         Left            =   360
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   83
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   13
         Left            =   360
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   80
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   17
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   82
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   12
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   81
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   16
         Left            =   360
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   84
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   11
         Left            =   360
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   87
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   15
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   86
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   10
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   85
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buy Small $90,000"
      Height          =   1335
      Left            =   1260
      TabIndex        =   49
      Top             =   5060
      Width           =   1575
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   1200
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   59
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   960
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   58
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   720
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   57
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   480
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   56
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   1200
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   55
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   960
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   54
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   720
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   53
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   480
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   52
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   720
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   51
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   480
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   50
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Ind."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Com."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Res."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buy BIG $300,000"
      Height          =   1335
      Left            =   2880
      TabIndex        =   63
      Top             =   5060
      Width           =   5175
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   32
         Left            =   4080
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   120
         Top             =   780
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   31
         Left            =   3600
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   119
         Top             =   780
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   30
         Left            =   4560
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   76
         Top             =   300
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   29
         Left            =   4080
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   75
         Top             =   300
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   28
         Left            =   3600
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   74
         Top             =   300
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   27
         Left            =   1920
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   73
         Top             =   780
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   26
         Left            =   1440
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   72
         Top             =   780
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   25
         Left            =   960
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   71
         Top             =   780
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   24
         Left            =   2400
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   70
         Top             =   300
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   23
         Left            =   1920
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   69
         Top             =   300
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   22
         Left            =   1440
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   68
         Top             =   300
         Width           =   450
      End
      Begin VB.PictureBox Selector 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   21
         Left            =   960
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   67
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Safety"
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   121
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Industrial"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   66
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Commercial"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   65
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Residential"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.PictureBox pbSCREEN 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   0
      Top             =   120
      Width           =   5710
   End
   Begin VB.PictureBox BGPB2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   128
      Top             =   120
      Width           =   5710
   End
   Begin VB.PictureBox BGPB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   88
      Top             =   120
      Visible         =   0   'False
      Width           =   5710
   End
   Begin VB.PictureBox Buffer2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   116
      Top             =   120
      Visible         =   0   'False
      Width           =   5710
   End
   Begin VB.Label lblINCOME2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   5880
      TabIndex        =   103
      Top             =   3840
      Width           =   2145
   End
   Begin VB.Label lblINCOME 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Left            =   5880
      TabIndex        =   90
      Top             =   3600
      Width           =   2145
   End
   Begin VB.Label lblPOP 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Left            =   5880
      TabIndex        =   89
      Top             =   2760
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RNDS As Integer
Sub GenerateTURF()
For i = 1 To 4
For ii = 0 To 11
For iii = 0 To 13
For iiii = 0 To 13
rn = Int((Rnd * 3) + 0.5)
Select Case i
Case 1
Select Case rn
Case 1: TURFWinter(ii).PSet (iii, iiii), &HFFFFFF
Case 2: TURFWinter(ii).PSet (iii, iiii), &HE0E0E0
Case 3: TURFWinter(ii).PSet (iii, iiii), &HFFFFC0
End Select
Case 2
Select Case rn
Case 1: TURFSpring(ii).PSet (iii, iiii), &H8080FF
Case 2: TURFSpring(ii).PSet (iii, iiii), &H8000&
Case 3: TURFSpring(ii).PSet (iii, iiii), &HC000&
End Select
Case 3
Select Case rn
Case 1: TURFSummer(ii).PSet (iii, iiii), &HC000&
Case 2: TURFSummer(ii).PSet (iii, iiii), &H8000&
Case 3: TURFSummer(ii).PSet (iii, iiii), &H4000&
End Select
Case 4
Select Case rn
Case 1: TURFFall(ii).PSet (iii, iiii), &H77B2&
Case 2: TURFFall(ii).PSet (iii, iiii), &H9DD5&
Case 3: TURFFall(ii).PSet (iii, iiii), &H5BBD&
End Select
End Select
Next
Next
Next
Next
End Sub

Private Sub BLTSELECTORS_Timer()
End Sub

Private Sub BirdTimer_Timer()
If Not CurrentSeason = 1 Then NewBird
rn2 = Rnd * 100
If rn2 > 49 Then FreakChangeInDirection
BirdTimer.Interval = RndRange(250, 2000)
End Sub

Private Sub Form_Load(): CURL = 1: CURC = 2
W = pbSCREEN.ScaleWidth
H = pbSCREEN.ScaleHeight

Math_BTT
initTILES
GenerateTURF
DrawBacks
Selector_Click 0
fileload

If CurrentSeason = 0 Then initTILES
DrawBoard

Form1.Caption = "CitySimulation - $" & Cash & "  " & ReturnMstr(CurMonth) & "  " & CurYear

DrawSelectors
Me.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
filesave
End
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub lblGrowth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub lblINCOME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub lblINCOME2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub lblLV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub lblPOP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub lblSPOP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub lblZNum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub lblZone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub MainDrawTimer_Timer()
'skipBlt: DoEvents
'If Drawing Then GoTo skipBlt
BitBlt Buffer2.hDC, 0, 0, W, H, GFX.Turf(CurrentSeason).hDC, 0, 0, SRCCOPY

BitBlt Buffer2.hDC, 0, 0, W, H, BGPB2.hDC, 0, 0, SRCAND
BitBlt Buffer2.hDC, 0, 0, W, H, BGPB.hDC, 0, 0, SRCPAINT

If Not MouseOUT Then DoCURS Int(CX), Int(CY), Buffer2, CURC, CURL, CURS
DoBird

'Final Draw...
BitBlt pbSCREEN.hDC, 0, 0, W, H, Buffer2.hDC, 0, 0, SRCCOPY
Form1.Caption = "CitySimulation - $" & Cash & "  " & ReturnMstr(CurMonth) & " " & CurYear & "   Crime:" & Crime
End Sub

Private Sub pbSCREEN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): On Error Resume Next
CURL = 3
NX1 = Int(X / 13)
NY1 = Int(Y / 13)

If Button = 1 Then
If CURC = 1 Then Exit Sub

'Exit if purchase is not possible.
If Cash < MS.price Then Exit Sub

If Not T(NX1, NY1).StructureID > 9 And Not T(NX1, NY1).StructureID <= 20 Then

If MS.selectedPurchase > 20 And MS.selectedPurchase <= 32 Then
    If Not T(NX1, NY1).StructureID = 100 Then Exit Sub
    If Not T(NX1 + 1, NY1).StructureID = 100 Then Exit Sub
    If Not T(NX1, NY1 + 1).StructureID = 100 Then Exit Sub
    If Not T(NX1 + 1, NY1 + 1).StructureID = 100 Then Exit Sub
Else
    If Not T(NX1, NY1).StructureID = 100 Then Exit Sub
End If

End If

T(NX1, NY1).StructureID = MS.selectedPurchase

If T(NX1, NY1).StructureID <= 1 Then T(NX1, NY1).ColorFlag = Int(Rnd * 3): T(NX1, NY1).Name = RandResidence & " Residence"
If T(NX1, NY1).StructureID = 33 Then T(NX1, NY1).ColorFlag = Int(Rnd * 2)
If T(NX1, NY1).StructureID = 33 Or T(NX1, NY1).StructureID = 34 Then T(NX1, NY1).Name = "Trees"
If T(NX1, NY1).StructureID = 35 Then T(NX1, NY1).Name = "Bus Station"
If T(NX1, NY1).StructureID = 36 Then T(NX1, NY1).ClassFlag = Rnd * 2: T(NX1, NY1).Name = "Mini-Golf"
If T(NX1, NY1).StructureID > 1 And T(NX1, NY1).StructureID <= 5 Then If Not T(NX1, NY1).StructureID = 4 Then T(NX1, NY1).Name = RandSmallCom Else T(NX1, NY1).Name = "Public Pool"
If T(NX1, NY1).StructureID > 5 And T(NX1, NY1).StructureID <= 9 Then T(NX1, NY1).Name = RandSmallInd
If T(NX1, NY1).StructureID > 20 And T(NX1, NY1).StructureID <= 24 Then T(NX1, NY1).Name = RandBigRes
If T(NX1, NY1).StructureID > 24 And T(NX1, NY1).StructureID <= 27 Then T(NX1, NY1).Name = RandBigCom
If T(NX1, NY1).StructureID > 27 And T(NX1, NY1).StructureID <= 30 Then T(NX1, NY1).Name = RandBigInd
If T(NX1, NY1).StructureID = 31 Then T(NX1, NY1).Name = "The Fuzz"
If T(NX1, NY1).StructureID = 32 Then T(NX1, NY1).Name = "Fire Dept."

If T(NX1, NY1).StructureID > 20 And T(NX1, NY1).StructureID <= 32 Then
T(NX1 + 1, NY1).StructureID = 200
T(NX1 + 1, NY1 + 1).StructureID = 200
T(NX1, NY1 + 1).StructureID = 200
T(NX1 + 1, NY1).Name = T(NX1, NY1).Name
T(NX1 + 1, NY1 + 1).Name = T(NX1, NY1).Name
T(NX1, NY1 + 1).Name = T(NX1, NY1).Name
End If

If Not T(NX1, NY1).StructureID > 9 Then T(NX1, NY1).Growth = 1
If Not T(NX1, NY1).StructureID <= 20 Then T(NX1, NY1).Growth = 2
Cash = Cash - MS.price


End If

If Button = 2 Then
If T(NX1, NY1).StructureID = 100 Or T(NX1, NY1).StructureID = 200 Then Exit Sub
Cash = Cash + T(NX1, NY1).LandValue

If T(NX1, NY1).StructureID > 20 And T(NX1, NY1).StructureID <= 32 Then
T(NX1 + 1, NY1).StructureID = 100
T(NX1 + 1, NY1 + 1).StructureID = 100
T(NX1, NY1 + 1).StructureID = 100      'Kill 4 tile structure
T(NX1 + 1, NY1).Name = "Open Space"
T(NX1 + 1, NY1 + 1).Name = "Open Space"
T(NX1, NY1 + 1).Name = "Open Space"
End If

T(NX1, NY1).StructureID = 100
T(NX1, NY1).EarthTile = Rnd * 8
T(NX1, NY1).LandValue = 100
T(NX1, NY1).Population = 0               'Reset tile Vars
T(NX1, NY1).Growth = 0
T(NX1, NY1).ColorFlag = 0
T(NX1, NY1).Name = "Open Space"
End If

DrawBoard
End Sub

Private Sub pbSCREEN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
MouseOUT = False
CX = X: CY = Y
NX = Int(X / 13)
NY = Int(Y / 13)

Select Case T(NX, NY).StructureID
Case Is < 2: lblZone = "Residential"
Case Is < 6: lblZone = "Commercial"
Case Is < 10: lblZone = "Industrial"
Case Is < 20: lblZone = "Major Road"
Case Is < 24: lblZone = "Residential"
Case Is < 27: lblZone = "Commercial"
Case Is < 30: lblZone = "Industrial"
Case Is < 33: lblZone = "Safety"
Case 100: lblZone = "Free"
Case 200: lblZone = "Structure"
Case Is > 32: lblZone = "FreeBe"
End Select
lblZNum = Trim(T(NX, NY).Name)
lblGrowth = T(NX, NY).Growth
lblSPOP = T(NX, NY).Population
lblLV = "$" & T(NX, NY).LandValue

If MS.selectedPurchase > 20 And MS.selectedPurchase <= 32 Then
CURS = 2
If Cash < MS.price Then CURC = 1: Exit Sub
If Not T(NX, NY).StructureID = 100 Then CURC = 1: Exit Sub
If Not T(NX + 1, NY).StructureID = 100 Then CURC = 1: Exit Sub
If Not T(NX, NY + 1).StructureID = 100 Then CURC = 1: Exit Sub
If Not T(NX + 1, NY + 1).StructureID = 100 Then CURC = 1: Exit Sub
CURC = 2
Else
CURS = 1
If Cash < MS.price Then CURC = 1: Exit Sub
If Not T(NX, NY).StructureID = 100 Then CURC = 1: Exit Sub
CURC = 2
End If

If Button = 1 Then pbSCREEN_MouseDown 1, Shift, X, Y
End Sub
Private Sub pbSCREEN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CURL = 2
DrawBoard
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub Selector_Click(Index As Integer)
For iiii = 0 To 36
Selector(iiii).BorderStyle = 0
Next
 
MS.selectedPurchase = Index
Selector(Index).BorderStyle = 1

If Index <= 9 Then MS.price = 90000
If Index > 9 And Index <= 20 Then MS.price = 2000
If Index > 20 And Index <= 32 Then MS.price = 300000
If Index > 32 Then MS.price = 0
End Sub

Private Sub Selector_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOUT = True
End Sub

Private Sub TILEDATAUPDATE_Timer(): Dim TEMPIC As Long, PopC As Integer, RoadC As Integer
TotalPOP = 0: TEMPIC = 0: SafetyCount = 0: PopC = 0

For iii = 0 To MAPROWS 'Get Local Counts
For iiii = 0 To MAPCOLS
If T(iii, iiii).StructureID > 9 And T(iii, iiii).StructureID <= 20 Then RoadC = RoadC + 1
If T(iii, iiii).StructureID = 31 Or T(iii, iiii).StructureID = 32 Then SafetyCount = SafetyCount + 1
If T(iii, iiii).StructureID < 31 And T(iii, iiii).StructureID < 10 And T(iii, iiii).StructureID > 20 Then PopC = PopC + 1
Next
Next

Crime = (PopC / 4) - (2 * SafetyCount)

For iii = 0 To MAPROWS 'Update Tile Data/Stats
For iiii = 0 To MAPCOLS
If Not T(iii, iiii).StructureID = 100 Then
TotalPOP = TotalPOP + T(iii, iiii).Population

'Evade Filling Full Houses...
If T(iii, iiii).StructureID = 31 Or T(iii, iiii).StructureID = 32 Or T(iii, iiii).StructureID = 200 Then GoTo skipGROW
If T(iii, iiii).StructureID <= 1 And T(iii, iiii).Population >= 5 Then GoTo skipGROW
If T(iii, iiii).StructureID > 1 And T(iii, iiii).StructureID <= 5 And T(iii, iiii).Population >= 10 Then GoTo skipGROW
If T(iii, iiii).StructureID > 5 And T(iii, iiii).StructureID <= 9 And T(iii, iiii).Population >= 20 Then GoTo skipGROW
If T(iii, iiii).StructureID > 9 And T(iii, iiii).StructureID <= 20 Then GoTo skipGROW
If T(iii, iiii).StructureID > 20 And T(iii, iiii).StructureID <= 24 And T(iii, iiii).Population >= 400 Then GoTo skipGROW
If T(iii, iiii).StructureID > 24 And T(iii, iiii).StructureID <= 30 And T(iii, iiii).Population >= 300 Then GoTo skipGROW
If T(iii, iiii).StructureID > 32 Then GoTo skipALL

T(iii, iiii).Population = T(iii, iiii).Population + Rnd * (T(iii, iiii).Growth)
If T(iii, iiii).Population <= 0 Then T(iii, iiii).Population = 0
skipGROW:

T(iii, iiii).LandValue = T(iii, iiii).LandValue + 1 + (0.5 * T(iii, iiii).Growth)

skipALL:
If T(iii, iiii).StructureID <= 1 Then T(iii, iiii).Growth = T(iii, iiii).Growth + (30 - Crime)
If T(iii, iiii).StructureID > 20 And T(iii, iiii).StructureID <= 30 Then T(iii, iiii).Growth = T(iii, iiii).Growth + (50 - Crime)

If T(iii, iiii).Growth > 100 Then T(iii, iiii).Growth = 100
If T(iii, iiii).Growth < -100 Then T(iii, iiii).Growth = -100
End If
Next
Next

BitBlt BGPB.hDC, 0, 0, W, H, GFX.Turf(CurrentSeason).hDC, 0, 0, SRCCOPY
lblPOP = TotalPOP
CurMonth = CurMonth + 1
If CurMonth = 13 Then
CurMonth = 1

DrawBoard
CurYear = CurYear + 1
'1% income tax, incomes WILL vary
TEMPIC = TEMPIC + (0.01 * (TotalPOP * RndRange(1250, 1750)))
'1% sales tax, incomes WILL vary
TEMPIC = TEMPIC + (0.01 * (TotalPOP * RndRange(100, 500)))
'DIRTY MONEY (face the truth)
TEMPIC = TEMPIC + RndRange(10, 1000)
TEMPIC = TEMPIC - ((RoadC * 5) * 12)
TEMPIC = TEMPIC - ((SafetyCount * 100) * 12)
lblINCOME2 = lblINCOME
lblINCOME = TEMPIC
Cash = Cash + TEMPIC
filesave
Else
ReturnMstr CurMonth
'DrawBoardM
DrawBoard
End If

End Sub
Public Function RandResidence() As String
RNDS = Rnd * 20
Select Case RNDS
Case 1: RandResidence = "Smith"
Case 2: RandResidence = "Doe"
Case 3: RandResidence = "Brown"
Case 4: RandResidence = "Jenkins"
Case 5: RandResidence = "Wills"
Case 6: RandResidence = "Walls"
Case 7: RandResidence = "O'tool"
Case 8: RandResidence = "Xou"
Case 9: RandResidence = "Chee"
Case 10: RandResidence = "Xao"
Case 11: RandResidence = "Davis"
Case 12: RandResidence = "Connor"
Case 13: RandResidence = "Kramer"
Case 14: RandResidence = "Simpson"
Case 15: RandResidence = "Simons"
Case 16: RandResidence = "Marsh"
Case 17: RandResidence = "Jabip"
Case 18: RandResidence = "Benis"
Case 19: RandResidence = "Miller"
Case Else: RandResidence = "Lundy"
End Select
End Function
Public Function RandSmallCom() As String
RNDS = Rnd * 14
Select Case RNDS
Case 1: RandSmallCom = "Fixit Shop"
Case 2: RandSmallCom = "Pawn Shop"
Case 3: RandSmallCom = "Barber Shop"
Case 4: RandSmallCom = "GoodWill"
Case 5: RandSmallCom = "Kwick-E-Mart"
Case 6: RandSmallCom = "Gun Shop"
Case 7: RandSmallCom = "Skate Shop"
Case 8: RandSmallCom = "Dank Bar"
Case 9: RandSmallCom = "Classy Bar"
Case 10: RandSmallCom = "Singles Bar"
Case 11: RandSmallCom = "Karate Class"
Case 12: RandSmallCom = "Local Bank"
Case 13: RandSmallCom = "Big Corp. Bank"
Case Else: RandSmallCom = "Hardware Store"

End Select
End Function
Public Function RandSmallInd() As String
RNDS = Rnd * 12
Select Case RNDS
Case 1: RandSmallInd = "Semi-Inc"
Case 2: RandSmallInd = "Chems-o-Keembo"
Case 3: RandSmallInd = "UAW Meeting Hall"
Case 4: RandSmallInd = "Local Warez"
Case 5: RandSmallInd = "Local Factory"
Case 6: RandSmallInd = "Pump House"
Case 7: RandSmallInd = "Robot Central"
Case 8: RandSmallInd = "Simsense Corp."
Case 9: RandSmallInd = "Small Refinery"
Case 10: RandSmallInd = "Stadech Tech"
Case 11: RandSmallInd = "Locaar Labs"
Case Else: RandSmallInd = "Small Arms Inc."

End Select
End Function
Public Function RandBigRes() As String
RNDS = Rnd * 9
Select Case RNDS
Case 1: RandBigRes = "Overlook Appartments"
Case 2: RandBigRes = "El Cheap-o' Condo'"
Case 3: RandBigRes = "Loke's Den Condos"
Case 4: RandBigRes = "Pent's House"
Case 5: RandBigRes = "Local Appartments"
Case 6: RandBigRes = "Brook's Delta Inn"
Case 7: RandBigRes = "Sleaze Central"
Case 8: RandBigRes = "Del Boka Vista."
Case Else: RandBigRes = "Phase 5"

End Select
End Function
Public Function RandBigCom() As String
RNDS = Rnd * 9
Select Case RNDS
Case 1: RandBigCom = "CrazyLegs.com"
Case 2: RandBigCom = "Reynold's Appraisal"
Case 3: RandBigCom = "Golden Real Estate"
Case 4: RandBigCom = "FreelanceMerc.com"
Case 5: RandBigCom = "ishoponline.com"
Case 6: RandBigCom = "WebServ.com"
Case 7: RandBigCom = "Mecco Reliance"
Case 8: RandBigCom = "Large Law Firm"
Case Else: RandBigCom = "Phase5.com"

End Select
End Function
Public Function RandBigInd() As String
RNDS = Rnd * 9
Select Case RNDS
Case 1: RandBigInd = "WebServe Tech"
Case 2: RandBigInd = "Mecco Weaps. Plant"
Case 3: RandBigInd = "Auto Manufacturer"
Case 4: RandBigInd = "Large Refinery"
Case 5: RandBigInd = "Can Manufacturer"
Case 6: RandBigInd = "ORE Mine"
Case 7: RandBigInd = "Mecco Tech"
Case 8: RandBigInd = "Mecco S-conductors"
Case Else: RandBigInd = "Mecco R-Branch"

End Select
End Function

Sub DrawSelectors()
For iii = 0 To 36
BitBlt Selector(iii).hDC, 0, 0, 13, 13, TURFFall(1).hDC, 0, 0, SRCCOPY
BitBlt Selector(iii).hDC, 13, 0, 13, 13, TURFFall(2).hDC, 0, 0, SRCCOPY
BitBlt Selector(iii).hDC, 0, 13, 13, 13, TURFFall(3).hDC, 0, 0, SRCCOPY
BitBlt Selector(iii).hDC, 13, 13, 13, 13, TURFFall(4).hDC, 0, 0, SRCCOPY
Next

BitBlt Selector(0).hDC, 0, 0, 13, 13, GFX.h1M.hDC, 0, 0, SRCAND
BitBlt Selector(0).hDC, 0, 0, 13, 13, GFX.h1SMon.hDC, 0, 0, SRCPAINT

BitBlt Selector(1).hDC, 0, 0, 13, 13, GFX.Picture1.hDC, 0, 0, SRCAND
BitBlt Selector(1).hDC, 0, 0, 13, 13, GFX.h2SMon.hDC, 0, 0, SRCPAINT

BitBlt Selector(2).hDC, 0, 0, 13, 13, GFX.c1M.hDC, 0, 0, SRCAND
BitBlt Selector(2).hDC, 0, 0, 13, 13, GFX.c1S.hDC, 0, 0, SRCPAINT

BitBlt Selector(3).hDC, 0, 0, 13, 13, GFX.c2M.hDC, 0, 0, SRCAND
BitBlt Selector(3).hDC, 0, 0, 13, 13, GFX.c2S.hDC, 0, 0, SRCPAINT

BitBlt Selector(4).hDC, 0, 0, 13, 13, GFX.c3M.hDC, 0, 0, SRCAND
BitBlt Selector(4).hDC, 0, 0, 13, 13, GFX.c3S.hDC, 0, 0, SRCPAINT

BitBlt Selector(5).hDC, 0, 0, 13, 13, GFX.c4M.hDC, 0, 0, SRCAND
BitBlt Selector(5).hDC, 0, 0, 13, 13, GFX.c4S.hDC, 0, 0, SRCPAINT

BitBlt Selector(6).hDC, 0, 0, 13, 13, GFX.i1M.hDC, 0, 0, SRCAND
BitBlt Selector(6).hDC, 0, 0, 13, 13, GFX.i1S.hDC, 0, 0, SRCPAINT

BitBlt Selector(7).hDC, 0, 0, 13, 13, GFX.i2M.hDC, 0, 0, SRCAND
BitBlt Selector(7).hDC, 0, 0, 13, 13, GFX.i2S.hDC, 0, 0, SRCPAINT

BitBlt Selector(8).hDC, 0, 0, 13, 13, GFX.i3M.hDC, 0, 0, SRCAND
BitBlt Selector(8).hDC, 0, 0, 13, 13, GFX.i3S.hDC, 0, 0, SRCPAINT

BitBlt Selector(9).hDC, 0, 0, 13, 13, GFX.i4M.hDC, 0, 0, SRCAND
BitBlt Selector(9).hDC, 0, 0, 13, 13, GFX.i4S.hDC, 0, 0, SRCPAINT

BitBlt Selector(10).hDC, 3, 3, 13, 13, GFX.rT4M.hDC, 0, 0, SRCAND
BitBlt Selector(10).hDC, 3, 3, 13, 13, GFX.rT4s.hDC, 0, 0, SRCPAINT

BitBlt Selector(11).hDC, 0, 3, 13, 13, GFX.rT1M.hDC, 0, 0, SRCAND
BitBlt Selector(11).hDC, 0, 3, 13, 13, GFX.rT1s.hDC, 0, 0, SRCPAINT

BitBlt Selector(12).hDC, 3, 3, 13, 13, GFX.rC3M.hDC, 0, 0, SRCAND
BitBlt Selector(12).hDC, 3, 3, 13, 13, GFX.rC3s.hDC, 0, 0, SRCPAINT

BitBlt Selector(13).hDC, 0, 3, 13, 13, GFX.rC4M.hDC, 0, 0, SRCAND
BitBlt Selector(13).hDC, 0, 3, 13, 13, GFX.rC4s.hDC, 0, 0, SRCPAINT

BitBlt Selector(14).hDC, 2, 2, 13, 13, GFX.rLRM.hDC, 0, 0, SRCAND
BitBlt Selector(14).hDC, 2, 2, 13, 13, GFX.rLRs.hDC, 0, 0, SRCPAINT

BitBlt Selector(15).hDC, 3, 0, 13, 13, GFX.rT2M.hDC, 0, 0, SRCAND
BitBlt Selector(15).hDC, 3, 0, 13, 13, GFX.rT2s.hDC, 0, 0, SRCPAINT

BitBlt Selector(16).hDC, 0, 0, 13, 13, GFX.rT3M.hDC, 0, 0, SRCAND
BitBlt Selector(16).hDC, 0, 0, 13, 13, GFX.rT3s.hDC, 0, 0, SRCPAINT

BitBlt Selector(17).hDC, 3, 0, 13, 13, GFX.rC1M.hDC, 0, 0, SRCAND
BitBlt Selector(17).hDC, 3, 0, 13, 13, GFX.rC1s.hDC, 0, 0, SRCPAINT

BitBlt Selector(18).hDC, 0, 0, 13, 13, GFX.rC2M.hDC, 0, 0, SRCAND
BitBlt Selector(18).hDC, 0, 0, 13, 13, GFX.rC2s.hDC, 0, 0, SRCPAINT

BitBlt Selector(19).hDC, 2, 2, 13, 13, GFX.rUDM.hDC, 0, 0, SRCAND
BitBlt Selector(19).hDC, 2, 2, 13, 13, GFX.rUDs.hDC, 0, 0, SRCPAINT

BitBlt Selector(20).hDC, 0, 0, 26, 26, GFX.RoadIm.hDC, 0, 0, SRCAND
BitBlt Selector(20).hDC, 0, 0, 26, 26, GFX.RoadIs.hDC, 0, 0, SRCPAINT

BitBlt Selector(21).hDC, 0, 0, 26, 26, GFX.BR1M.hDC, 0, 0, SRCAND
BitBlt Selector(21).hDC, 0, 0, 26, 26, GFX.BR1s.hDC, 0, 0, SRCPAINT

BitBlt Selector(22).hDC, 0, 0, 26, 26, GFX.BR2M.hDC, 0, 0, SRCAND
BitBlt Selector(22).hDC, 0, 0, 26, 26, GFX.BR2s.hDC, 0, 0, SRCPAINT

BitBlt Selector(23).hDC, 0, 0, 26, 26, GFX.BR3M.hDC, 0, 0, SRCAND
BitBlt Selector(23).hDC, 0, 0, 26, 26, GFX.BR3s.hDC, 0, 0, SRCPAINT

BitBlt Selector(24).hDC, 0, 0, 26, 26, GFX.BR4M.hDC, 0, 0, SRCAND
BitBlt Selector(24).hDC, 0, 0, 26, 26, GFX.BR4s.hDC, 0, 0, SRCPAINT

BitBlt Selector(25).hDC, 0, 0, 26, 26, GFX.BC1M.hDC, 0, 0, SRCAND
BitBlt Selector(25).hDC, 0, 0, 26, 26, GFX.BC1s.hDC, 0, 0, SRCPAINT

BitBlt Selector(26).hDC, 0, 0, 26, 26, GFX.BC2M.hDC, 0, 0, SRCAND
BitBlt Selector(26).hDC, 0, 0, 26, 26, GFX.BC2s.hDC, 0, 0, SRCPAINT

BitBlt Selector(27).hDC, 0, 0, 26, 26, GFX.BC3M.hDC, 0, 0, SRCAND
BitBlt Selector(27).hDC, 0, 0, 26, 26, GFX.BC3s.hDC, 0, 0, SRCPAINT

BitBlt Selector(28).hDC, 0, 0, 26, 26, GFX.BI1M.hDC, 0, 0, SRCAND
BitBlt Selector(28).hDC, 0, 0, 26, 26, GFX.BI1s.hDC, 0, 0, SRCPAINT

BitBlt Selector(29).hDC, 0, 0, 26, 26, GFX.BI2M.hDC, 0, 0, SRCAND
BitBlt Selector(29).hDC, 0, 0, 26, 26, GFX.BI2s.hDC, 0, 0, SRCPAINT

BitBlt Selector(30).hDC, 0, 0, 26, 26, GFX.BI3M.hDC, 0, 0, SRCAND
BitBlt Selector(30).hDC, 0, 0, 26, 26, GFX.BI3s.hDC, 0, 0, SRCPAINT

BitBlt Selector(31).hDC, 0, 0, 26, 26, GFX.ESM.hDC, 0, 0, SRCAND
BitBlt Selector(31).hDC, 0, 0, 26, 26, GFX.ES1.hDC, 0, 0, SRCPAINT

BitBlt Selector(32).hDC, 0, 0, 26, 26, GFX.ESM.hDC, 0, 0, SRCAND
BitBlt Selector(32).hDC, 0, 0, 26, 26, GFX.ES2.hDC, 0, 0, SRCPAINT

BitBlt Selector(33).hDC, 0, 0, 13, 13, GFX.Tree1M.hDC, 0, 0, SRCAND
BitBlt Selector(33).hDC, 0, 0, 13, 13, GFX.Tree1sSS.hDC, 0, 0, SRCPAINT

BitBlt Selector(34).hDC, 0, 0, 13, 13, GFX.Tree2M.hDC, 0, 0, SRCAND
BitBlt Selector(34).hDC, 0, 0, 13, 13, GFX.Tree2s.hDC, 0, 0, SRCPAINT

BitBlt Selector(35).hDC, 0, 0, 13, 13, GFX.BusM.hDC, 0, 0, SRCAND
BitBlt Selector(35).hDC, 0, 0, 13, 13, GFX.BusS.hDC, 0, 0, SRCPAINT

BitBlt Selector(36).hDC, i * 13, ii * 13, 13, 13, GFX.GC1(5).hDC, 0, 0, SRCAND
BitBlt Selector(36).hDC, 0, 0, 13, 13, GFX.GC1(0).hDC, 0, 0, SRCPAINT

End Sub

