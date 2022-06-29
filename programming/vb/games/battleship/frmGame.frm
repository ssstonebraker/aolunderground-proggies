VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battleship v. 1.0"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Levels"
      Height          =   2175
      Left            =   5280
      TabIndex        =   208
      Top             =   600
      Width           =   2055
      Begin VB.OptionButton optLevel 
         Caption         =   "Admiral"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   215
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Rear admiral"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   214
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Captain"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   213
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Commander"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   212
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Lieutenant"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   211
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Ensign"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   210
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Mate"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   209
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Starter"
      Height          =   975
      Left            =   5280
      TabIndex        =   205
      Top             =   2880
      Width           =   2055
      Begin VB.OptionButton optStart 
         Caption         =   "Computer"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   207
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optStart 
         Caption         =   "You"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   206
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   204
      Top             =   3480
      Width           =   5055
   End
   Begin VB.CommandButton Resign 
      Caption         =   "Resign"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   203
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton NewGame 
      Caption         =   "New game"
      Height          =   375
      Left            =   120
      TabIndex        =   200
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Your ships"
      Height          =   255
      Left            =   2760
      TabIndex        =   202
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Computers ships"
      Height          =   255
      Left            =   120
      TabIndex        =   201
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   99
      Left            =   2280
      TabIndex        =   199
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   98
      Left            =   2040
      TabIndex        =   198
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   97
      Left            =   1800
      TabIndex        =   197
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   96
      Left            =   1560
      TabIndex        =   196
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   95
      Left            =   1320
      TabIndex        =   195
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   94
      Left            =   1080
      TabIndex        =   194
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   93
      Left            =   840
      TabIndex        =   193
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   92
      Left            =   600
      TabIndex        =   192
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   91
      Left            =   360
      TabIndex        =   191
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   90
      Left            =   120
      TabIndex        =   190
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   89
      Left            =   2280
      TabIndex        =   189
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   88
      Left            =   2040
      TabIndex        =   188
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   87
      Left            =   1800
      TabIndex        =   187
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   86
      Left            =   1560
      TabIndex        =   186
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   85
      Left            =   1320
      TabIndex        =   185
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   84
      Left            =   1080
      TabIndex        =   184
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   83
      Left            =   840
      TabIndex        =   183
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   82
      Left            =   600
      TabIndex        =   182
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   81
      Left            =   360
      TabIndex        =   181
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   80
      Left            =   120
      TabIndex        =   180
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   79
      Left            =   2280
      TabIndex        =   179
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   78
      Left            =   2040
      TabIndex        =   178
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   77
      Left            =   1800
      TabIndex        =   177
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   76
      Left            =   1560
      TabIndex        =   176
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   75
      Left            =   1320
      TabIndex        =   175
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   74
      Left            =   1080
      TabIndex        =   174
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   73
      Left            =   840
      TabIndex        =   173
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   72
      Left            =   600
      TabIndex        =   172
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   71
      Left            =   360
      TabIndex        =   171
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   70
      Left            =   120
      TabIndex        =   170
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   69
      Left            =   2280
      TabIndex        =   169
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   68
      Left            =   2040
      TabIndex        =   168
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   67
      Left            =   1800
      TabIndex        =   167
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   66
      Left            =   1560
      TabIndex        =   166
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   65
      Left            =   1320
      TabIndex        =   165
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   64
      Left            =   1080
      TabIndex        =   164
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   63
      Left            =   840
      TabIndex        =   163
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   62
      Left            =   600
      TabIndex        =   162
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   61
      Left            =   360
      TabIndex        =   161
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   60
      Left            =   120
      TabIndex        =   160
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   59
      Left            =   2280
      TabIndex        =   159
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   58
      Left            =   2040
      TabIndex        =   158
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   57
      Left            =   1800
      TabIndex        =   157
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   56
      Left            =   1560
      TabIndex        =   156
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   55
      Left            =   1320
      TabIndex        =   155
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   54
      Left            =   1080
      TabIndex        =   154
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   53
      Left            =   840
      TabIndex        =   153
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   52
      Left            =   600
      TabIndex        =   152
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   51
      Left            =   360
      TabIndex        =   151
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   50
      Left            =   120
      TabIndex        =   150
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   49
      Left            =   2280
      TabIndex        =   149
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   48
      Left            =   2040
      TabIndex        =   148
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   1800
      TabIndex        =   147
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   1560
      TabIndex        =   146
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   1320
      TabIndex        =   145
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   1080
      TabIndex        =   144
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   840
      TabIndex        =   143
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   600
      TabIndex        =   142
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   360
      TabIndex        =   141
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   120
      TabIndex        =   140
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   2280
      TabIndex        =   139
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   2040
      TabIndex        =   138
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   1800
      TabIndex        =   137
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   1560
      TabIndex        =   136
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   1320
      TabIndex        =   135
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   1080
      TabIndex        =   134
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   840
      TabIndex        =   133
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   600
      TabIndex        =   132
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   360
      TabIndex        =   131
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   120
      TabIndex        =   130
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   2280
      TabIndex        =   129
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   128
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   1800
      TabIndex        =   127
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   1560
      TabIndex        =   126
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   1320
      TabIndex        =   125
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   1080
      TabIndex        =   124
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   840
      TabIndex        =   123
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   600
      TabIndex        =   122
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   360
      TabIndex        =   121
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   120
      Top             =   960
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   2280
      TabIndex        =   119
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   2040
      TabIndex        =   118
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   1800
      TabIndex        =   117
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   1560
      TabIndex        =   116
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   1320
      TabIndex        =   115
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   1080
      TabIndex        =   114
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   840
      TabIndex        =   113
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   600
      TabIndex        =   112
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   111
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   110
      Top             =   720
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   109
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   108
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   107
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   106
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   105
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   104
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   103
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   102
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   101
      Top             =   480
      Width           =   255
   End
   Begin VB.Label computer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmGame.frx":030A
      TabIndex        =   100
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   99
      Left            =   4920
      TabIndex        =   99
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   98
      Left            =   4680
      TabIndex        =   98
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   97
      Left            =   4440
      TabIndex        =   97
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   96
      Left            =   4200
      TabIndex        =   96
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   95
      Left            =   3960
      TabIndex        =   95
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   94
      Left            =   3720
      TabIndex        =   94
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   93
      Left            =   3480
      TabIndex        =   93
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   92
      Left            =   3240
      TabIndex        =   92
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   91
      Left            =   3000
      TabIndex        =   91
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   90
      Left            =   2760
      TabIndex        =   90
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   89
      Left            =   4920
      TabIndex        =   89
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   88
      Left            =   4680
      TabIndex        =   88
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   87
      Left            =   4440
      TabIndex        =   87
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   86
      Left            =   4200
      TabIndex        =   86
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   85
      Left            =   3960
      TabIndex        =   85
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   84
      Left            =   3720
      TabIndex        =   84
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   83
      Left            =   3480
      TabIndex        =   83
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   82
      Left            =   3240
      TabIndex        =   82
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   81
      Left            =   3000
      TabIndex        =   81
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   80
      Left            =   2760
      TabIndex        =   80
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   79
      Left            =   4920
      TabIndex        =   79
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   78
      Left            =   4680
      TabIndex        =   78
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   77
      Left            =   4440
      TabIndex        =   77
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   76
      Left            =   4200
      TabIndex        =   76
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   75
      Left            =   3960
      TabIndex        =   75
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   74
      Left            =   3720
      TabIndex        =   74
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   73
      Left            =   3480
      TabIndex        =   73
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   72
      Left            =   3240
      TabIndex        =   72
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   71
      Left            =   3000
      TabIndex        =   71
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   70
      Left            =   2760
      TabIndex        =   70
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   69
      Left            =   4920
      TabIndex        =   69
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   68
      Left            =   4680
      TabIndex        =   68
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   67
      Left            =   4440
      TabIndex        =   67
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   66
      Left            =   4200
      TabIndex        =   66
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   65
      Left            =   3960
      TabIndex        =   65
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   64
      Left            =   3720
      TabIndex        =   64
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   63
      Left            =   3480
      TabIndex        =   63
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   62
      Left            =   3240
      TabIndex        =   62
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   61
      Left            =   3000
      TabIndex        =   61
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   60
      Left            =   2760
      TabIndex        =   60
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   59
      Left            =   4920
      TabIndex        =   59
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   58
      Left            =   4680
      TabIndex        =   58
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   57
      Left            =   4440
      TabIndex        =   57
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   56
      Left            =   4200
      TabIndex        =   56
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   55
      Left            =   3960
      TabIndex        =   55
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   54
      Left            =   3720
      TabIndex        =   54
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   53
      Left            =   3480
      TabIndex        =   53
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   52
      Left            =   3240
      TabIndex        =   52
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   51
      Left            =   3000
      TabIndex        =   51
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   50
      Left            =   2760
      TabIndex        =   50
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   49
      Left            =   4920
      TabIndex        =   49
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   48
      Left            =   4680
      TabIndex        =   48
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   4440
      TabIndex        =   47
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   4200
      TabIndex        =   46
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   3960
      TabIndex        =   45
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   3720
      TabIndex        =   44
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   3480
      TabIndex        =   43
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   3240
      TabIndex        =   42
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   3000
      TabIndex        =   41
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   2760
      TabIndex        =   40
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   4920
      TabIndex        =   39
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   4680
      TabIndex        =   38
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   4440
      TabIndex        =   37
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   4200
      TabIndex        =   36
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   3960
      TabIndex        =   35
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   3720
      TabIndex        =   34
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   3480
      TabIndex        =   33
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   3240
      TabIndex        =   32
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   3000
      TabIndex        =   31
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   2760
      TabIndex        =   30
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   4920
      TabIndex        =   29
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   4680
      TabIndex        =   28
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   4440
      TabIndex        =   27
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   4200
      TabIndex        =   26
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   3960
      TabIndex        =   25
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   3720
      TabIndex        =   24
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   3480
      TabIndex        =   23
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   3240
      TabIndex        =   22
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   3000
      TabIndex        =   21
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   2760
      TabIndex        =   20
      Top             =   960
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   4920
      TabIndex        =   19
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   4680
      TabIndex        =   18
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   4440
      TabIndex        =   17
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   4200
      TabIndex        =   16
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   3960
      TabIndex        =   15
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   3720
      TabIndex        =   14
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   3480
      TabIndex        =   13
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   12
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   11
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.Label human 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim my_game As New game
Dim iLevel As Integer
Dim iStarter As Integer
Dim running As Boolean
Const HUMAN_MISS = 0
Const HUMAN_HIT = 1
Const HUMAN_WINNER = 2
Const HUMAN_TAKEN = 3
Const HUMAN_WRONG = 4

Const COMPUTER_MISS = 0
Const COMPUTER_HIT = 1
Const COMPUTER_WINNER = 2
Const COMPUTER_SUNK = 3

Private Sub Command1_Click()
    'show the help-dialog
    'the 1 at the end ensures a modal dialog. you must end it
    'before you can continue
    frmHelp.Show 1
End Sub

Private Sub NewGame_Click()
    Dim i As Integer
    Dim j As Integer
    Randomize
    my_game.init_board iLevel
    'populating the left board with water (that is we don't know which is which)
    For i = 0 To 9
        For j = 0 To 9
            computer(i * 10 + j).BackColor = vbBlue
        Next
    Next
    
    'populating the right board with the humans ships
    For i = 0 To 9
        For j = 0 To 9
            If my_game.get_player(1, i + 1, j + 1) > 11 Then
                human(i * 10 + j).BackColor = vbGreen
            Else
                human(i * 10 + j).BackColor = vbBlue
            End If
        Next
    Next
    
    'disabling level-options
    For i = 0 To 6
        optLevel(i).Enabled = False
    Next
    optStart(0).Enabled = False
    optStart(1).Enabled = False
    'updating certain variables
    'the game is running
    running = True
    
    'the two command buttons for gameplay
    Resign.Enabled = True
    NewGame.Enabled = False
    
    If iStarter = 1 Then DoComputer
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    'init a game based upon mate-level
    Randomize
    
    'set both grids to blue color
    For i = 0 To 9
        For j = 0 To 9
            computer(i * 10 + j).BackColor = vbBlue
        Next
    Next
    For i = 0 To 9
        For j = 0 To 9
            human(i * 10 + j).BackColor = vbBlue
        Next
    Next
    
    'not a running game
    running = False
End Sub

Private Sub computer_Click(Index As Integer)
    If running Then
        'do a new turn based upon the cell clicked
        Select Case my_game.human_turn(Index \ 10 + 1, Index Mod 10 + 1)
            Case HUMAN_MISS
                computer(Index).BackColor = vbYellow
            Case HUMAN_HIT
                computer(Index).BackColor = vbRed
                Caption = "Score is " + CStr(my_game.get_hit(0)) + "-" + CStr(my_game.get_hit(1))
            Case HUMAN_WINNER
                computer(Index).BackColor = vbRed
                Caption = "Once in a lifetime punk!"
                close_grid
            Case HUMAN_TAKEN To HUMAN_WRONG
                Beep
                MsgBox "Can" + Chr(39) + "t you shoot straight!", vbExclamation, "Moron!!!"
                Exit Sub
        End Select
        
        'thereafter it's the computers turn
        DoComputer
    End If
End Sub
Private Sub DoComputer()
    Dim nRow As Integer
    Dim nCol As Integer
    
    'since we're doing a ByRef, the nRow and nCol contains the appropriate destination
    Select Case my_game.computer_turn(nRow, nCol)
        Case COMPUTER_MISS
            human(nRow * 10 + nCol - 11).BackColor = vbYellow
        Case COMPUTER_HIT
            human(nRow * 10 + nCol - 11).BackColor = vbRed
            Caption = "Score is " + CStr(my_game.get_hit(0)) + "-" + CStr(my_game.get_hit(1))
        Case COMPUTER_WINNER
            human(nRow * 10 + nCol - 11).BackColor = vbRed
            Caption = "I am so good it hurts!"
            close_grid
        Case COMPUTER_SUNK
            human(nRow * 10 + nCol - 11).BackColor = vbRed
            Caption = "Score is " + CStr(my_game.get_hit(0)) + "-" + CStr(my_game.get_hit(1))
            MsgBox "Another big ship is a goner"
    End Select
End Sub

Private Sub optLevel_Click(Index As Integer)
    iLevel = Index
End Sub

Public Sub close_grid()
    Dim i As Integer
    Dim j As Integer
    
    'reveal grid to the player
    For i = 0 To 9
        For j = 0 To 9
            If my_game.get_player(0, i + 1, j + 1) > 11 And my_game.get_player(0, i + 1, j + 1) < 16 Then
                computer(i * 10 + j).BackColor = vbGreen
            End If
        Next
    Next
    
    'en/dis-able items on the board
    For i = 0 To 6
        optLevel(i).Enabled = True
    Next

    optStart(0).Enabled = True
    optStart(1).Enabled = True
    
    Resign.Enabled = False
    NewGame.Enabled = True
    
    'disable a running game
    running = False
End Sub

Private Sub optStart_Click(Index As Integer)
    iStarter = Index
End Sub

Private Sub Resign_Click()
    close_grid
    Caption = "Chicken-shit!!!"
End Sub
