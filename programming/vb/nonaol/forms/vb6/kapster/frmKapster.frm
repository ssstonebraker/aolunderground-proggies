VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKapster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kapster Client - Beta .2"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKapster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   192
      Text            =   "frmKapster.frx":0CCA
      Top             =   4920
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4080
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock99 
      Left            =   4560
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6700
   End
   Begin VB.Frame Frame2 
      Caption         =   "Instant Messaging"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   8
      Left            =   120
      TabIndex        =   175
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Timer Timer6 
         Interval        =   1
         Left            =   6480
         Top             =   3600
      End
      Begin VB.CommandButton Command51 
         Caption         =   "-"
         Height          =   255
         Left            =   6240
         TabIndex        =   184
         Top             =   3360
         Width           =   375
      End
      Begin VB.CommandButton Command50 
         Caption         =   "+"
         Height          =   255
         Left            =   5880
         TabIndex        =   183
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox Text44 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2655
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   179
         Top             =   600
         Width           =   5055
      End
      Begin VB.ListBox List9 
         Height          =   3120
         Left            =   5160
         TabIndex        =   178
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text45 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   177
         Top             =   3360
         Width           =   3735
      End
      Begin VB.CommandButton Command49 
         Caption         =   "Say it"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   176
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label49 
         Caption         =   "<null>"
         Height          =   255
         Left            =   1920
         TabIndex        =   185
         Top             =   275
         Width           =   2295
      End
      Begin VB.Label Label48 
         Caption         =   "0"
         Height          =   255
         Left            =   4920
         TabIndex        =   182
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "0 Users"
         Height          =   255
         Left            =   5280
         TabIndex        =   181
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   5160
         X2              =   5160
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Label Label21 
         Caption         =   "You are chatting with:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   180
         Top             =   275
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "More Dumb Stuff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   7
      Left            =   120
      TabIndex        =   151
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame Frame24 
         Caption         =   "Change Info"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   2640
         TabIndex        =   162
         Top             =   1680
         Width           =   4095
         Begin VB.CommandButton Command48 
            Caption         =   "Change!"
            Height          =   285
            Left            =   2880
            TabIndex        =   174
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox Text43 
            Height          =   270
            Left            =   1200
            TabIndex        =   173
            Text            =   "6699"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Change!"
            Height          =   285
            Left            =   2880
            TabIndex        =   171
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox Text42 
            Height          =   270
            Left            =   1200
            TabIndex        =   170
            Text            =   "anon@napster.com"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Command46 
            Caption         =   "Change!"
            Height          =   285
            Left            =   2880
            TabIndex        =   168
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox Text41 
            Height          =   270
            Left            =   1200
            TabIndex        =   167
            Text            =   "1"
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton Command45 
            Caption         =   "Change!"
            Height          =   285
            Left            =   2880
            TabIndex        =   165
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text40 
            Height          =   270
            Left            =   1200
            TabIndex        =   164
            Text            =   "newone"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label47 
            Caption         =   "Data Port #:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label46 
            Caption         =   "Email Addy:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   169
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label45 
            Caption         =   "Link Speed:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label44 
            Caption         =   "My Password:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "User Pinger/Ponger"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2640
         TabIndex        =   161
         Top             =   360
         Width           =   4095
         Begin VB.TextBox Text36 
            Height          =   270
            Left            =   600
            TabIndex        =   191
            Text            =   "100"
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Pong"
            Height          =   255
            Left            =   2640
            TabIndex        =   189
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Ping"
            Height          =   255
            Left            =   1560
            TabIndex        =   188
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text19 
            Height          =   270
            Left            =   960
            TabIndex        =   187
            Text            =   "wizistheman"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label13 
            Caption         =   "X:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   190
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   186
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Server Ping"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   158
         Top             =   3000
         Width           =   2415
         Begin VB.CommandButton Command43 
            Caption         =   "Ping it!"
            Height          =   285
            Left            =   120
            TabIndex        =   159
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label43 
            Caption         =   "idle"
            Height          =   255
            Left            =   1200
            TabIndex        =   160
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Channel Detailer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   152
         Top             =   240
         Width           =   2415
         Begin VB.ListBox List6 
            Height          =   1500
            Left            =   120
            TabIndex        =   157
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton Command42 
            Caption         =   "Detail"
            Height          =   255
            Left            =   1440
            TabIndex        =   154
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text38 
            Height          =   270
            Left            =   840
            TabIndex        =   153
            Text            =   "Blues"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label42 
            Caption         =   "Results:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label41 
            Caption         =   "Channel:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "About Kapster Client"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   9
      Left            =   120
      TabIndex        =   147
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame Frame20 
         Height          =   1935
         Left            =   840
         TabIndex        =   149
         Top             =   1680
         Width           =   5175
         Begin VB.TextBox Text17 
            ForeColor       =   &H00006060&
            Height          =   1575
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   150
            Text            =   "frmKapster.frx":0CD5
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00E0E0E0&
         X1              =   6000
         X2              =   6000
         Y1              =   1560
         Y2              =   360
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E0E0E0&
         X1              =   840
         X2              =   6000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         X1              =   6000
         X2              =   840
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   840
         X2              =   840
         Y1              =   1560
         Y2              =   360
      End
      Begin VB.Image Image3 
         Height          =   1200
         Left            =   840
         Picture         =   "frmKapster.frx":0EF0
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Download Files"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   6
      Left            =   120
      TabIndex        =   105
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame Frame19 
         Caption         =   "Download File"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3120
         TabIndex        =   129
         Top             =   2400
         Width           =   3615
         Begin VB.Timer Timer3 
            Interval        =   3000
            Left            =   3120
            Top             =   840
         End
         Begin VB.CommandButton Command40 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   136
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Download"
            Height          =   285
            Left            =   2400
            TabIndex        =   135
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command38 
            Caption         =   "..."
            Height          =   285
            Left            =   3000
            TabIndex        =   132
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text35 
            Height          =   270
            Left            =   840
            TabIndex        =   131
            Text            =   "C:\kapster.mp3"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00E0E0E0&
            X1              =   480
            X2              =   480
            Y1              =   600
            Y2              =   1200
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00E0E0E0&
            X1              =   120
            X2              =   480
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00404040&
            X1              =   480
            X2              =   120
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   120
            Y1              =   1200
            Y2              =   600
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            Caption         =   "0/0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   138
            Top             =   890
            Width           =   1695
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            Caption         =   "0k a second"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   137
            Top             =   650
            Width           =   1695
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   120
            TabIndex        =   133
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label36 
            Caption         =   "Save As:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Download Flood"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3120
         TabIndex        =   123
         Top             =   960
         Width           =   3615
         Begin VB.CommandButton Command17 
            Caption         =   "Attack'em"
            Height          =   255
            Left            =   2280
            TabIndex        =   127
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text34 
            Height          =   270
            Left            =   2280
            TabIndex        =   125
            Text            =   "wizistheman"
            Top             =   480
            Width           =   1215
         End
         Begin VB.ListBox List5 
            Height          =   960
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label35 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   126
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Query Results"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   118
         Top             =   240
         Width           =   3615
         Begin VB.TextBox Text33 
            Height          =   270
            Left            =   2760
            TabIndex        =   122
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text32 
            Height          =   270
            Left            =   480
            TabIndex        =   121
            Text            =   "0.0.0.0"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label34 
            Caption         =   "IP:                                         Port:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "IP/Port Query"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   113
         Top             =   1680
         Width           =   2895
         Begin VB.CommandButton Command20 
            Caption         =   "Query"
            Height          =   285
            Left            =   1800
            TabIndex        =   119
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox Text16 
            Height          =   270
            Left            =   120
            TabIndex        =   115
            Text            =   "C:\goodmusic1.mp3"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox Text18 
            Height          =   270
            Left            =   120
            TabIndex        =   114
            Text            =   "wizistheman"
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label18 
            Caption         =   "File String:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Upload/Download Stats"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   106
         Top             =   480
         Width           =   2895
         Begin VB.OptionButton Option4 
            Caption         =   "rem"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   112
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "add"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   600
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox Text15 
            Height          =   270
            Left            =   840
            TabIndex        =   109
            Text            =   "100"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command16 
            Caption         =   "'Download'"
            Height          =   285
            Left            =   1680
            TabIndex        =   108
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command18 
            Caption         =   "'Upload'"
            Height          =   285
            Left            =   1680
            TabIndex        =   107
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Amount:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search for Songs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   5
      Left            =   120
      TabIndex        =   91
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton Command36 
         Caption         =   "Copy Song"
         Height          =   255
         Left            =   5520
         TabIndex        =   103
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Copy Name"
         Height          =   255
         Left            =   4320
         TabIndex        =   101
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Move 2 Query"
         Height          =   255
         Left            =   2880
         TabIndex        =   100
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   5880
         TabIndex        =   99
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Search"
         Height          =   255
         Left            =   5040
         TabIndex        =   98
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text31 
         Height          =   270
         Left            =   3960
         TabIndex        =   97
         Text            =   "100"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text30 
         Height          =   270
         Left            =   720
         TabIndex        =   96
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text29 
         Height          =   270
         Left            =   720
         TabIndex        =   94
         Top             =   240
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   120
         TabIndex        =   92
         Top             =   960
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Song Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Song Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Length"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "User"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Speed"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label33 
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "Artist:                                                   Max Results:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "1 On 1 Stuff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   4
      Left            =   120
      TabIndex        =   77
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame Frame13 
         Caption         =   "User's Song List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   3120
         TabIndex        =   84
         Top             =   240
         Width           =   3615
         Begin VB.CommandButton Command19 
            Caption         =   "2 Flood"
            Height          =   255
            Left            =   2520
            TabIndex        =   128
            Top             =   3120
            Width           =   975
         End
         Begin VB.CommandButton Command37 
            Caption         =   "Copy Song"
            Height          =   255
            Left            =   1440
            TabIndex        =   104
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Move 2 Query"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox Text28 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   89
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Find"
            Height          =   285
            Left            =   2760
            TabIndex        =   87
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text27 
            Height          =   270
            Left            =   1080
            TabIndex        =   85
            Text            =   "wizistheman"
            Top             =   240
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   90
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Song Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Song Size"
               Object.Width           =   2275
            EndProperty
         End
         Begin VB.Label Label31 
            Caption         =   "User's ID:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   88
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label30 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Finger User"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   2895
         Begin VB.TextBox Text14 
            Height          =   270
            Left            =   1080
            TabIndex        =   81
            Text            =   "wizistheman"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Finger"
            Height          =   285
            Left            =   1920
            TabIndex        =   80
            Top             =   600
            Width           =   855
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   2040
            Left            =   120
            TabIndex        =   79
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label29 
            Caption         =   "Results:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "My Files and Stuff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   3
      Left            =   120
      TabIndex        =   49
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   71
         Top             =   480
         Width           =   3735
         Begin VB.CheckBox Check3 
            Caption         =   "Single Send"
            Height          =   255
            Left            =   1920
            TabIndex        =   74
            Top             =   240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Update File List"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "File List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   2775
         Begin VB.CommandButton Command29 
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2360
            TabIndex        =   76
            Top             =   2520
            Width           =   255
         End
         Begin VB.CommandButton Command28 
            Caption         =   "-"
            Height          =   255
            Left            =   2080
            TabIndex        =   75
            Top             =   2520
            Width           =   255
         End
         Begin VB.CommandButton Command26 
            Caption         =   "+"
            Height          =   255
            Left            =   2400
            TabIndex        =   70
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Random Song Length"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2880
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Random File Size"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   3120
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.ListBox List4 
            Height          =   2220
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text26 
            Height          =   270
            Left            =   960
            TabIndex        =   65
            Text            =   "kapster.mp3"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "FileName:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "File Flooder"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   3000
         TabIndex        =   50
         Top             =   1440
         Width           =   3735
         Begin VB.CommandButton Command4 
            Caption         =   "Halt Flood"
            Height          =   285
            Left            =   1440
            TabIndex        =   73
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Left            =   1200
            TabIndex        =   57
            Text            =   "goodmusic"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Begin Flood"
            Height          =   285
            Left            =   240
            TabIndex        =   56
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text5 
            Height          =   270
            Left            =   3000
            TabIndex        =   55
            Text            =   "0"
            Top             =   1440
            Width           =   615
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   3000
            Top             =   1800
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Single Send"
            Height          =   285
            Left            =   240
            TabIndex        =   54
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Height          =   270
            Left            =   1560
            TabIndex        =   53
            Text            =   "500"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox Text7 
            Height          =   270
            Left            =   1080
            TabIndex        =   52
            Text            =   "4294967296"
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1320
            TabIndex        =   51
            Text            =   "999999999"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "File String:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "X.mp3"
            Height          =   255
            Left            =   3000
            TabIndex        =   62
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "X:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   61
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   2400
            TabIndex        =   60
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "File Size:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label11 
            Caption         =   "File Length:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   1080
            Width           =   2415
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nifty/Dumb Stuff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   2
      Left            =   120
      TabIndex        =   39
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame Frame9 
         Caption         =   "Napster News"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   2400
         TabIndex        =   46
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton Command25 
            Caption         =   "Clear"
            Height          =   255
            Left            =   3000
            TabIndex        =   48
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox Text25 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   2775
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   47
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Chat Rooms"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton Command24 
            Caption         =   "Copy Room Name"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2680
            Width           =   1935
         End
         Begin VB.CommandButton Command23 
            Caption         =   "View All"
            Height          =   255
            Left            =   1080
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Root"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   42
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Currently 0 files (0 gigabytes) available in 0 libraries."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   6615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Chat Stuff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Timer Timer5 
         Interval        =   1
         Left            =   5880
         Top             =   2880
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Clear"
         Height          =   255
         Left            =   4080
         TabIndex        =   37
         Top             =   2880
         Width           =   735
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   6240
         Top             =   2880
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Say it"
         Height          =   285
         Left            =   3960
         TabIndex        =   36
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Text            =   "hey room i'm on kapster !"
         Top             =   3360
         Width           =   3735
      End
      Begin VB.ListBox List2 
         Height          =   2760
         Left            =   5160
         TabIndex        =   34
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   600
         Width           =   5055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Scorch It!"
         Height          =   285
         Left            =   5760
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Leave!"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Go There!"
         Height          =   285
         Left            =   3840
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1320
         TabIndex        =   28
         Text            =   "Punk"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Chatter(s)"
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Channel Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Login Stuff"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6855
      Begin VB.Frame Frame7 
         Height          =   680
         Left            =   480
         TabIndex        =   24
         Top             =   2760
         Width           =   2535
         Begin VB.CommandButton Command2 
            Caption         =   "Logout"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Login"
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   3255
         Begin VB.OptionButton Option2 
            Caption         =   "Main Server"
            Height          =   255
            Left            =   1800
            TabIndex        =   22
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Resolve Server"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.TextBox Text24 
            Height          =   270
            Left            =   2520
            TabIndex        =   19
            Text            =   "8875"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text23 
            Height          =   270
            Left            =   360
            TabIndex        =   18
            Text            =   "server.napster.com"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "IP                                        Port"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Socket Info"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   3480
         TabIndex        =   10
         Top             =   1560
         Width           =   3255
         Begin VB.TextBox Text22 
            Height          =   270
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text21 
            Height          =   270
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0.0.0.0"
            Top             =   240
            Width           =   1695
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   975
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1720
            _Version        =   393217
            TextRTF         =   $"frmKapster.frx":1E5E
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "status: idle"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label25 
            Caption         =   "Server:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label24 
            Caption         =   "Last Incoming Packet:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   3015
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Advanced"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   3255
         Begin VB.TextBox Text8 
            Height          =   270
            Left            =   1320
            TabIndex        =   23
            Text            =   "1 3697 0"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   270
            Left            =   1320
            TabIndex        =   7
            Text            =   "Kapster Client"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "??? String:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Client Type:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Basics"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   1200
            TabIndex        =   3
            Text            =   "krapster"
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1200
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IM"
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   148
      Top             =   200
      Width           =   495
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. 2"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   146
      Top             =   195
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   6600
      Picture         =   "frmKapster.frx":1F21
      ToolTipText     =   "About Kapster Client"
      Top             =   150
      Width           =   270
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   145
      Top             =   195
      Width           =   615
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chat !"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   144
      Top             =   195
      Width           =   735
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Misc."
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   143
      Top             =   195
      Width           =   735
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Library"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   142
      Top             =   195
      Width           =   735
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1 On 1"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   141
      Top             =   195
      Width           =   735
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   140
      Top             =   195
      Width           =   735
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Get'em"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   139
      Top             =   195
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00E0E0E0&
      Height          =   360
      Left            =   120
      Top             =   120
      Width           =   6840
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   360
      Left            =   120
      Top             =   120
      Width           =   6840
   End
End
Attribute VB_Name = "frmKapster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Soced and Coded By Wizdum(Robbie Saunders), Module from misc. coders
Dim TheData As String, A1, A2, A3, A4, A5, A6, A7, A8, A9, LB, F1 As Boolean, F2, F3, F4, F5, F6, F7, F8, F9
Dim IncomingStuff As String, TheDataa, PacketLen, ThePacket, LeftOver As Boolean
Dim ASDF1, MP3FILE, MP3LEN, MP3START As Boolean, T1 As Boolean

Private Sub Command1_Click()
On Error Resume Next
Command1.Enabled = False
Command2.Enabled = True
On Error Resume Next
Winsock1.Connect Text23, Text24
LeftOver = False
End Sub

Private Sub Command10_Click()
On Error Resume Next
If Command10.Caption = "Scorch It!" Then
Timer2.Enabled = True
S1 = True
Command10.Caption = "Stop It!"
Else
Command10.Caption = "Scorch It!"
Timer2.Enabled = False
Pause 40
S1 = False
End If
End Sub

Private Sub Command13_Click()
On Error Resume Next
For i = 1 To Text36
Winsock1.SendData Chr(Len(Text19)) & Chr(0) & Chr(239) & Chr(2) & Text19
DoEvents
Next i
End Sub

Private Sub Command14_Click()
On Error Resume Next
For i = 1 To Text36
Winsock1.SendData Chr(Len(Text19)) & Chr(0) & Chr(240) & Chr(2) & Text19
DoEvents
Next i
End Sub

Private Sub Command15_Click()
On Error Resume Next
F1 = True
List1.clear
LB = Len(Text14)
Winsock1.SendData Chr(LB) & Chr(0) & Chr(91) & Chr(2) & Text14
End Sub

Private Sub Command16_Click()
On Error Resume Next
If Option3.Value = 1 Then
For i = 1 To Text15
Winsock1.SendData Chr(0) & Chr(0) & Chr(218) & Chr(0)
DoEvents
Next i
Else
For i = 1 To Text15
Winsock1.SendData Chr(0) & Chr(0) & Chr(219) & Chr(0)
DoEvents
Next i
End If
End Sub

Private Sub Command17_Click()
On Error Resume Next
For i = 0 To List5.ListCount - 1
LB = Len(Text34 & " " & Chr(34) & List5.List(i) & Chr(34))
Winsock1.SendData Chr(LB) & Chr(0) & Chr(203) & Chr(0) & Text34 & " " & Chr(34) & List5.List(i) & Chr(34)
DoEvents
Next i
End Sub

Private Sub Command18_Click()
On Error Resume Next
If Option3.Value = 1 Then
For i = 1 To Text15
Winsock1.SendData Chr(0) & Chr(0) & Chr(220) & Chr(0)
DoEvents
Next i
Else
For i = 1 To Text15
Winsock1.SendData Chr(0) & Chr(0) & Chr(221) & Chr(0)
DoEvents
Next i
End If
End Sub

Private Sub Command19_Click()
On Error Resume Next
Text34 = Text27
List5.clear
For i = 1 To ListView1.ListItems.Count
List5.AddItem ListView1.ListItems(i)
DoEvents
Next i
Command6_Click (6)
End Sub

Private Sub Command2_Click()
On Error Resume Next
Command8_Click
Command1.Enabled = True
Command2.Enabled = False
Label4 = "status: idle"
Winsock1.Close
End Sub

Private Sub Command20_Click()
On Error Resume Next
LB = Len(Text18 & " " & Chr(34) & Text16 & Chr(34))
Winsock1.SendData Chr(LB) & Chr(0) & Chr(203) & Chr(0) & Text18 & " " & Chr(34) & Text16 & Chr(34)
End Sub

Private Sub Command21_Click()
On Error Resume Next
Text20 = ""
End Sub

Private Sub Command22_Click()
On Error Resume Next
List3.clear
Winsock1.SendData Chr(0) & Chr(0) & Chr(105) & Chr(2)
End Sub

Private Sub Command23_Click()
On Error Resume Next
List3.clear
Winsock1.SendData Chr(0) & Chr(0) & Chr(59) & Chr(3)
End Sub

Private Sub Command24_Click()
On Error Resume Next
Clipboard.clear
Clipboard.SetText List3.Text
End Sub

Private Sub Command25_Click()
On Error Resume Next
Text25 = ""
End Sub

Private Sub Command26_Click()
On Error Resume Next
For i = 0 To List4.ListCount - 1
If List4.List(i) = Text26 Then
MsgBox "The Server Ignores Files Of The Same Name ; )", vbCritical, "Error"
Exit Sub
End If
DoEvents
Next i
List4.AddItem Text26
End Sub

Private Sub Command27_Click()
On Error Resume Next
If Check3.Value = 1 Then
A1 = Chr(34) & "C:\" & Chr(34)
For i = 0 To List4.ListCount - 1
If Check1.Value = 1 Then
A3 = GetRandomInteger(1, 999999999)
Else
A3 = 999999999
End If
If Check2.Value = 1 Then
A4 = GetRandomInteger(1, 999999999)
Else
A4 = 999999999
End If
A1 = A1 & " " & Chr(34) & List4.List(i) & Chr(34) & " 5f08b9e815d392ecea30d4b5a8a01489-3431264 " & A3 & " 128 44100 " & A4
DoEvents
Next i
Winsock1.SendData VPbackwards(RealLen(Len(A1))) & Chr(102) & Chr(3) & A1
Else
If Check1.Value = 1 Then
A3 = GetRandomInteger(1, 999999999)
Else
A3 = 999999999
End If
If Check2.Value = 1 Then
A4 = GetRandomInteger(1, 999999999)
Else
A4 = 999999999
End If
For i = 0 To List4.ListCount - 1
Text5 = Text5 + 1
A1 = Chr(34) & "C:\" & Chr(34)
A1 = A1 & " " & Chr(34) & List4.List(i) & ".mp3" & Chr(34) & " 5f08b9e815d392ecea30d4b5a8a01489-3431264 " & A3 & " 128 44100 " & A4
Winsock1.SendData VPbackwards(RealLen(Len(A1))) & Chr(102) & Chr(3) & A1
DoEvents
Next i
End If
End Sub

Private Sub Command28_Click()
On Error Resume Next
List4.RemoveItem List4.ListIndex
End Sub

Private Sub Command29_Click()
On Error Resume Next
List4.clear
End Sub

Private Sub Command3_Click()
On Error Resume Next
Timer1.Enabled = True
End Sub

Private Sub Command30_Click()
On Error Resume Next
ListView1.ListItems.clear
Winsock1.SendData Chr(Len(Text27)) & Chr(0) & Chr(211) & Chr(0) & Text27
End Sub

Private Sub Command31_Click()
On Error Resume Next
ListView2.ListItems.clear
LB = Len("FILENAME CONTAINS " & Chr(34) & Text29 & Chr(34) & " MAX_RESULTS " & Text31 & " FILENAME CONTAINS " & Chr(34) & Text30 & Chr(34))
Winsock1.SendData Chr(LB) & Chr(0) & Chr(200) & Chr(0) & "FILENAME CONTAINS " & Chr(34) & Text29 & Chr(34) & " MAX_RESULTS " & Text31 & " FILENAME CONTAINS " & Chr(34) & Text30 & Chr(34)
End Sub

Private Sub Command33_Click()
On Error Resume Next
Text18 = TrimSpaces(ListView2.ListItems.Item(ListView2.SelectedItem.Index).SubItems(3))
Text16 = ListView2.ListItems.Item(ListView2.SelectedItem.Index)
Command6_Click (6)
End Sub

Private Sub Command34_Click()
On Error Resume Next
Clipboard.clear
Clipboard.SetText ListView2.ListItems.Item(ListView2.SelectedItem.Index).SubItems(3)
End Sub

Private Sub Command35_Click()
On Error Resume Next
Text18 = Text27
Text16 = ListView1.SelectedItem
Command6_Click (6)
End Sub

Private Sub Command36_Click()
On Error Resume Next
Clipboard.clear
Clipboard.SetText ListView2.ListItems.Item(ListView2.SelectedItem.Index)
End Sub

Private Sub Command38_Click()
On Error Resume Next
CommonDialog1.Filter = "MP3 File (*.mp3)|*.mp3|WMA File (*.wma)|*.wma|All Files (*.*)|*.*"
CommonDialog1.ShowSave
Text35 = CommonDialog1.FileName
End Sub

Private Sub Command39_Click()
On Error Resume Next
Dim RandomPort
RandomPort = GetRandomInteger(1, 9998) + 1
Winsock99.Close
Winsock99.LocalPort = RandomPort
Winsock99.Close
Winsock99.Listen
Winsock1.SendData Chr(Len(RandomPort)) & Chr(0) & Chr(191) & Chr(2) & RandomPort
Pause 0.3
LB = Len(Text18 & " " & Chr(34) & Text16 & Chr(34))
Winsock1.SendData Chr(LB) & Chr(0) & Chr(244) & Chr(1) & Text18 & " " & Chr(34) & Text16 & Chr(34)
Command39.Enabled = False
Command40.Enabled = True
End Sub

Private Sub Command4_Click()
On Error Resume Next
Timer1.Enabled = False
End Sub

Private Sub Command41_Click()
On Error Resume Next
Winsock1.SendData Chr(0) & Chr(0) & Chr(Text36) & Chr(Text37)
End Sub

Private Sub Command40_Click()
On Error Resume Next
Dim E1
If Len(MP3FILE) > 1000 Then
E1 = MsgBox("Save Existing File???", vbYesNo, "Save it?")
If E1 = vbYes Then
Open Text35 For Output As #87
Print #87, MP3FILE
DoEvents
Close #87
MsgBox "Download Saved!!"
MP3FILE = ""
End If
End If
Command40.Enabled = False
Command39.Enabled = True
Winsock2.Close
End Sub

Private Sub Command42_Click()
On Error Resume Next
If Text38 = Text9 Then
MsgBox "Error !"
Exit Sub
End If
List6.clear
T1 = True
LB = Len(Text38)
Winsock1.SendData Chr(LB) & Chr(0) & Chr(144) & Chr(1) & Text38
Pause 3
T1 = False
LB = Len(Text38)
Winsock1.SendData Chr(LB) & Chr(0) & Chr(145) & Chr(1) & Text38
End Sub

Private Sub Command43_Click()
On Error Resume Next
Label43 = "sending"
Winsock1.SendData Chr(0) & Chr(0) & Chr(238) & Chr(2)
End Sub

Private Sub Command45_Click()
On Error Resume Next
Winsock1.SendData Chr(Len(Text40)) & Chr(0) & Chr(189) & Chr(2) & Text40
End Sub

Private Sub Command46_Click()
On Error Resume Next
Winsock1.SendData Chr(Len(Text41)) & Chr(0) & Chr(188) & Chr(2) & Text41
End Sub

Private Sub Command47_Click()
On Error Resume Next
Winsock1.SendData Chr(Len(Text42)) & Chr(0) & Chr(190) & Chr(2) & Text42
End Sub

Private Sub Command48_Click()
On Error Resume Next
Winsock1.SendData Chr(Len(Text43)) & Chr(0) & Chr(191) & Chr(2) & Text43
End Sub

Private Sub Command49_Click(Index As Integer)
On Error Resume Next
Text44(Index) = Text44(Index) & Text1 & " : " & Text45(Index) & vbCrLf
LB = Len(Label49 & " " & Text45(Index))
Winsock1.SendData VPbackwards(RealLen(LB)) & Chr(205) & Chr(0) & Label49 & " " & Text45(Index)
End Sub

Private Sub Command5_Click()
On Error Resume Next
A1 = Chr(34) & "C:\" & Chr(34)
A2 = Text6
For i = 1 To Text6
Text6 = A2 - i
A1 = A1 & " " & Chr(34) & Text4 & i & ".mp3" & Chr(34) & " 5f08b9e815d392ecea30d4b5a8a01489-3431264 " & Text7 & " 128 44100 " & Text11
DoEvents
Next i
Winsock1.SendData VPbackwards(RealLen(Len(A1))) & Chr(102) & Chr(3) & A1
End Sub

Private Sub Command50_Click()
On Error Resume Next
NewMessageUser InputBox("What's His Name??", "Username")
End Sub

Private Sub Command51_Click()
On Error Resume Next
DelMessageUser TrimSpaces(Left(List9.Text, Len(List9.Text) - 40))
End Sub

Private Sub Command6_Click(Index As Integer)
On Error Resume Next
For i = 0 To 15
Frame2(i).Visible = False
Command6(i).FontBold = False
Next i
Frame2(Index).Visible = True
Command6(Index).FontBold = True
Image2.BorderStyle = 0
End Sub

Private Sub Command7_Click()
On Error Resume Next
LB = Len(Text9)
Winsock1.SendData Chr(LB) & Chr(0) & Chr(144) & Chr(1) & Text9
Command8.Enabled = True
Command7.Enabled = False
Text9.Enabled = False
End Sub

Private Sub Command8_Click()
On Error Resume Next
If S1 = False Then
List2.clear
Command7.Enabled = True
Command8.Enabled = False
Text9.Enabled = True
End If
LB = Len(Text9)
Winsock1.SendData Chr(LB) & Chr(0) & Chr(145) & Chr(1) & Text9
End Sub

Private Sub Command9_Click()
On Error Resume Next
LB = Len(Text9) + Len(Text10) + 1
Winsock1.SendData Chr(LB) & Chr(0) & Chr(146) & Chr(1) & Text9 & " " & Text10
Text10 = ""
End Sub

Private Sub Form_Load()
LoadSettings
Winsock99.Close
Winsock99.Listen
StayOnTop Me
End Sub

Private Sub form_Unload(Cancel As Integer)
SaveSettings
End
End Sub

Private Sub Image2_Click()
On Error Resume Next
For i = 0 To 15
Frame2(i).Visible = False
Command6(i).FontBold = False
Next i
Frame2(9).Visible = True
End Sub

Private Sub List9_Click()
On Error Resume Next
ShowMessages Left(List9.Text, Len(List9.Text) - 40)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command9_Click
End If
End Sub

Private Sub Text45_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command49_Click (Index)
Text45(Index) = ""
End If
End Sub

Private Sub Timer1_Timer()
Text5 = Text5 + 1
A1 = Chr(34) & "C:\" & Chr(34)
A1 = A1 & " " & Chr(34) & Text4 & Text5 & ".mp3" & Chr(34) & " 5f08b9e815d392ecea30d4b5a8a01489-3431264 " & Text7 & " 128 44100 " & Text11
Winsock1.SendData VPbackwards(RealLen(Len(A1))) & Chr(102) & Chr(3) & A1
End Sub

Private Sub Timer2_Timer()
Command7_Click
Command8_Click
End Sub

Private Sub Timer3_Timer()
Dim ASDF2, ASDF3
ASDF2 = (ASDF1 / 3) / 1000
ASDF3 = InStr(1, ASDF2, ".")
ASDF1 = Left(ASDF2, ASDF3 + 1)
Label39 = ASDF1 & "k a second"
ASDF1 = 0
End Sub

Private Sub Timer4_Timer()
LB = Len(Text12 & " " & Text13)
Winsock1.SendData VPbackwards(RealLen(LB)) & Chr(205) & Chr(0) & Text12 & " " & Text13
LB = Len(Text12 & " " & Text13)
Winsock1.SendData VPbackwards(RealLen(LB)) & Chr(205) & Chr(0) & Text12 & " " & Text13
LB = Len(Text12 & " " & Text13)
Winsock1.SendData VPbackwards(RealLen(LB)) & Chr(205) & Chr(0) & Text12 & " " & Text13
LB = Len(Text12 & " " & Text13)
Winsock1.SendData VPbackwards(RealLen(LB)) & Chr(205) & Chr(0) & Text12 & " " & Text13
LB = Len(Text12 & " " & Text13)
Winsock1.SendData VPbackwards(RealLen(LB)) & Chr(205) & Chr(0) & Text12 & " " & Text13
Label15 = Label15 + 5
End Sub

Private Sub Timer5_Timer()
Label27 = List2.ListCount & " Chatter(s)"
End Sub

Private Sub Timer6_Timer()
Label22 = List9.ListCount & " Users"
End Sub

Private Sub Winsock1_Connect()
If Option1.Value = True And Winsock1.RemotePort = Text24 Then
Label4 = "status: connected, resolving..."
Else
Label4 = "status: connected, logging in..."
LB = Len(Text1 & " " & Text2 & " 6699 " & Chr(34) & Text3 & Chr(34) & " " & Text8)
Winsock1.SendData Chr(LB) & Chr(0) & Chr(2) & Chr(0) & Text1 & " " & Text2 & " 6699 " & Chr(34) & Text3 & Chr(34) & " " & Text8
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Winsock1.GetData TheData
Open "C:\g.f" For Output As #87
Print #87, TheData
Close #87
RichTextBox1.LoadFile "C:\g.f"
'resolve the server
If Option1.Value = True And Winsock1.RemotePort = Text24 Then
A1 = InStr(1, TheData, ":")
If A1 <> 0 Then
A2 = Mid(TheData, 1, A1 - 1)
A3 = Mid(TheData, A1 + 1, 4)
Text21 = A2
Text22 = A3
Label4 = "status: " & A2 & " " & A3
Winsock1.Close
Winsock1.Connect A2, A3
Else
Winsock1.Close
MsgBox "Protocol Error! Disconnecting...", vbCritical, "Kapster"
Label4 = "status: idle"
End If
Else
'my new super leet splitter heh
IncomingStuff = TheData
TheDataa = TheDataa & IncomingStuff
NewOne:
If LeftOver = False Then PacketLen = Asc(Mid(TheDataa, 1, 1)) + (Asc(Mid(TheDataa, 2, 1)) * 256) + 4
If PacketLen > Len(TheDataa) Then
    LeftOver = True
    Exit Sub
End If
ThePacket = Left(TheDataa, PacketLen)
ProcessData ThePacket
TheDataa = Right(TheDataa, Len(TheDataa) - PacketLen)
LeftOver = False
If Len(TheDataa) <> 0 Then GoTo NewOne:
End If
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Label8 = bytesRemaining
End Sub

Sub ProcessData(TheData)
'u're logged in!
If Mid(TheData, 3, 1) = Chr(3) Then
For i = 1 To 8
Command6(i).Enabled = True
DoEvents
Next i
Label4 = "status: logged in"
End If
'grab finger results
If F1 = True Then
F1 = False
F2 = InStr(1, TheData, " " & Chr(34))
List1.AddItem "nickname: " & Mid(TheData, 5, F2 - 4)
TheData = Right(TheData, Len(TheData) - 1 - F2)
F2 = InStr(1, TheData, Chr(34) & " ")
List1.AddItem "user level: " & Mid(TheData, 1, F2 - 1)
TheData = Right(TheData, Len(TheData) - 1 - F2)
F2 = InStr(1, TheData, " " & Chr(34))
List1.AddItem "time online: " & Mid(TheData, 1, F2 - 1) & " seconds"
TheData = Right(TheData, Len(TheData) - 1 - F2)
F2 = InStr(1, TheData, Chr(34) & " ")
List1.AddItem "channel(s): " & Replace(Mid(TheData, 1, F2 - 2), " ", ",")
TheData = Right(TheData, Len(TheData) - 2 - F2)
F2 = InStr(1, TheData, Chr(34) & " ")
List1.AddItem "ability: " & Mid(TheData, 1, F2 - 1)
TheData = Right(TheData, Len(TheData) - 1 - F2)
F2 = InStr(1, TheData, " ")
List1.AddItem "# of files: " & Mid(TheData, 1, F2 - 1)
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
List1.AddItem "# downloading: " & Mid(TheData, 1, F2 - 1)
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
List1.AddItem "# uploading: " & Mid(TheData, 1, F2 - 1)
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
List1.AddItem "connection: " & Mid(TheData, 1, F2 - 1)
TheData = Right(TheData, Len(TheData) - F2)
List1.AddItem "client: " & Replace(TheData, Chr(34), "")
End If
'grab room enter (fixed)
If S1 = False Then
If Mid(TheData, 3, 1) = Chr(152) Then
If Mid(TheData, 5, Len(Text38)) = Text38 Then
TheData = Right(TheData, Len(TheData) - 5 - Len(Text38))
If InStr(1, TheData, Text1) = 0 Then
List6.AddItem Left(TheData, F2)
End If
ElseIf Mid(TheData, 5, Len(Text9)) = Text9 Then
TheData = Right(TheData, Len(TheData) - 5 - Len(Text9))
F2 = InStr(1, TheData, " ")
For i = 0 To List2.ListCount - 1
If Left(List2.List(i), Len(Left(TheData, F2))) = Left(TheData, F2) Then Exit Sub
DoEvents
Next i
List2.AddItem Left(TheData, F2)
End If
End If
End If
'grab room exit
If T1 = False Then
If Mid(TheData, 3, 1) = Chr(151) Then
If Mid(TheData, 5, Len(Text9)) = Text9 Then
TheData = Right(TheData, Len(TheData) - 5 - Len(Text9))
F2 = InStr(1, TheData, " ")
AddChatText "KAPSTER ", UCase(Left(TheData, F2)) & " LEFT THE ROOM"
For i = 0 To List2.ListCount - 1
If List2.List(i) = Left(TheData, F2) Then
List2.RemoveItem i
End If
DoEvents
Next i
End If
End If
'grab chat text
If Mid(TheData, 3, 1) = Chr(147) Then
If Mid(TheData, 5, Len(Text9)) = Text9 Then
TheData = Right(TheData, Len(TheData) - 5 - Len(Text9))
F2 = InStr(1, TheData, " ")
AddChatText Left(TheData, F2), Right(TheData, Len(TheData) - F2)
End If
End If
End If
'grab server stats
If Mid(TheData, 3, 1) = Chr(214) Then
TheData = Right(TheData, Len(TheData) - 4)
F2 = InStr(1, TheData, " ")
F3 = Mid(TheData, 1, F2 - 1)
TheData = Replace(TheData, F3 & " ", "")
F2 = InStr(1, TheData, " ")
F4 = Mid(TheData, 1, F2 - 1)
TheData = Replace(TheData, F4 & " ", "")
F5 = Right(TheData, Len(TheData))
Label28 = "Currently " & F4 & " files (" & F5 & " gigabytes) available in " & F3 & " libraries @ " & Time
End If
'grab chat rooms
If Mid(TheData, 3, 1) = Chr(106) Or Mid(TheData, 3, 1) = Chr(60) Then
TheData = Right(TheData, Len(TheData) - 4)
F2 = InStr(1, TheData, " ")
F3 = Left(TheData, F2)
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
List3.AddItem F3 & "(" & Left(TheData, F2 - 1) & ")"
End If
'grab napster news
If Mid(TheData, 3, 1) = Chr(109) Then
Text25 = Text25 & Right(TheData, Len(TheData) - 4) & vbCrLf
End If
'grab user's file list
If Mid(TheData, 3, 1) = Chr(212) Then
If Mid(TheData, 5, Len(Text27)) = Text27 Then
TheData = Right(TheData, Len(TheData) - 6 - Len(Text27))
F2 = InStr(1, TheData, Chr(34) & " ")
F3 = Left(TheData, F2 - 1)
TheData = Right(TheData, Len(TheData) - 1 - F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
AddSongHotList F3, Left(TheData, F2)
End If
End If
'grab user's ip #
If Mid(TheData, 3, 1) = Chr(213) Then
If Mid(TheData, 5, Len(Text27)) = Text27 Then
TheData = Right(TheData, Len(TheData) - 5 - Len(Text27))
Text28 = Val(TheData)
End If
End If
'grab search results
If Mid(TheData, 3, 1) = Chr(201) Then
TheData = Right(TheData, Len(TheData) - 5)
F2 = InStr(1, TheData, Chr(34) & " ")
F3 = Left(TheData, F2 - 1)
TheData = Right(TheData, Len(TheData) - 1 - F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
F4 = Left(TheData, F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
F5 = Left(TheData, F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
F6 = Left(TheData, F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - F2)
F2 = InStr(1, TheData, " ")
F7 = Left(TheData, F2)
AddSongSearch F3, F4, F5 & "s", F6, F7
End If
'grab search finished
If Mid(TheData, 3, 1) = Chr(202) Then
MsgBox "Search Results Finished!!", vbInformation, "Done!"
End If
'grab user's ip and port
If Mid(TheData, 3, 1) = Chr(204) Then
TheData = Right(TheData, Len(TheData) - 4)
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - 1 - F2)
F2 = InStr(1, TheData, " ")
Text32 = IPToStrin(Val(Left(TheData, F2)))
F2 = InStr(1, TheData, " ")
TheData = Right(TheData, Len(TheData) - 1 - F2)
Text33 = Val(Left(TheData, F2))
End If
'grab server ping
If Mid(TheData, 3, 1) = Chr(238) Then
Label43 = "success!"
End If
'grab incoming ims
If Mid(TheData, 3, 1) = Chr(205) Then
TheData = Right(TheData, Len(TheData) - 4)
F2 = InStr(1, TheData, " ")
F3 = Left(TheData, F2 - 1)
TheData = Right(TheData, Len(TheData) - F2)
NewMessageUser F3
AddIMessage F3, Left(TheData, Len(TheData))
End If
'grab user offline in ims
If Mid(TheData, 3, 1) = Chr(148) Then
TheData = Right(TheData, Len(TheData) - 4)
TheData = Replace(TheData, "user ", "")
TheData = Replace(TheData, " is not online", "")
NewMessageUser TheData
i = GetUserIndex(TheData)
Text44(i) = Text44(i) & "KAPSTER : " & UCase(TheData) & " IS NOT ONLINE!" & vbCrLf
End If
End Sub

Sub AddChatText(Namer, Texter)
Text20 = Text20 & Namer & ": " & Texter & vbCrLf
End Sub

Sub AddSongHotList(SongNaMe, SongSiZe)
Set listsub = ListView1.ListItems.Add(, , SongNaMe)
listsub.SubItems(1) = SongSiZe
End Sub

Sub AddSongSearch(SongNaMe, SongSiZe, SongLeN, UserNaMe, UserSpEeD)
Set listsub = ListView2.ListItems.Add(, , SongNaMe)
listsub.SubItems(1) = SongSiZe
listsub.SubItems(2) = SongLeN
listsub.SubItems(3) = UserNaMe
listsub.SubItems(4) = UserSpEeD
End Sub

Sub NewMessageUser(UserNaMe)
On Error Resume Next
For i = 0 To List9.ListCount - 1
If Left(List9.List(i), Len(UserNaMe)) = UserNaMe Then
Exit Sub
End If
DoEvents
Next i
Label48 = Label48 + 1
List9.AddItem UserNaMe & String(40, " ") & Label48
Load Text44(Label48)
Load Text45(Label48)
Load Command49(Label48)
Text45(Label48) = "howdy partner"
ShowMessages UserNaMe
End Sub

Sub DelMessageUser(UserNaMe)
On Error Resume Next
Text44(0).Visible = True
Text45(0).Visible = True
Command49(0).Visible = True
Unload Text44(GetUserIndex(UserNaMe))
Unload Text45(GetUserIndex(UserNaMe))
Unload Command49(GetUserIndex(UserNaMe))
Text45(0) = "user deleted"
List9.RemoveItem List9.ListIndex
End Sub

Sub AddIMessage(UserNaMe, IMesSaGe)
On Error Resume Next
i = GetUserIndex(UserNaMe)
Text44(i) = Text44(i) & UserNaMe & " : " & IMesSaGe & vbCrLf
End Sub

Sub ShowMessages(UserNaMe)
On Error Resume Next
For i = 0 To Label48
Text44(i).Visible = False
Text45(i).Visible = False
Command49(i).Visible = False
DoEvents
Next i
Label49 = UserNaMe
Text44(GetUserIndex(UserNaMe)).Left = 120
Text44(GetUserIndex(UserNaMe)).Top = 600
Text44(GetUserIndex(UserNaMe)).Visible = True
Text44(GetUserIndex(UserNaMe)).Enabled = True
Text45(GetUserIndex(UserNaMe)).Left = 120
Text45(GetUserIndex(UserNaMe)).Top = 3360
Text45(GetUserIndex(UserNaMe)).Visible = True
Text45(GetUserIndex(UserNaMe)).Enabled = True
Command49(GetUserIndex(UserNaMe)).Left = 3960
Command49(GetUserIndex(UserNaMe)).Top = 3360
Command49(GetUserIndex(UserNaMe)).Visible = True
Command49(GetUserIndex(UserNaMe)).Enabled = True
End Sub

Function GetUserIndex(UserNaMe)
On Error Resume Next
For i = 0 To List9.ListCount - 1
If Left(List9.List(i), Len(UserNaMe)) = UserNaMe Then
GetUserIndex = TrimSpaces(Right(List9.List(i), 10))
End If
DoEvents
Next i
End Function

Public Function GetRandomInteger(LowerBound, UpperBound) As Long
GetRandomInteger = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim GR1 As String, K1, K2, K3, K4
Winsock2.GetData GR1
If GR1 = "SEND" Then MsgBox "user started send": Exit Sub
If InStr(1, GR1, Text18) <> 0 Then
    GR1 = Right(GR1, Len(GR1) - Len(Text18 & "  " & Text16 & "  "))
    For i = 1 To 11
        If Asc(Mid(GR1, i, 1)) < 48 Or Asc(Mid(GR1, i, 1)) > 57 Then
            F3 = Left(GR1, i - 1)
            GR1 = Right(GR1, Len(GR1) - (i - 1))
            GoTo GotLength:
        End If
    Next i
GotLength:
    MP3LEN = F3
    ASDF1 = 0
    MP3FILE = ""
End If
MP3FILE = MP3FILE & GR1
Label40 = Len(MP3FILE) & "/" & MP3LEN
ASDF1 = ASDF1 + bytesTotal
DisplayPercent Len(MP3FILE) / MP3LEN
If Len(MP3FILE) > MP3LEN - 1000 Then
    Open Text35 For Output As #87
        Print #87, MP3FILE
        DoEvents
    Close #87
    MsgBox "Download Complete!!"
    Winsock2.Close
    MP3FILE = ""
End If
DoEvents
End Sub

Private Sub Winsock99_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Winsock2.Accept requestID
MsgBox "remote connection received!"
LB = Len(Text18 & " " & Chr(34) & Text16 & Chr(34))
Winsock2.SendData Chr(LB) & Chr(0) & Chr(203) & Chr(0) & Text18 & " " & Chr(34) & Text16 & Chr(34)
End Sub

Sub DisplayPercent(Percenter)
Dim Y1, Y2, Y3
Y1 = InStr(1, Percenter * 100, ".")
If Y1 <> 0 Then
Y2 = Left(Percenter * 100, Y1 - 1)
End If
Label38 = Y2
Label37.Height = Percenter * 615
Label37.Top = 600 + (615 - Label37.Height)
End Sub

Sub SaveSettings()
Dim TT1 As String
Text12 = Text1 & vbCrLf & base64_encode(Text2)
TT1 = App.Path
If Right(TT1, 1) <> "\" Then TT1 = TT1 & "\kap.ini"
Open TT1 For Output As #42
Print #42, Text12
Close #42
End Sub

Sub LoadSettings()
Dim TT1 As String, TT2
TT1 = App.Path
If Right(TT1, 1) <> "\" Then TT1 = TT1 & "\kap.ini"
If FileExist(TT1) = True Then
Open TT1 For Input As #43
Text12.Text = Input$(LOF(43), #43)
Close #43
TT2 = Split(Text12.Text, vbCrLf)
Text1 = TT2(0)
Text2 = base64_decode(TT2(1))
Else
Text1 = "krapster"
Text2 = ""
End If
End Sub

