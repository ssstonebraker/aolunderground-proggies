VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form MainLarge 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quad-Ball By Arvinder Sehmi 1999, Arvinder@Sehmi.Co.Uk"
   ClientHeight    =   8985
   ClientLeft      =   150
   ClientTop       =   405
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainLarge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox StatBox 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   0
      ScaleHeight     =   1350
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Score"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   510
         Left            =   900
         TabIndex        =   12
         Top             =   1890
         Width           =   3660
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Planet Satatus"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   555
         Left            =   900
         TabIndex        =   11
         Top             =   3015
         Width           =   4110
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Top Speed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   555
         Left            =   900
         TabIndex        =   10
         Top             =   5670
         Width           =   3885
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Number Of Bounces"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   645
         Left            =   900
         TabIndex        =   9
         Top             =   6750
         Width           =   6180
      End
      Begin VB.Label TotalScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SCORE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   825
         Left            =   4005
         TabIndex        =   8
         Top             =   1755
         Width           =   2490
      End
      Begin VB.Label SavedPlanets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Earth"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   825
         Index           =   1
         Left            =   4770
         TabIndex        =   7
         Top             =   2970
         Width           =   1785
      End
      Begin VB.Label TopSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TopSpeed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   825
         Left            =   4005
         TabIndex        =   6
         Top             =   5490
         Width           =   3090
      End
      Begin VB.Label TotalBounces 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TotalBounces"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   825
         Left            =   7200
         TabIndex        =   5
         Top             =   6615
         Width           =   4185
      End
      Begin VB.Label SavedPlanets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mars"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   825
         Index           =   2
         Left            =   4770
         TabIndex        =   4
         Top             =   3735
         Width           =   1635
      End
      Begin VB.Label SavedPlanets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neptune"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   825
         Index           =   3
         Left            =   4770
         TabIndex        =   3
         Top             =   4455
         Width           =   2580
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Statistics"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   1095
         Left            =   900
         TabIndex        =   2
         Top             =   270
         Width           =   5985
      End
      Begin VB.Label OK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Continue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   825
         Left            =   8460
         TabIndex        =   1
         Top             =   7650
         Width           =   2805
      End
   End
   Begin VB.PictureBox StoryScreen 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   2000
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   2000
      Begin VB.Label StoryLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2460
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   12075
         WordWrap        =   -1  'True
      End
      Begin VB.Label Start2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Game"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   1800
         TabIndex        =   15
         Top             =   8145
         Width           =   3225
      End
      Begin VB.Label Quit2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quit Game"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   6750
         TabIndex        =   14
         Top             =   8145
         Width           =   2955
      End
   End
   Begin VB.PictureBox TitleScreen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8805
      Left            =   0
      ScaleHeight     =   8805
      ScaleWidth      =   7590
      TabIndex        =   45
      Top             =   0
      Width           =   7590
      Begin QuadBall_Story.ArviScroll Scroll 
         Height          =   330
         Left            =   225
         TabIndex        =   46
         Top             =   8550
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Start 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Game"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   6840
         TabIndex        =   51
         Top             =   5040
         Width           =   3225
      End
      Begin VB.Label Quit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quit Game"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   6885
         TabIndex        =   50
         Top             =   6705
         Width           =   2955
      End
      Begin VB.Label Story 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Story"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   6885
         TabIndex        =   49
         Top             =   5580
         Width           =   3075
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Quad-Ball Story Mode."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   555
         Left            =   225
         TabIndex        =   48
         Top             =   7830
         Width           =   11490
      End
      Begin VB.Label TrainingMode 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Play In Training Mode"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   630
         Left            =   5940
         TabIndex        =   47
         Top             =   6165
         Width           =   4965
      End
   End
   Begin VB.PictureBox Area 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H000000FF&
      Height          =   9075
      Left            =   0
      MouseIcon       =   "MainLarge.frx":09BA
      MousePointer    =   99  'Custom
      ScaleHeight     =   9075
      ScaleWidth      =   10320
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   10320
      Begin VB.PictureBox Ball 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   1  'Dash
         Enabled         =   0   'False
         FillColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5085
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   39
         Top             =   4455
         Width           =   375
      End
      Begin VB.PictureBox TPad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         ScaleHeight     =   330
         ScaleWidth      =   2085
         TabIndex        =   38
         Tag             =   "Top"
         Top             =   0
         Width           =   2085
      End
      Begin VB.PictureBox BPad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         ScaleHeight     =   330
         ScaleWidth      =   2085
         TabIndex        =   37
         Tag             =   "Top"
         Top             =   8685
         Width           =   2085
      End
      Begin VB.PictureBox LPad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2085
         Left            =   0
         ScaleHeight     =   2085
         ScaleWidth      =   330
         TabIndex        =   36
         Tag             =   "Side"
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox RPad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2085
         Left            =   9990
         ScaleHeight     =   2085
         ScaleWidth      =   330
         TabIndex        =   35
         Tag             =   "Side"
         Top             =   45
         Width           =   330
      End
      Begin PicClip.PictureClip AniBall 
         Left            =   2160
         Top             =   2205
         _ExtentX        =   2646
         _ExtentY        =   1323
         _Version        =   393216
         Rows            =   2
         Cols            =   4
      End
      Begin PicClip.PictureClip BallBlow 
         Left            =   2025
         Top             =   3375
         _ExtentX        =   5292
         _ExtentY        =   5080
         _Version        =   393216
         Rows            =   4
         Cols            =   4
      End
      Begin VB.Shape Limit 
         BackColor       =   &H00000000&
         BorderColor     =   &H0000FF00&
         Height          =   8370
         Left            =   360
         Top             =   315
         Visible         =   0   'False
         Width           =   9675
      End
      Begin VB.Label GameStory 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2085
         Left            =   2895
         TabIndex        =   40
         Top             =   3555
         Width           =   5970
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Info 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   8985
      Left            =   10320
      ScaleHeight     =   8955
      ScaleWidth      =   1650
      TabIndex        =   17
      Top             =   0
      Width           =   1680
      Begin VB.Timer TimeKeeper 
         Enabled         =   0   'False
         Interval        =   900
         Left            =   1260
         Top             =   3690
      End
      Begin VB.Label Reset 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Left            =   225
         TabIndex        =   29
         Tag             =   "no"
         Top             =   5895
         Width           =   1185
      End
      Begin VB.Image InfoArea 
         Height          =   9075
         Left            =   -45
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   1770
      End
      Begin VB.Image Life 
         Height          =   375
         Index           =   2
         Left            =   630
         Top             =   2610
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF80FF&
         Height          =   960
         Left            =   90
         Top             =   3240
         Width           =   1410
      End
      Begin VB.Label LYLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   330
         Left            =   1080
         TabIndex        =   33
         Top             =   3870
         Width           =   375
      End
      Begin VB.Label CommetSpeed 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   45
         TabIndex        =   32
         Top             =   3870
         Width           =   1005
      End
      Begin VB.Label CometSpeedLab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comet Speed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Left            =   45
         TabIndex        =   31
         Top             =   3240
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   825
         Left            =   90
         Top             =   90
         Width           =   1410
      End
      Begin VB.Label PlanetsLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Planets"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   420
         Left            =   90
         TabIndex        =   30
         Top             =   2115
         Width           =   1410
      End
      Begin VB.Shape PlanetsBox 
         BorderColor     =   &H00FF8080&
         Height          =   1050
         Left            =   90
         Top             =   2115
         Width           =   1410
      End
      Begin VB.Image Life 
         Height          =   375
         Index           =   3
         Left            =   180
         Top             =   2610
         Width           =   375
      End
      Begin VB.Image Life 
         Height          =   375
         Index           =   1
         Left            =   1080
         Top             =   2610
         Width           =   375
      End
      Begin VB.Label AimTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   90
         TabIndex        =   28
         Top             =   1710
         Width           =   1410
      End
      Begin VB.Label TimeSoFar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   90
         TabIndex        =   27
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Label TimeLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   420
         Left            =   90
         TabIndex        =   26
         Top             =   990
         Width           =   1410
      End
      Begin VB.Shape TimeBox 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   1050
         Left            =   90
         Top             =   990
         Width           =   1410
      End
      Begin VB.Label Score 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   90
         TabIndex        =   25
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label ScoreLable 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   420
         Left            =   90
         TabIndex        =   24
         Top             =   90
         Width           =   1410
      End
      Begin VB.Shape ScoreBox 
         Height          =   825
         Left            =   90
         Top             =   -5000
         Width           =   1410
      End
      Begin VB.Label LifeTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   3
         Left            =   90
         TabIndex        =   23
         Top             =   2610
         Width           =   510
      End
      Begin VB.Label LifeTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   2
         Left            =   585
         TabIndex        =   22
         Top             =   2610
         Width           =   510
      End
      Begin VB.Label LifeTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   1
         Left            =   1035
         TabIndex        =   21
         Top             =   2610
         Width           =   510
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFF00&
         Height          =   1320
         Left            =   90
         Top             =   4275
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "High Score"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   90
         TabIndex        =   20
         Top             =   4275
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label HighestScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   90
         TabIndex        =   19
         Top             =   4545
         Width           =   1455
      End
      Begin VB.Label HighName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   600
         Left            =   90
         TabIndex        =   18
         Top             =   4950
         Width           =   1455
      End
   End
   Begin VB.PictureBox KillEarth 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   0
      ScaleHeight     =   6000
      ScaleWidth      =   6495
      TabIndex        =   41
      Top             =   0
      Width           =   6495
      Begin QuadBall_Story.EarthBlow1 BlowUpEarth 
         Height          =   1380
         Left            =   450
         TabIndex        =   42
         Top             =   810
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   2434
      End
      Begin VB.PictureBox KillBall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   1  'Dash
         Enabled         =   0   'False
         FillColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   43
         Top             =   0
         Width           =   375
      End
      Begin VB.Label GameStory2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2085
         Left            =   3150
         TabIndex        =   44
         Top             =   3375
         Width           =   5970
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "MainLarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' TO SEE DOCUMENTATION ON WHAT THESE SUBS DO,   '
' GOTO "MainSmall", As This Code Is Very Similar'
' And As Only Been Slightly Adjusted To Work    '
' On Displays Larger Than 800x600               '
'_______________________________________________'
Dim ExitBounce As Boolean
Dim StartTime As Date
Private Sub TrainingMode_Click()
 On Error GoTo a:
 ThisDir
 Shell "QuadBall_Training.Exe " & Trim(Str(CmdSpeedParam)), vbNormalFocus
 Reset_Click
 Unload Me
 End
a:
MsgBox "Cannot Find QuadBall_Training.Exe," & Chr(13) & "Please Re-install This Game To Fix The Problem." & Chr(13) & _
        "If You Can Find QuadBall_Training.exe On Your PC," & Chr(13) & "Please Place It In The Directory:" & Chr(13) & _
        App.Path, vbCritical, "Error"
End Sub
Private Sub TrainingMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight Start
 Highlight TrainingMode
 UnHighlight Quit
 UnHighlight Story
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Reset_Click
 Call StopSounds(True, True)
 ExitBounce = True
 Call StopSounds(True, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Call StopSounds(True, True)
End Sub
Private Sub GameStory_Change()
With GameStory
 .Left = Int(Me.Width - Info.Width) / 2
 .Top = Int(Me.Height / 2)
 .Left = .Left - Int(.Width / 2)
 .Top = .Top - Int(.Height / 2)
 .Refresh
End With
End Sub
Private Sub Info_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight Reset
End Sub

Private Sub InfoArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 X = Int(Area.Width - (BPad.Width / 2))
 Call Area_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub OK_Click()
 Reset_Click
 StatBox.Visible = False
End Sub
Private Sub OK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Highlight OK
End Sub
Private Sub Quit_Click()
 ExitBounce = True
 WAVPlay "exit.qbs"
 Unload Me
 End
End Sub
Private Sub Quit2_Click()
 Call Quit_Click
End Sub
Public Sub CentreBall()
 Ball.Visible = False
 Ball.Top = Int((Me.Height + Ball.Height) / 2)
 Ball.Left = Int(Int((Area.Width - Info.Width) / 2) - Int(Ball.Width / 2))
End Sub
Private Sub Reset_Click()
 WAVPlay "exit.qbs"
 ExitBounce = True
 ExitBounce = True
 CentreBall
 TitleScreen.Visible = True
 Scroll.StartScroll
 Area.Visible = False
 LifeTxt(1).caption = ""
 LifeTxt(2).caption = ""
 LifeTxt(3).caption = ""
 LivesLeft = 3
 XSpeed = CmdSpeedParam
 YSpeed = CmdSpeedParam
 For i = 1 To 3
  Life(1).Visible = True
 Next i
 TimeSoFar = "00:00:00"
 AimTime = "00:00:00"
 Score = "0"
End Sub
Private Sub Reset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Highlight Reset
End Sub
Private Sub Start_Click()
 WAVPlay "start.qbs"
 Area.Visible = True
 Area.Cls
 TitleScreen.Visible = False
 Scroll.ExitScroll
 LoadScore
 LivesLeft = 3
 XSpeed = CmdSpeedParam
 YSpeed = CmdSpeedParam
  For i = 1 To 3
   Life(i).Visible = True
   LifeTxt(i).caption = ""
  Next i
 TimeSoFar = "00:00:00"
 AimTime = "00:00:25"
 Score = "0"
 Ball.Visible = False
 GameStory = "Your First Mission Is To Save neptune!" & Chr(13) & " Your Target Time is 0.25 Minuets"
 Delay 1.5
 GameStory = "Ready"
 Delay 0.8
 GameStory = "GO!"
 Delay 0.6
 GameStory = ""
 Ball.Visible = True
 ExitBounce = False
 Bounce
End Sub
Private Sub Start2_Click()
 StoryScreen.Visible = False
 Call Start_Click
End Sub
Private Sub Start_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight TrainingMode
 UnHighlight Story
 UnHighlight Quit
 Highlight Start
End Sub
Private Sub Start2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight Quit2
 Highlight Start2
End Sub
Private Sub Quit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight TrainingMode
 UnHighlight Story
 UnHighlight Start
 Highlight Quit
End Sub
Private Sub Quit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight Start2
 Highlight Quit2
End Sub
Private Sub ShowStats()
TotalScore = Val(Score)
TopSpeed = FastSpeed
TotalBounces = NumBounces
 If LifeTxt(3).caption = "OK" Then
  SavedPlanets(3).caption = "Neptune: Saved!"
 Else
  SavedPlanets(3).caption = "Neptune: Destroyed!"
 End If
 If LifeTxt(2).caption = "OK" Then
  SavedPlanets(2).caption = "Mars: Saved!"
 Else
  SavedPlanets(2).caption = "Mars: Destroyed!"
 End If
 If LifeTxt(1).caption = "OK" Then
  SavedPlanets(1).caption = "Earth: Saved!"
 Else
  SavedPlanets(1).caption = "Earth: Destroyed!"
 End If
StatBox.Visible = True
 If TotalScore > Val(GetKeyValue(HKEY_LOCAL_MACHINE, _
  "SOFTWARE\arvisehmi\QuadBall", "TopScore")) Then
  On Error Resume Next
  InputWindow.Visible = True
 End If
End Sub

Private Sub StatBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight OK
End Sub
Private Sub totalbounces_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight OK
End Sub
Private Sub label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight OK
End Sub

Private Sub Story_Click()
 LoadStoryString
 StoryScreen.Visible = True
End Sub
Private Sub Story_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight TrainingMode
 UnHighlight Start
 UnHighlight Quit
 Highlight Story
End Sub
Private Sub StoryLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight Quit2
 UnHighlight Start2
End Sub
Private Sub TimeKeeper_Timer()
 Dim TempTime As Date
 TempTime = Time - StartTime
 TimeSoFar = TempTime
 If TimeSoFar = AimTime Then SavePlanet
End Sub
Public Sub SavePlanet()
 LivesLeft = LivesLeft - 1
 ExitBounce = True
 WAVLoop "Shoot.qbs"
 Ball.Visible = True
 For i = 300 To 0 Step -10
  Area.Line (0, i)-(Ball.Left, Ball.Top), RGB(180, 0, 0)
  Area.Line (i, 0)-(Ball.Left, Ball.Top), RGB(180, 0, 0)
  Delay 1E-19
  Area.Cls
 Next i
 Area.DrawWidth = 12
 Area.Line (0, 0)-(Ball.Left, Ball.Top), RGB(180, 0, 0)
 Area.DrawWidth = 6
 Area.Line (0, 0)-(Ball.Left, Ball.Top), RGB(255, 0, 0)
 Area.DrawWidth = 3
 Area.Line (0, 0)-(Ball.Left, Ball.Top), vbWhite
 Delay 0.2
 Area.Cls
 StopSounds True
 BlowUpBall
 Ball.Visible = False
 CentreBall
 GameStory = "The Comet Has Been Safely Destroyed!"
 Sleep 1.6
 GameStory = ""
 Select Case LivesLeft
 Case 2
  GameStory = "You've Saved The Planet neptune!"
  Score = Str(Int(Val(Score) + 5000))
  AimTime = "00:00:45"
  Sleep 1.9
  GameStory = "Your next Mission Is To Save Mars!" & Chr(13) & "The Target Time Is 0.45 Minuets."
  Sleep 3
  LifeTxt(3).caption = "OK"
  ExitBounce = False
 Case 1
  GameStory = "You've Saved Mars!!"
  Score = Str(Int(Val(Score) + 10000))
  AimTime = "00:01:30"
  Sleep 1.9
  GameStory = "Your next Mission Is To Save Earth!" & Chr(13) & "The Target Time Is 1.30 Minuets."
  Sleep 3
  LifeTxt(2).caption = "OK"
  ExitBounce = False
 Case 0
  GameStory = "You've Saved The Planet Earth!!!"
  Score = Str(Int(Val(Score) + 20000))
  LifeTxt(1).caption = "OK"
  Sleep 3
  ExitBounce = True
  ExitBounce = True
  ExitBounce = True
  ShowStats
  Exit Sub
 End Select
 Ball.Visible = True
 Area.Refresh
 GameStory = "Ready"
 Delay 0.8
 GameStory = "GO!"
 Delay 0.6
 GameStory = ""
 XSpeed = CmdSpeedParam
 YSpeed = CmdSpeedParam
 ExitBounce = False
 ExitBounce = False
 Bounce
End Sub
Private Sub TitleScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ExitBounce = True
 UnHighlight Story
 UnHighlight TrainingMode
 UnHighlight Start
 UnHighlight Quit
End Sub
Public Sub RefreshForm()
 Me.Show
 Me.Refresh
 Area.Refresh
 TitleScreen.Refresh
 Info.Refresh
End Sub
Private Sub Area_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Reset.Tag = "yes" Then UnHighlight Reset
 TPad.Left = Int(X - (TPad.Width / 2))
 BPad.Left = TPad.Left
 RPad.Top = Int(Y - (RPad.Height / 2))
 LPad.Top = RPad.Top
End Sub
Sub Form_Load()
 Load LoadUp
 LoadUp.Visible = True
 ExitBounce = True
 Area.Visible = False
 LivesLeft = 3
 XSpeed = CmdSpeedParam
 YSpeed = CmdSpeedParam
 LoadScore
 LoadTitlePics
 LoadGamePics
 LoadScrollText
 Unload LoadUp
 Call SetWindowRgn(KillBall.hWnd, CreateEllipticRgn(0, 0, Ball.Width / Screen.TwipsPerPixelX, Ball.Height / Screen.TwipsPerPixelY), True)
 Show
 Scroll.StartScroll
 RefreshForm
End Sub
Public Sub LoadScrollText()
LoadPercent = LoadPercent + 1
LoadUp.caption = LoadPercent
LoadUp.CurrLoad.caption = "Loading Text..."
LoadUp.Refresh
Sleep (0.3)
Scroll.Font.Name = Terminal
Scroll.Font.Size = 14
Scroll.Font.Bold = True
Scroll.caption = "< Q U A D - B A L L > - < P r o g r a m m i n g ,   G r a p h i c s ,   S o u n d   E f f e c t s   B y   A r v i n d e r   S e h m i > - < A r v i n d e r @ S e h m i . C o . U k > - "
End Sub
Private Sub LoadGamePics()
 If GamePicsLoaded = True Then Exit Sub
 Call LoadPic(Area, "Space.Img")
 Call LoadPic(Info, "QBallI.Img")
 Call LoadPic(AniBall, "Ball.Img")
 Call LoadPic(TPad, "TBPad.Img")
 Call LoadPic(BPad, "TBPad.Img")
 Call LoadPic(LPad, "LRPad.Img")
 Call LoadPic(RPad, "LRPad.Img")
 Call LoadPic(BallBlow, "BallBlow.Img")
 Call LoadPic(Life(1), "Planet1.Img")
 Call LoadPic(Life(2), "Planet2.Img")
 Call LoadPic(Life(3), "Planet3.Img")
 KillEarth.Picture = Area.Picture
 Call LoadAniPic(Ball, AniBall, 0)
 Call LoadAniPic(KillBall, AniBall, "0")
 GamePicsLoaded = True
End Sub
Private Sub LoadTitlePics()
 If TitlePicsLoaded = True Then Exit Sub
 Call LoadPic(StoryScreen, "QBallS.Img")
 Call LoadPic(TitleScreen, "QBallT.Img")
 Call LoadPic(StatBox, "QBallS.Img")
 UnHighlight Start
 UnHighlight Quit
 UnHighlight Story
 TitlePicsLoaded = True
End Sub
Sub LoadStoryString()
Dim NewLine As String
NewLine = Chr(13) & Chr(13)
StoryLabel.FontBold = True
StoryLabel.caption = NewLine & "The Year is 2500. " & _
 "The Human Civilisation Has Expanded And now;" & _
 " Earth, Mars And neptune Are All Inhabited by Our" & _
 " Species." & NewLine & "The Colonists On The Planet neptune Have Discovered" & _
 " Fifteen Large Comets Heading Towards Our Solar System. Three Of These are" & _
 " Heading For The Civilised Planets, These Planets" & _
 " Will be Destroyed If The Comets Are not Destroyed First." & NewLine & _
 " You Have Been Assigned To Take Control Of A new Protection" & _
 " System. You Will Have Keep The Comets In Their Current Area Location (One At a Time)" & _
 " So That There Will Be Enough Time To Destroy Each Of The Dangerous Three Saftly." & _
 NewLine & "The Human Civilisation Is Depending On YOU!"
End Sub
Public Sub SpinBall(Direction)
Static BallCell As Single
 Select Case Direction
  Case Clock
   BallCell = BallCell + 0.5
   If Int(BallCell) = 8 Then BallCell = 0
   Call LoadAniPic(Ball, AniBall, Int(BallCell))
  Case AntiClock
   BallCell = BallCell - 0.5
   If Int(BallCell) = -1 Then BallCell = 7
   Call LoadAniPic(Ball, AniBall, Int(BallCell))
 End Select
End Sub
Public Sub Bounce()
Static BallXCent As Integer, BallYCent As Integer
Static Bounces As Integer
Call GetSpeed
If ExitBounce = True Then GoTo Nd2:
Ball.Visible = True
Call SetWindowRgn(Ball.hWnd, CreateEllipticRgn(0, 0, (Ball.Width / Screen.TwipsPerPixelX), (Ball.Height / Screen.TwipsPerPixelY)), True)
StartTime = Time
Do
DoEvents
KeepMouseOnForm
Ball.Left = Ball.Left + XSpeed
Ball.Top = Ball.Top + YSpeed
BallXCent = Int(Ball.Left + (Ball.Width / 2))
BallYCent = Int(Ball.Top + (Ball.Height / 2))
 If Ball.Top <= Limit.Top And YSpeed < 0 Then
  TPadLeft = TPad.Left
  TPadRight = Int(TPad.Left + TPad.Width)
  If BallXCent > TPadLeft And BallXCent < TPadRight Then
   Call WAVPlay("Hit.qbs")
   YSpeed = -YSpeed
   AddScore
   GoTo nd:
  Else
   Die
  End If
 ElseIf Ball.Left <= Limit.Left And XSpeed < 0 Then
  LPadTop = LPad.Top
  LPadBottom = Int(LPad.Top + LPad.Height)
  If BallYCent > LPadTop And BallYCent < LPadBottom Then
   Call WAVPlay("Hit.qbs")
   XSpeed = -XSpeed
   AddScore
   GoTo nd:
  Else
   Die
  End If
 ElseIf Int(Ball.Left + Ball.Width) >= Int(Limit.Left + Limit.Width) And XSpeed > 0 Then
  RPadTop = RPad.Top
  RPadBottom = Int(RPad.Top + RPad.Height)
  If BallYCent > RPadTop And BallYCent < RPadBottom Then
   WAVPlay ("Hit.qbs")
   XSpeed = -XSpeed
   AddScore
   GoTo nd:
  Else
   Die
  End If
 ElseIf Int(Ball.Top + Ball.Height) >= Int(Limit.Top + Limit.Height) And YSpeed > 0 Then
  BPadLeft = BPad.Left
  BPadRight = Int(BPad.Left + BPad.Width)
  If BallXCent > BPadLeft And BallXCent < BPadRight Then
   Call WAVPlay("Hit.qbs")
   YSpeed = -YSpeed
   AddScore
   GoTo nd:
  Else
   Die
  End If
 End If
nd:
If Bounces = 15 Then
 KeepMouseOnForm
 If XSpeed > 0 Then SpinBall Clock Else SpinBall AntiClock
 Bounces = 0
 Dim TempTime As Date
 TempTime = Time - StartTime
 TimeSoFar = TempTime
 If TimeSoFar = AimTime Then SavePlanet
End If
Bounces = Bounces + 1
Loop Until ExitBounce = True
Nd2:
StopSounds False, True
Ball.Visible = False
TimeSoFar = "00:00:00"
End Sub
Public Sub Die()
 ExitBounce = True
 LivesLeft = LivesLeft - 1
If LivesLeft > 0 Then BlowUpBall
 Select Case LivesLeft
  Case 2
   GameStory = "The Planet neptune Has Been Destroyed!" & Chr(13) & "Four Billion People Dead!"
   Area.Refresh
   Sleep 2.5
   GameStory = "Your next Mission Is To Save Mars!" & Chr(13) & "The Target Time Is 0.45 Minuets."
   Sleep 3
   GameStory = "Ready"
   Delay 0.8
   GameStory = "GO!"
   Delay 0.6
   GameStory = ""
   AimTime = "00:00:45"
   LifeTxt(3).caption = "X"
  Case 1
   GameStory = "The Planet Mars Has Been Destroyed!" & Chr(13) & "Six Billion People Dead!"
   Area.Refresh
   Sleep 2.5
   GameStory = "Your next Mission Is To Save Earth!" & Chr(13) & "The Target Time Is 1.30 Minuets."
   Sleep 3
   GameStory = "Ready"
   Delay 0.8
   GameStory = "GO!"
   Delay 0.6
   GameStory = ""
   AimTime = "00:01:35"
   LifeTxt(2).caption = "X"
  Case 0
   LifeTxt(1).caption = "X"
   ShowDie
   Reset_Click
   Exit Sub
 End Select
 CentreBall
 YSpeed = CmdSpeedParam
 XSpeed = CmdSpeedParam
 GetSpeed
 ExitBounce = False
 Bounce
End Sub
Sub ShowDie()
'centre earth and place the killer ball to the top
 KillBall.Top = 200
 KillBall.Left = 0
 BlowUpEarth.Left = Int((Me.Width - BlowUpEarth.Width) / 2)
 BlowUpEarth.Top = BlowUpEarth.Left 'Int((Me.Height - BlowUpEarth.Height) / 2)
'hide all pic boxes and show only the animation one
 KillEarth.Visible = True
 Area.Visible = False
 TitleScreen.Visible = False
 Scroll.ExitScroll
 StoryScreen.Visible = False
' set earth round
' kill earth
  BlowUpEarth.SetShape True, False
 Dim Pos As Integer
 Dim Moves As Integer
 Static BallCell As Single
 With KillBall
  For Pos = 0 To (BlowUpEarth.Left - (.Width / 2)) + 250 Step 24
   .Top = .Top + 24
   .Left = .Top - 100
   '.Refresh
   KillEarth.Refresh
   BallCell = BallCell + 0.5
   If Int(BallCell) = 8 Then BallCell = 0
   Call LoadAniPic(KillBall, AniBall, Int(BallCell))
  Next
  BlowUpEarth.SetShape False, True
  Call WAVLoop("blowup.qbs")
  BlowUpEarth.Animate
  Call StopSounds(True, False)
 End With
KillBall.Visible = False
BlowUpEarth.Visible = False
GameStory2 = "The Planet Earth Has Been Destroyed!" & Chr(13) & "Eight Billion People Dead!"
Area.Refresh
Sleep 2.5
GameStory2 = ""
'show all other picboxes
 KillEarth.Visible = False
 Area.Visible = True
 StoryScreen.Visible = False
 TitleScreen.Visible = True
 ShowStats
End Sub
Public Sub BlowUpBall()
 Ball.Visible = True
 Dim i As Integer
 WAVPlay "blowup.qbs"
 Area.Cls
 Area.Refresh
 For i = 0 To (BallBlow.Rows * BallBlow.Cols) - 1
  Call LoadAniPic(Ball, BallBlow, i)
  Ball.Refresh
  Call Sleep(0.1)
 Next i
 CentreBall
 Call LoadAniPic(Ball, AniBall, 0)
End Sub
Sub AddScore()
On Error GoTo ScoreTooHigh
 Call IncSpeed
 Call GetSpeed
 Dim AveSpeed As Integer
 AveSpeed = Int(XSpeed + YSpeed) / 2
 If AveSpeed < 0 Then AveSpeed = -AveSpeed
 Score = CLng(Val(Score) + CLng(5 * AveSpeed))
 If Val(Score.caption) > 99999 Then Score.FontSize = 16
 NumBounces = NumBounces + 1
 Exit Sub
ScoreTooHigh:
 Score.FontSize = 14
 Score = "Maxed Out"
End Sub
Public Sub GetSpeed()
 Dim TempX, TempY
 TempX = XSpeed
 TempY = YSpeed
 If TempX < 0 Then TempX = -TempX
 If TempY < 0 Then TempY = -TempY
 CommetSpeed = Val(Int((TempX + TempY) / 2))
 If Val(CommetSpeed) > FastSpeed Then FastSpeed = Val(CommetSpeed)
End Sub
Private Sub KeepMouseOnForm()
If ExitMainMouse = True Then GoTo nd:
If Get_Mouse_X >= Int((Me.Left + Me.Width) _
 / Screen.TwipsPerPixelX) - 5 Then _
 Set_Mouse_X ((Me.Left + Me.Width) / Screen.TwipsPerPixelX) - 5
If Get_Mouse_X <= Int((Me.Left) _
 / Screen.TwipsPerPixelX) Then _
 Set_Mouse_X ((Me.Left) / Screen.TwipsPerPixelX) + 5
If Get_Mouse_Y <= Int(Me.Top / Screen.TwipsPerPixelY) + 5 Then _
 Set_Mouse_Y (Me.Top / Screen.TwipsPerPixelY) + 35
If Get_Mouse_Y >= Int((Me.Top + Me.Height) / Screen.TwipsPerPixelY) - 5 Then _
 Set_Mouse_Y ((Me.Top + Me.Height) / Screen.TwipsPerPixelY) - 5
nd:
End Sub

