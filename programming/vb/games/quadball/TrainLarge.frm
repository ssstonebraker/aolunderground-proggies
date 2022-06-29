VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form TrainLarge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quad-Ball Training Mode. Arvinder Sehmi 1999, Arvinder@Sehmi.Co.Uk"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "TrainLarge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox StatBox 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   0
      ScaleHeight     =   3030
      ScaleWidth      =   3570
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3570
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Left            =   -90
         TabIndex        =   11
         Top             =   5535
         Width           =   3885
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bounces"
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
         Left            =   135
         TabIndex        =   10
         Top             =   6705
         Width           =   3660
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
         Left            =   4185
         TabIndex        =   9
         Top             =   1755
         Width           =   2490
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
         Left            =   4185
         TabIndex        =   8
         Top             =   5400
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
         Left            =   4230
         TabIndex        =   7
         Top             =   6615
         Width           =   4185
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
         TabIndex        =   6
         Top             =   270
         Width           =   5985
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Time"
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
         Left            =   945
         TabIndex        =   5
         Top             =   3105
         Width           =   2850
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Highest Time"
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
         Left            =   315
         TabIndex        =   4
         Top             =   4320
         Width           =   3480
      End
      Begin VB.Label AllTime 
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
         Left            =   4185
         TabIndex        =   3
         Top             =   3015
         Width           =   2490
      End
      Begin VB.Label HighTime 
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
         Left            =   4185
         TabIndex        =   2
         Top             =   4230
         Width           =   2490
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
         Left            =   7605
         TabIndex        =   1
         Tag             =   "No"
         Top             =   7380
         Width           =   2805
      End
   End
   Begin VB.PictureBox TitleScreen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   6690
      TabIndex        =   13
      Top             =   0
      Width           =   6690
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
         Left            =   6795
         TabIndex        =   17
         Top             =   5220
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
         Left            =   6930
         TabIndex        =   16
         Top             =   6435
         Width           =   2955
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Quad-Ball Training Mode"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1005
         Left            =   135
         TabIndex        =   15
         Top             =   7875
         Width           =   11805
      End
      Begin VB.Label Story 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Play Story Mode"
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
         Left            =   6615
         TabIndex        =   14
         Top             =   5805
         Width           =   3675
      End
   End
   Begin VB.PictureBox Area 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H000000FF&
      Height          =   9030
      Left            =   0
      MouseIcon       =   "TrainLarge.frx":09BA
      MousePointer    =   99  'Custom
      ScaleHeight     =   9030
      ScaleWidth      =   10320
      TabIndex        =   38
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
         Tag             =   "Side"
         Top             =   45
         Width           =   330
      End
      Begin PicClip.PictureClip BallBlow 
         Left            =   6705
         Top             =   3240
         _ExtentX        =   5292
         _ExtentY        =   5080
         _Version        =   393216
         Rows            =   4
         Cols            =   4
      End
      Begin PicClip.PictureClip AniBall 
         Left            =   6975
         Top             =   1755
         _ExtentX        =   2646
         _ExtentY        =   1323
         _Version        =   393216
         Rows            =   2
         Cols            =   4
      End
      Begin VB.Shape Limit 
         BackColor       =   &H00000000&
         BorderColor     =   &H0000FF00&
         Height          =   8370
         Left            =   315
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
         TabIndex        =   44
         Top             =   3555
         Width           =   5970
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Info 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   10320
      ScaleHeight     =   8970
      ScaleWidth      =   1650
      TabIndex        =   18
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
         Left            =   360
         TabIndex        =   22
         Tag             =   "no"
         Top             =   7110
         Width           =   1005
      End
      Begin VB.Image Life 
         Height          =   375
         Index           =   2
         Left            =   630
         Top             =   2520
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF80FF&
         Height          =   690
         Left            =   90
         Top             =   3105
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
         TabIndex        =   37
         Top             =   3420
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
         Left            =   0
         TabIndex        =   36
         Top             =   3420
         Width           =   1005
      End
      Begin VB.Label CometSpeedLab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
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
         Left            =   45
         TabIndex        =   35
         Top             =   3105
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   780
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
         TabIndex        =   34
         Top             =   2070
         Width           =   1410
      End
      Begin VB.Shape PlanetsBox 
         BorderColor     =   &H00FF8080&
         Height          =   1005
         Left            =   90
         Top             =   2025
         Width           =   1410
      End
      Begin VB.Image Life 
         Height          =   375
         Index           =   3
         Left            =   180
         Top             =   2520
         Width           =   375
      End
      Begin VB.Image Life 
         Height          =   375
         Index           =   1
         Left            =   1080
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label TotalTime 
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
         TabIndex        =   33
         Top             =   1665
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
         TabIndex        =   32
         Top             =   1350
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
         TabIndex        =   31
         Top             =   945
         Width           =   1410
      End
      Begin VB.Shape TimeBox 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         Height          =   1005
         Left            =   90
         Top             =   945
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
         TabIndex        =   30
         Top             =   405
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
         TabIndex        =   29
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
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   28
         Top             =   2565
         Width           =   375
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
         Height          =   285
         Index           =   2
         Left            =   630
         TabIndex        =   27
         Top             =   2565
         Width           =   375
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
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   26
         Top             =   2565
         Width           =   375
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFF00&
         Height          =   1320
         Left            =   90
         Top             =   3870
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
         Height          =   330
         Left            =   90
         TabIndex        =   25
         Top             =   3870
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label HighestScore 
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
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   90
         TabIndex        =   24
         Top             =   4095
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
         TabIndex        =   23
         Top             =   4500
         Width           =   1455
      End
      Begin VB.Image InfoArea 
         Height          =   9075
         Left            =   -45
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   1770
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H000080FF&
         Height          =   1680
         Left            =   90
         Top             =   5265
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Longest Time Total"
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
         Left            =   90
         TabIndex        =   21
         Top             =   5310
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label HighestTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   90
         TabIndex        =   20
         Top             =   5895
         Width           =   1455
      End
      Begin VB.Label HighTimeName 
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
         TabIndex        =   19
         Top             =   6300
         Width           =   1455
      End
   End
End
Attribute VB_Name = "TrainLarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' TO SEE DOCUMENTATION ON WHAT THESE SUBS DO    '
' GOTO "TrainFRM", As This Code Is Very Similar '
' And As Only Been Slightly Adjusted To Work    '
' On Displays Larger Than 800x600               '
'_______________________________________________'
Dim ExitBounce As Boolean
Dim StartTime As Date
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight Story
 UnHighlight Start
 UnHighlight Quit
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Call StopSounds(True, True)
 ExitBounce = True
 Call StopSounds(True, True)
 Reset_Click
 StatBox.Visible = False
 TitleScreen.Visible = True
 Reset_Click
 Unload Me
 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call StopSounds(True, True)
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
 LoadScoreTraining
 LoadTimeTraining
 CentreBall
 TitleScreen.Visible = True
 Area.Visible = False
 LifeTxt(1).Caption = ""
 LifeTxt(2).Caption = ""
 LifeTxt(3).Caption = ""
 LivesLeft = 3
 XSpeed = CmdSpeedParam
 YSpeed = CmdSpeedParam
 For i = 1 To 3
  Life(1).Visible = True
 Next i
 TopTime = "00:00:00"
 TimeSoFar = "00:00:00"
 OldTime = "00:00:00"
 TotalTime = "00:00:00"
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
 LoadScoreTraining
 LoadTimeTraining
 LivesLeft = 3
 XSpeed = CmdSpeedParam
 YSpeed = CmdSpeedParam
  For i = 1 To 3
   Life(i).Visible = True
   LifeTxt(i).Caption = ""
  Next i
 TopTime = "00:00:00"
 TimeSoFar = "00:00:00"
 OldTime = "00:00:00"
 TotalTime = "00:00:00"
 Score = "0"
 Ball.Visible = False
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
 UnHighlight Quit
 UnHighlight Story
 Highlight Start
End Sub
Private Sub Quit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight Start
 UnHighlight Story
 Highlight Quit
End Sub
Private Sub ShowStats()
TotalScore = Val(Score)
TopSpeed = FastSpeed
TotalBounces = NumBounces
HighTime.Caption = TopTime
AllTime.Caption = TotalTime
StatBox.Visible = True
 If TotalScore > Val(GetKeyValue(HKEY_LOCAL_MACHINE, _
  "SOFTWARE\arvisehmi\QuadBall\Training", "TopScore")) Then
  On Error Resume Next
  InputWindow.Visible = True
 End If
Dim TmpOldHigh As String
Dim OldHigh As Date
  TmpOldHigh = (GetKeyValue(HKEY_LOCAL_MACHINE, _
  "SOFTWARE\arvisehmi\QuadBall\Training", "TopTime"))
  OldHigh = TmpOldHigh
  If TotalTime > OldHigh Then
   On Error Resume Next
   InputWindow2.Visible = True
  End If
Reset_Click
End Sub
Private Sub StatBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight OK
End Sub
Private Sub Story_Click()
 On Error GoTo a:
 ThisDir
 Shell "QuadBall.Exe " & Trim(Str(CmdSpeedParam)), vbNormalFocus
 Reset_Click
 Unload Me
 End
a:
 MsgBox "Cannot Find QuadBall.Exe," & Chr(13) & "Please Re-install This Game To Fix The Problem." & Chr(13) & _
        "If You Can Find QuadBall.exe On Your PC," & Chr(13) & "Please Place It In The Directory:" & Chr(13) & _
        App.Path, vbCritical, "Error"
End Sub
Private Sub totalbounces_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight OK
End Sub
Private Sub label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UnHighlight OK
End Sub
Private Sub Story_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Highlight Story
 UnHighlight Start
 UnHighlight Quit
End Sub
Private Sub TitleScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ExitBounce = True
 UnHighlight Start
 UnHighlight Quit
End Sub
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' TO SEE DOCUMENTATION ON WHAT THESE SUBS DO    '
' GOTO "TrainFRM", As This Code Is Very Similar '
' And As Only Been Slightly Adjusted To Work    '
' On Displays Larger Than 800x600               '
'_______________________________________________'
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
 Load TrainLoadUp
 TrainLoadUp.Visible = True
 ExitBounce = True
 Area.Visible = False
 LivesLeft = 3
 XSpeed = CmdSpeedParam
 YSpeed = CmdSpeedParam
 LoadScoreTraining
 LoadTimeTraining
 LoadTitlePics
 LoadGamePics
 Unload TrainLoadUp
 Show
 RefreshForm
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
 Call LoadAniPic(Ball, AniBall, 0)
 GamePicsLoaded = True
End Sub
Private Sub LoadTitlePics()
 If TitlePicsLoaded = True Then Exit Sub
 Call LoadPic(TitleScreen, "QBallT.Img")
 Call LoadPic(StatBox, "QBallS.Img")
 UnHighlight Start
 UnHighlight Quit
 TitlePicsLoaded = True
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

'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' Control The Mouse'
'__________________'
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
Public Sub Bounce()
Static BallXCent As Integer, BallYCent As Integer
Static Bounces As Integer
GetSpeed

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
 TotalTime = OldTime + TempTime
End If
Bounces = Bounces + 1
Loop Until ExitBounce = True
Nd2:
StopSounds True, True
Ball.Visible = False
'TimeSoFar = "00:00:00"
End Sub
Public Sub Die()
 ExitBounce = True
 LivesLeft = LivesLeft - 1
 BlowUpBall
 Dim TempTime As Date
 TempTime = TimeSoFar
 OldTime = OldTime + TempTime
 Debug.Print LivesLeft; "  "; OldTime
 Select Case LivesLeft
  Case 2
   GameStory = "Ready"
   Delay 0.8
   GameStory = "GO!"
   Delay 0.6
   GameStory = ""
   LifeTxt(3).Caption = "X"
   If TempTime > TopTime Then TopTime = TempTime
  Case 1
   GameStory = "Ready"
   Delay 0.8
   GameStory = "GO!"
   Delay 0.6
   GameStory = ""
   LifeTxt(2).Caption = "X"
   If TempTime > TopTime Then TopTime = TempTime
  Case 0
   LifeTxt(1).Caption = "X"
   '+-----------+
    ' code here '
   '+-----------+
   'Reset_Click
   If TempTime > TopTime Then TopTime = TempTime
   ShowStats
   Exit Sub
 End Select
 CentreBall
 YSpeed = CmdSpeedParam
 XSpeed = CmdSpeedParam
 GetSpeed
 ExitBounce = False
 TimeSoFar = "00:00:00"
 Bounce
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
 If Val(Score.Caption) > 99999 Then Score.FontSize = 16
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
