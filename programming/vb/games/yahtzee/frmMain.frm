VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahtzee Deluxe"
   ClientHeight    =   4860
   ClientLeft      =   5565
   ClientTop       =   2760
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4800
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      Height          =   400
      Index           =   4
      Left            =   2160
      TabIndex        =   49
      Top             =   4115
      Width           =   2535
      Begin VB.Label lblScores 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus Yahtzee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   13
         Left            =   120
         TabIndex        =   51
         Top             =   130
         Width           =   1920
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   13
         Left            =   2055
         TabIndex        =   50
         Top             =   130
         Width           =   150
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   3255
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   225
         Left            =   75
         TabIndex        =   52
         Top             =   210
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Max             =   13
         Scrolling       =   1
      End
      Begin VB.Label lblRollNumber 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   2970
         TabIndex        =   48
         Top             =   225
         Width           =   120
      End
      Begin VB.Label lblStatic 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll #"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   25
         Top             =   225
         Width           =   525
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   3
      Left            =   2160
      TabIndex        =   7
      Top             =   1830
      Width           =   2535
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   12
         Left            =   2050
         TabIndex        =   42
         Top             =   2025
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   11
         Left            =   2050
         TabIndex        =   41
         Top             =   1710
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   10
         Left            =   2050
         TabIndex        =   40
         Top             =   1410
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   2050
         TabIndex        =   39
         Top             =   1095
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   8
         Left            =   2050
         TabIndex        =   38
         Top             =   780
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   7
         Left            =   2050
         TabIndex        =   37
         Top             =   480
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   2050
         TabIndex        =   36
         Top             =   165
         Width           =   150
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Chance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   120
         TabIndex        =   24
         Top             =   2025
         Width           =   1875
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Yahtzee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   11
         Left            =   120
         TabIndex        =   23
         Top             =   1725
         Width           =   1935
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Large Straight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   120
         TabIndex        =   22
         Top             =   1410
         Width           =   1755
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Small Straight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   21
         Top             =   1095
         Width           =   1935
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Full House"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   795
         Width           =   1845
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "4 of a kind"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "3 of a kind"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   165
         Width           =   1800
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      Begin MSComctlLib.ImageList ilDice 
         Left            =   690
         Top             =   465
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":07EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0D02
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":124E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":17CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1DB6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1530
         Top             =   465
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   73
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":23A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2726
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2AA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D32
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2FE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":35FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkDice 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4000
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   195
      End
      Begin VB.CheckBox chkDice 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3090
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   195
      End
      Begin VB.CheckBox chkDice 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2180
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   195
      End
      Begin VB.CheckBox chkDice 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1270
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   195
      End
      Begin VB.CheckBox chkDice 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   195
      End
      Begin VB.Label labClick 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   4
         Left            =   3800
         TabIndex        =   47
         Top             =   120
         Width           =   615
      End
      Begin VB.Label labClick 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   3
         Left            =   2880
         TabIndex        =   46
         Top             =   120
         Width           =   615
      End
      Begin VB.Label labClick 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   2
         Left            =   1950
         TabIndex        =   45
         Top             =   120
         Width           =   615
      End
      Begin VB.Label labClick 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   1
         Left            =   1080
         TabIndex        =   44
         Top             =   120
         Width           =   615
      End
      Begin VB.Label labClick 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   615
      End
      Begin VB.Image imgDice 
         Height          =   495
         Index           =   1
         Left            =   1120
         Top             =   190
         Width           =   495
      End
      Begin VB.Image imgDice 
         Height          =   495
         Index           =   2
         Left            =   2030
         Top             =   190
         Width           =   495
      End
      Begin VB.Image imgDice 
         Height          =   495
         Index           =   3
         Left            =   2940
         Top             =   190
         Width           =   495
      End
      Begin VB.Image imgDice 
         Height          =   495
         Index           =   4
         Left            =   3850
         Top             =   195
         Width           =   495
      End
      Begin VB.Image imgDice 
         Height          =   495
         Index           =   0
         Left            =   210
         Top             =   195
         Width           =   495
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   2
      Left            =   135
      TabIndex        =   5
      Top             =   1830
      Width           =   1935
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   1560
         TabIndex        =   35
         Top             =   1730
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   1560
         TabIndex        =   34
         Top             =   1416
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   1560
         TabIndex        =   33
         Top             =   1102
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   1560
         TabIndex        =   32
         Top             =   788
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1560
         TabIndex        =   27
         Top             =   480
         Width           =   150
      End
      Begin VB.Label lblNums 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   1560
         TabIndex        =   26
         Top             =   165
         Width           =   150
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Six"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1725
         Width           =   1290
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Five"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1410
         Width           =   1425
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Four"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1095
         Width           =   1440
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Three"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   795
         Width           =   1335
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Two"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label lblScores 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   180
         Width           =   1260
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   4800
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Image imgRoll 
      Height          =   330
      Left            =   3520
      Picture         =   "frmMain.frx":3A7E
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label lblGTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4225
      TabIndex        =   31
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label lblBonusCountDown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-63"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Left            =   1005
      TabIndex        =   30
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label lblLScoreTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1680
      TabIndex        =   29
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label lblScoreBonus 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1680
      TabIndex        =   28
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label lblStatic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Game Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   3
      Left            =   2280
      TabIndex        =   11
      Top             =   4560
      Width           =   1230
   End
   Begin VB.Label lblStatic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus (     )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   1170
   End
   Begin VB.Label lblStatic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   0
      X2              =   4800
      Y1              =   45
      Y2              =   45
   End
   Begin VB.Menu mnuGameMain 
      Caption         =   "&Game"
      Begin VB.Menu mnuGame 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Undo Last"
         Index           =   2
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Statistics"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&High Score"
         Index           =   5
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuGame 
         Caption         =   "E&xit"
         Index           =   7
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelpMain 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help &Contents"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help &Index"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'*  Name:    Yahtzee Deluxe (couldn't think of a better name)     *
'*  Author:  Shannon Harmon                                       *
'*  Email:   sharmon@microtechcomputers.com                       *
'*  Date:    March 3, 1999                                        *
'*  Updated: September 25, 1999                                   *
'*                                                                *
'*  Source Copyright - Shannon Harmon                             *
'*  Yahtzee is owned by Milton Bradley                            *
'*                                                                *
'*  Notice:  Made available only for personal (non commercial)    *
'*           use.  Do NOT try to sell this software, use only for *
'*           learning and ideas.  Source code made free available *
'*           to the public domain.  Please email me any updates   *
'*           you make and please keep this header with any source *
'*           you distribute.                                      *
'*                                                                *
'*  PS:  Help file class from http://www.vbexplorer.com           *
'*                                                                *
'******************************************************************

Option Explicit

Dim RollCount As Integer '//Track current round total rolls (max 3)
Dim die(4) As Integer '//Tracks dices number value
Dim uIndex As Integer '//Tracks last score clicked
Dim uGameTotal As Integer '//Tracks last total score before undo
Dim uBonus63 As Integer '//Tracks last bonus amount
Dim ulColTotal As Integer '//Tracks last left column total points
Dim uBonusCountDown As Integer '//Tracks last bonus countdown number
Dim uYatzeeBonus As Boolean '//Tracks if bonus yahtzee should test if undo mode
Dim uRollCount As Integer '//Tracks rollcount before undo
Private hHelp As New HTMLHelp '//Help class object
'

Private Sub Form_Load()
  
  Call CheckReg '//Check and setup registry
  Call NewGame '//Start new game
  
End Sub

Private Sub NewGame() '//Initialze new game
Dim i As Integer
    
  For i = 0 To 13 '//Disable and reset all scoreing buttons
    
    lblNums(i).Enabled = False
    lblNums(i) = "X"
    lblNums(i).FontStrikethru = False
    lblScores(i).Enabled = False
    lblScores(i).FontStrikethru = False
  
  Next i

  Bonus63 = 0 '//Reset score variables
  lColTotal = 0
  rColTotal = 0
  GameTotal = 0
    
  lblScoreBonus = "0" '//Left Column bonus score
  lblLScoreTotal = "0" '//Left Column Score total
  lblGTotal = "0" '//Game total
  lblBonusCountDown = "-63" '//Amount left till bonus available
  pBar.Value = 0 '//Total Roll Progress Bar
  lblRollNumber = "0" '//Visible Text Roll Number Counter
  RollCount = 0 '//Actual Roll Counter Variable
  imgRoll.Enabled = True '//Enable clicking of the Roll Button
  imgRoll.Picture = ImageList1.ListImages(2).Picture '//Set picture to normal
    
  For i = 0 To 4
    
    imgDice(i).Picture = ilDice.ListImages(1).Picture '//Set dice pics to #1
    chkDice(i).Enabled = False '//Disable dice checkboxes
    chkDice(i).Value = 0 '//Set dice checkboxes to greyed
    labClick(i).Enabled = False '//Disable top labels over dice
  
  Next i
  
  '//Undo variables
  mnuGame(2).Enabled = False '//Disable undo menu item
  uGameTotal = 0
  uBonusCountDown = 0
  uBonus63 = 0
  ulColTotal = 0
  uYatzeeBonus = False
    
End Sub

Private Sub Roll() '//Roll Dice
Dim i As Integer, j As Integer
    
  If pBar.Value = 0 Then '//If first roll of game enable all score values

    For i = 0 To 13

      lblNums(i).Enabled = True
      lblScores(i).Enabled = True

    Next i

  End If
        
  Randomize  '//Initializes the random-number generator.
  
  For i = 0 To 4
    
    j = (Int(6 * Rnd)) '//Set j to random number between 1 and 6
    chkDice(i).Enabled = True '//Enable dice checkboxes
    
    '//Set enabled checkbox dice pictures to correct number
    '//If it's 0 we know it's not been checked
    If chkDice(i).Value = 0 Then imgDice(i).Picture = ilDice.ListImages(j + 1).Picture
    
    '//Set our variable so we know what the last roll dice value was...
    '//If it's 0 we know it's not been checked
    If chkDice(i).Value = 0 Then die(i) = j + 1
  
  Next i
    
  If RollCount = 3 Then '//Turn is over
    
    imgRoll.Enabled = False '//Disable Roll button
    imgRoll.Picture = ImageList1.ListImages(7).Picture '//Change to disabled pic
    
    For i = 0 To 4
      
      chkDice(i).Value = 2 '//Set all checkboxes to selected
      chkDice(i).Enabled = False '//Disable all checkboxes
      labClick(i).Enabled = False '//Disable all dice covering labels
    
    Next i
    
    RollCount = 0 '//Set our roll count back to zero
  
  End If
    
  Call CheckAllScores
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  hHelp.HHClose '//Close our help class
  Set hHelp = Nothing '//Give back it's memory

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

  '//Close all sub forms if any open
  For i = Forms.Count - 1 To 1 Step -1
    Unload Forms(i)
  Next
  
  '//Save screen position
  If Me.WindowState <> vbMinimized Then
        
    SaveSetting MyApp, "Settings", "MainLeft", Me.Left '//Forms left position
    SaveSetting MyApp, "Settings", "MainTop", Me.Top '//Forms top position
  
  End If

End Sub

Private Sub imgRoll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Button = 2 Then Exit Sub '//If it's the right mouse button just exit
  
  If mnuSound.Checked Then PlaySound App.Path & "\roll.wav" '//Play sound if available
  
  imgRoll.Picture = ImageList1.ListImages(1).Picture '//Change pic to down picture
  
End Sub

Private Sub imgRoll_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Button = 2 Then Exit Sub '//If it's the right mouse button just exit
  
  imgRoll.Picture = ImageList1.ListImages(2).Picture '//Change pic to up picture
  
End Sub

Private Sub imgRoll_Click()
Dim i As Integer

  mnuGame(2).Enabled = False '//Disable undo menu item
  
  If RollCount = 0 Then '//First roll of turn
    
    For i = 0 To 4
      
      chkDice(i).Value = 0 '//Reset checkboxes to zero
      chkDice(i).Enabled = True '//Allow clicking of checkboxes
      labClick(i).Enabled = True '//Allow clicking of dice label covers
    
    Next
  
  End If
  
  RollCount = RollCount + 1 '//Update our roll count
  lblRollNumber = RollCount '//Update our visible roll count
   
  Call Roll '//Do the dice roll subroutine
   
End Sub

Private Sub labClick_Click(Index As Integer)

  '//Set value to opposite
  If chkDice(Index).Value = 0 Then chkDice(Index).Value = 1: Exit Sub
  
  If chkDice(Index).Value = 1 Then chkDice(Index).Value = 0

End Sub

Private Sub lblNums_Click(Index As Integer)
  
  Call lblScores_Click(Index) '//Do same thing as lblscores_click

End Sub

Private Sub UndoPlay()
Dim i As Integer
  
  mnuGame(2).Enabled = False '//Disable undo menu item
  
  If uIndex = 11 Then
  
    lblNums(13).Enabled = True '//Set bonus yahtzee to available again
    lblNums(13) = "X"
    lblScores(13).Enabled = True
  
  End If
  
  '//Reset item score number
  With lblNums(uIndex)
    
    .Caption = "X"
    .FontStrikethru = False
    .Enabled = True
  
  End With
  
  '//Reset item text
  With lblScores(uIndex)
    
    .FontStrikethru = False
    .Enabled = True
  
  End With
  
  If pBar.Value <> 0 Then pBar.Value = pBar.Value - 1 '//Move pBar.Value back
  
  lblGTotal = uGameTotal '//Move last score back
  lblBonusCountDown = uBonusCountDown '//Set bonus countdown amount back
  lblScoreBonus = uBonus63 '//Set bonus back to last state
  lblLScoreTotal = ulColTotal '//Set left column total back
  uYatzeeBonus = True '//Let's check all scores no we are calling it from an undo
  
  If uRollCount <> 3 And uRollCount <> 0 Then '//Allows for roll again if not last turn
    
    RollCount = uRollCount '//Restores old roll number
    lblRollNumber = RollCount '//Makes roll number visible to user
    imgRoll.Enabled = True '//Enable Roll button incase it was off
    imgRoll.Picture = ImageList1.ListImages(2).Picture '//Reset picture to default
    
    For i = 0 To 4
    
      chkDice(i).Enabled = True '//Allow clicking of checkboxes
      labClick(i).Enabled = True '//Allow clicking of dice label covers
    
    Next i
    
  Else '//Turn was last so keep roll button disabled
  
    imgRoll.Enabled = False '//Disable Roll button incase it was off
    imgRoll.Picture = ImageList1.ListImages(7).Picture '//Reset picture to done
  End If
  
  Call CheckAllScores '//Update available scoring positions
  
End Sub

Private Sub lblScores_Click(Index As Integer)
Dim i As Integer


  If Index = 13 Then Exit Sub '//Bonus yahtzee - not a user clickable score
  
  If lblNums(Index) = "X" Then Exit Sub '//Exit if we clicked an unusable one
  
  '//Undo routine items
  uIndex = Index '//Track clicked item for undo routine
  mnuGame(2).Enabled = True '//Enable undo menu item
  uGameTotal = lblGTotal '//Tracks last game total
  uBonusCountDown = lblBonusCountDown '//Tracks bonus countdown amount
  uBonus63 = lblScoreBonus '//Tracks bonus amount
  ulColTotal = lblLScoreTotal '//Tracks left score total
  uYatzeeBonus = False '//Resets undo mode tracker to false
  uRollCount = RollCount '//Tracks roll number
  
  For i = 0 To 4
    
    chkDice(i).Enabled = False '//Not allow clicking of checkboxes
    labClick(i).Enabled = False '//Not allow clicking of dice label covers

  Next i
  
  lblScores(Index).Enabled = False '//Disable this item
  lblScores(Index).FontStrikethru = True '//Add font strikethrou so we know it's used
  lblNums(Index).Enabled = False '//Disable this item
  lblNums(Index).FontStrikethru = True '//Add font strikethrou so we know it's used
     
  lblLScoreTotal = 0 '//Reset left scoring column to zero
  rColTotal = 0 '//Reset left scoring column to zero
  
  '//Process for all score clicks
  
  '//---Left columns
  For i = 0 To 5
    
    If lblNums(i).Enabled = False Then '//Only used disabled items for scoring
      
      lblLScoreTotal = CStr(Int(lblLScoreTotal) + Int(lblNums(i))) '//Show left column score
      lColTotal = Int(lblLScoreTotal) '//Total left column score
    
    End If
    
    If lblNums(i).Enabled = True Then '//It's not been used, set it back to normal
      
      lblNums(i) = "X" '//This isn't used yet so put value back to "X"
    
    End If
    
    lblNums(i).ForeColor = vbBlack '//Reset scoreing colors back to black
  
  Next i
  
  If lColTotal >= 63 Then '//Bonus available
    
    lblScoreBonus = 35 '//They got their bonus, set the value
    Bonus63 = 35 '//Variable to hold the bonus amount
    lblBonusCountDown = 0 '//Show 0 since they got the bonus
  
  Else '//No bonus available
    
    lblScoreBonus = 0 '//No bonus yet, keep at zero
    Bonus63 = 0 '//Variable to hold the bonus amount
    lblBonusCountDown = -63 + lColTotal '//Tell how much is left to go till bonus
  
  End If
  '//---End Left columns
  
  '//---Right columns
  For i = 6 To 12
    
    If lblNums(i).Enabled = False Then '//Only use disabled items to score
      
      rColTotal = rColTotal + Int(lblNums(i)) '//Add right column total
    
    End If '//If enabled set it back to default
    
    If lblNums(i).Enabled Then lblNums(i) = "X" '//Not used so put back to "X"
    
    lblNums(i).ForeColor = vbBlack '//Reset scoreing colors back to black
  
  Next i
  
  '//Add in the bonus yahtzee if there...
  If lblNums(13).Enabled = True Then
    
    If lblNums(11).Enabled = False And lblNums(11) = "0" Then '//Yahtzee is used and 0
      
      lblNums(13).Enabled = False '//Disable bonus yahtzee
      lblScores(13).Enabled = False '//Disable bonus yahtzee
      lblNums(13) = 0 '//Set bonus yahtzee score to zero
    
    End If
  
  End If
  
  '//Set bonus yahtzee score to 0 if last turn and no bonus still
  If lblNums(13) = "X" And pBar.Value = 13 Then
    
    lblNums(13) = 0 '//Set it to zero score
    lblNums(13).Enabled = False '//Disable
    lblScores(13).Enabled = False '//Disable
  
  End If
  
  If lblNums(13) <> "X" Then rColTotal = rColTotal + Int(lblNums(13)) '//Add yahtzee bonus total to score
  '//---End right columns
  
  GameTotal = (Bonus63 + lColTotal + rColTotal) '//Game score total
  lblGTotal = GameTotal '//Form visible game score total label
  
  If pBar.Value = 12 Then '//Game is over
  
    pBar.Value = pBar.Value + 1 '//Fill progress bar to max level
    
    If CheckForHS = False Then '//Sub to determine if score is in top 5
      
      frmGameOver.Show vbModal '//Score was lower then top 5 show game over form
    
    Else '//Score was in top 5
    
      If frmMain.mnuSound.Checked Then PlaySound App.Path & "\Excited.wav" '//Sound:)
      
      frmHighScore.Show vbModal '//If score was in top 5 show high score form
    
    End If
      
    Call UpdateStats(GameTotal)  '//Add this game to the statistics
    Call NewGame '//Reset game to new
    
    Exit Sub
  
  End If
  
  imgRoll.Enabled = True '//Enable Roll button incase it was off
  imgRoll.Picture = ImageList1.ListImages(2).Picture '//Reset picture to default
  lblRollNumber = 0 '//Set roll count back to zero
  RollCount = 0 '//Set our roll count back to zero
  pBar.Value = pBar.Value + 1 '//Update game progress bar

End Sub

Private Sub mnuGame_Click(Index As Integer)
  
  Select Case Index
       
    Case 0
      NewGame '//Start game over...
    
    Case 2
      Call UndoPlay '//Undo last score click...
    
    Case 4
      frmStats.Show vbModal '//Show statistics form
    
    Case 5
      frmHighScore.Show vbModal '//Show high scores form
    
    Case 7
      Unload Me '//Exit program
      End '//Just incase
    
  End Select
  
End Sub

Private Sub mnuHelp_Click(Index As Integer)
  
  Select Case Index
  
    '//Show help file
    Case 0
      
      With hHelp '//Help contents
       .CHMFile = App.Path & "\yahtzee.chm"
       .HHWindow = ""
       .HHDisplayContents
      End With
      
    Case 1
      
      With hHelp '//Help index
        .CHMFile = App.Path & "\yahtzee.chm"
        .HHWindow = ""
        .HHDisplayIndex
      End With

    Case 3
      frmAbout.Show vbModal '//Show about form
    
  End Select
  
End Sub

Private Sub mnuSound_Click() '//Toggle sound on or off
  
  mnuSound.Checked = Not mnuSound.Checked '//Set to opposite value
  SaveSetting MyApp, "Settings", "Sound", mnuSound.Checked  '//Save to registry

End Sub

Private Sub CheckAllScores()
''//Hard to comment, hope you can understand this one!
'//This routine checks to see which scores can go where!
Dim i As Integer, j As Integer
Dim tmp As Integer
Dim RoundScore As Integer
Dim iPar As Boolean, iTriss As Boolean

  '//---Do left column numbers
  For i = 0 To 5
    
    tmp = 0
  
    If lblNums(i).Enabled = True Then '//Only check scores that are enabled
      
      For j = 0 To 4
        
        If die(j) = i + 1 Then tmp = tmp + (i + 1)
        
        lblNums(i) = tmp '//Update label with correct score value
      
      Next j
    
    End If
  
  Next i
  '//---End of left column numbers

  '//---Three of a kind
  If lblNums(6).Enabled = True Then
    
    For i = 0 To 5
      
      tmp = 0
      
      For j = 0 To 4
      
        If imgDice(j).Picture = ilDice.ListImages(i + 1).Picture Then tmp = tmp + 1
      
      Next j
    
    If tmp >= 3 Then
      
      lblNums(6) = die(0) + die(1) + die(2) + die(3) + die(4)
      Exit For
    
    Else
    
    lblNums(6) = 0
    
    End If
    
    Next i
  End If
  '//---End three of a kind

  '//---Four of a kind
  If lblNums(7).Enabled = True Then
    
    For i = 0 To 5
      
      tmp = 0
    
      For j = 0 To 4
      
        If imgDice(j).Picture = ilDice.ListImages(i + 1).Picture Then tmp = tmp + 1
      
      Next j
    
      If tmp >= 4 Then
        
        lblNums(7) = die(0) + die(1) + die(2) + die(3) + die(4)
        Exit For
      
      Else
      
        lblNums(7) = 0
      
      End If
    
    Next i

  End If
  '//---End four of a kind

  '//---Full house
  If lblNums(8).Enabled = True Then
    
    iPar = False
    iTriss = False

    For i = 0 To 5
      
      tmp = 0
      
      For j = 0 To 4
      
        If imgDice(j).Picture = ilDice.ListImages(i + 1).Picture Then tmp = tmp + 1
      
      Next j
    
      If tmp = 2 Then
        
        RoundScore = RoundScore + (i + 1) * 2
        iPar = True
      
      End If
            
      If tmp = 3 Then
      
        RoundScore = RoundScore + (i + 1) * 3
        iTriss = True
      
      End If

  Next i

  RoundScore = 25
  
  If Not iPar Or Not iTriss Then RoundScore = 0
  
  lblNums(8) = RoundScore
  
  End If
  '//---End full house

  '//---Small straight
  If lblNums(9).Enabled = True Then
    
    Dim a As Integer, b As Integer, c As Integer
    Dim d As Integer, e As Integer, g As Integer
    a = 9: b = 9: c = 9: d = 9: e = 9: g = 9
    lblNums(9) = 0
  
    For i = 0 To 4
      
      If die(i) = 1 Then a = 0
      If die(i) = 2 Then b = 1
      If die(i) = 3 Then c = 2
      If die(i) = 4 Then d = 3
      If die(i) = 5 Then e = 4
      If die(i) = 6 Then g = 5
    
    Next i

    If a = 0 And b = 1 And c = 2 And d = 3 Then lblNums(9) = 30
    If b = 1 And c = 2 And d = 3 And e = 4 Then lblNums(9) = 30
    If c = 2 And d = 3 And e = 4 And g = 5 Then lblNums(9) = 30

  End If
  '//---End small straight

  '//---Large straight
  If lblNums(10).Enabled = True Then
   
    lblNums(10) = 40

    For i = 0 To 4
    
      For j = 0 To 4
        
        If Not i = j Then
          
          If imgDice(i).Picture = imgDice(j).Picture Or _
          imgDice(i).Picture = ilDice.ListImages(1).Picture And imgDice(j).Picture = ilDice.ListImages(6).Picture Then _
          lblNums(10) = 0
      
        End If
      
      Next j
    
    Next i
  
  End If
  '//--End large straight

  '//---Yahtzee
  If lblNums(11).Enabled = True Then
    
    If die(0) = die(1) And die(0) = die(2) And die(0) = die(3) And die(0) = die(4) Then
      
      lblNums(11) = 50
      
      If mnuSound.Checked Then PlaySound App.Path & "\Tada.wav" 'Play sound if available
    
    Else
      
      lblNums(11) = 0
    
    End If
  
  End If
  '//---End yahtzee

  '//---Chance
  If lblNums(12).Enabled = True Then
  
    lblNums(12) = 0
    
    For i = 0 To 4
      
      lblNums(12) = lblNums(12) + die(i)
    
    Next i
  
  End If
  '//---End chance

  '//---If it's a yahtzee then we need to allow total item value if used there
  If die(0) = die(1) And die(0) = die(2) And die(0) = die(3) And die(0) = die(4) Then
    
    If lblNums(13).Enabled = False Then GoTo SkipMe '//Bonus has been closed
    
    '//Left Columns
    For i = 0 To 5
      
      If lblNums(i).Enabled = True Then
        
        lblNums(i) = (i + 1) * 5
      
      End If
    
    Next i
  
    '//Three of a kind
    If lblNums(6).Enabled = True Then lblNums(6) = die(0) * 5
  
    '//Four of a kind
    If lblNums(7).Enabled = True Then lblNums(7) = die(0) * 5
  
    '//Full house
    If lblNums(8).Enabled = True Then lblNums(8) = 25
  
    '//Small straight
    If lblNums(9).Enabled = True Then lblNums(9) = 30
  
    '//Large straight
    If lblNums(10).Enabled = True Then lblNums(10) = 40
  
    '//Bonus Yahtzee
    If uYatzeeBonus = False Then
      
      If lblNums(11).Enabled = False And lblNums(11) = "50" Then
      
        If mnuSound.Checked Then PlaySound App.Path & "\Tada.wav" '//Play sound if available
      
        If lblNums(13).Enabled = True Then
      
          If lblNums(13) = "X" Then
        
            lblNums(13) = 100
      
          Else
        
            lblNums(13) = CStr(Int(lblNums(13) + 100))
      
          End If
    
        End If
  
      End If
  
    End If
  
  End If
  
SkipMe:

  '//Set all enabled items to blue
  For i = 0 To 12
    
    If lblNums(i).Enabled = True Then lblNums(i).ForeColor = vbBlue
  
  Next i

End Sub



