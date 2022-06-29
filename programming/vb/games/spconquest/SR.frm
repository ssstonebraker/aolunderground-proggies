VERSION 5.00
Begin VB.Form SR 
   AutoRedraw      =   -1  'True
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   Icon            =   "SR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   7605
      Left            =   9930
      ScaleHeight     =   7545
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton Option1 
         Caption         =   "Seek and Destroy"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   360
         TabIndex        =   27
         Top             =   5880
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Text            =   "2"
         Top             =   7080
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Help"
         Height          =   375
         Left            =   600
         TabIndex        =   19
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label who 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Line Line10 
         X1              =   240
         X2              =   1320
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label13 
         Caption         =   "Skill Level  1-4"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Line Line8 
         BorderWidth     =   3
         X1              =   240
         X2              =   1320
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Player"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   480
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Health"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Power"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Health"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Armor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Power"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Armor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   5040
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   1320
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   1320
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   1320
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   1320
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   240
         Y1              =   960
         Y2              =   6600
      End
      Begin VB.Line Line6 
         X1              =   240
         X2              =   1320
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line7 
         X1              =   1320
         X2              =   1320
         Y1              =   960
         Y2              =   6600
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   840
         Picture         =   "SR.frx":164A
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   2
         Left            =   840
         Picture         =   "SR.frx":1954
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   1
         Left            =   840
         Picture         =   "SR.frx":1C5E
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   0
         Left            =   840
         Picture         =   "SR.frx":1F68
         Top             =   960
         Width           =   480
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   7500
      Left            =   9120
      Top             =   1560
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   11325
      TabIndex        =   3
      Top             =   7605
      Width           =   11385
      Begin VB.CommandButton Command2 
         Caption         =   "Clear Msg Log"
         Height          =   375
         Left            =   4200
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Audio"
         Height          =   375
         Left            =   3360
         TabIndex        =   22
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         CausesValidation=   0   'False
         ForeColor       =   &H80000009&
         Height          =   1035
         ItemData        =   "SR.frx":2272
         Left            =   5640
         List            =   "SR.frx":2274
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   5415
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   120
         X2              =   3240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label16 
         Caption         =   "V 1.8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "CheeZeWare Homepage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "SouthQuest "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "By Paul Bryan, 1999    Email: pb2012@mad.scientist.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label8 
         Caption         =   "Message Log:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   9120
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   75
      Left            =   9120
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   9120
      Top             =   120
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Press The ""Enter"" Key,  to Start a New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   3015
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   7320
      Picture         =   "SR.frx":2276
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Index           =   4
      Left            =   1200
      Picture         =   "SR.frx":2580
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Index           =   3
      Left            =   1200
      Picture         =   "SR.frx":288A
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   "SR.frx":2B94
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "SR.frx":2E9E
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Index           =   0
      Left            =   1200
      Picture         =   "SR.frx":31A8
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   -240
      TabIndex        =   1
      Top             =   120
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   6720
      Picture         =   "SR.frx":34B2
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   2280
      Shape           =   2  'Oval
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   6
      Left            =   6720
      Picture         =   "SR.frx":37BC
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   1560
      Picture         =   "SR.frx":4086
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   5
      Left            =   6240
      Picture         =   "SR.frx":4950
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   4
      Left            =   5640
      Picture         =   "SR.frx":521A
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   3
      Left            =   5040
      Picture         =   "SR.frx":5AE4
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   2
      Left            =   4440
      Picture         =   "SR.frx":63AE
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   3720
      Picture         =   "SR.frx":6C78
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   3120
      Picture         =   "SR.frx":7542
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   1080
      Picture         =   "SR.frx":7E0C
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   1560
      Picture         =   "SR.frx":86D6
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "SR.frx":8FA0
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   600
      Picture         =   "SR.frx":986A
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   840
      Picture         =   "SR.frx":A134
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "SR.frx":A9FE
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   6240
      Picture         =   "SR.frx":B2C8
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   5640
      Picture         =   "SR.frx":B5D2
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   4920
      Picture         =   "SR.frx":B8DC
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   4320
      Picture         =   "SR.frx":BBE6
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   3720
      Picture         =   "SR.frx":BEF0
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "SR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                   "Southpark Conquest v1.8"  by Paul Bryan 1999
' All code contained here-in, is the Sole Intellectual Property of the Author.
'
' The Bad Guys are Satan, Death, and Evil Damien Clones, Gang up on the Dark Forces before they Get your Characters.
' 'P' is to Pause the game,left-click on a character to view his status, then left-click on the screen
' and he'll move to that location. Right-click on the screen or an object,
' and all of your remaining charcters will do a formation move, to where you right-clicked.
'
'
' As you can see, this is still in development ... but it does do "something", as is, and now has Sound Effects.
'
'
Dim ch(10, 6) As Long 'ch(image, 0 = Health     ** ch = the good guys **
Dim em(10, 5) As Long          ' 1 = Hit Power  ** em = the bad guys **
Dim n(10) As String 'Char Name      2 = Speed
Dim en(10) As String 'Robots Name   3 = Armor
Public s As Integer               ' 4 = X destination
Public c As Integer               ' 5 = Y destination
                                  ' 6 = Seek & Destroy [on=1 off=0] )
Dim it(10) As Integer ' Bonus Item Values
Dim itn(10) As String ' Bonus Item Names
Public skd As Integer
Public strt As Integer
Public lv As Integer ' Level Counter
Public Build As Integer ' buildings
Public bh As Long ' building health

Private Sub Form_Load() ' initialize game / level
    If Command$ = "/nosnd" Then SR.Check1.Value = 0 ' check for start muted command line option.
    Randomize Timer()
    lv = lv + 1: If strt = 0 And (lv > 1 And lv < 5) Then lv = 5
    If lv > 1 Then GoTo pb
    List1.AddItem ("Game Started, Skill Level " & Text1.Text): ubx
    itn(0) = "+400 Health": itn(1) = "+400 Armor": itn(2) = "+10 Speed": itn(3) = "+100 Hit Power": itn(4) = "SUPER CHARGED !!!": itn(5) = " Brings "
    it(0) = 400: it(1) = 400: it(2) = 10: it(3) = 100: it(4) = 100: it(5) = 100 'initialize defaults and constants
    n(0) = "Mr Hanky": n(1) = "Cartman": n(2) = "Stan": n(3) = "Kyle": n(4) = "Kenny": n(5) = "Ike": n(6) = "Wendy"
    en(0) = "Satan": en(1) = "Death": For t = 2 To 6: en(t) = "Evil Damien #" + Str$(t - 1): Next t
pb:
    SR.Caption = "SouthQuest, Level" + Str$(lv)
    For t = 0 To 6: h = Int(Rnd(1) * 200) + 1: If h < 50 Then h = 50
    If lv > 1 Then GoTo pb2
    ch(t, 0) = h: ch(t, 1) = h ' good guys health and hit
    If strt = 0 Then ch(t, 0) = ch(t, 0) * 100: ch(t, 1) = ch(t, 1) * 20: skd = 1
pb2:
    
    If t < 2 Then em(t, 0) = (h * (lv * 10)): em(t, 1) = (h * (lv * 7)): GoTo pbb
    em(t, 0) = (h * Int(lv * 2)): em(t, 1) = (h * Int(lv * 3)) ' bad guys health and hit
pbb:
    h = Int(Rnd(1) * (100 + lv)) + 1: If h < 60 Then h = 60
    If lv > 1 Then GoTo pb3
    ch(t, 2) = h: ch(t, 3) = h ' good guys speed and armor
    If strt = 0 Then ch(t, 3) = ch(t, 3) * 20
    ch(t, 4) = Int(Rnd(1) * (SR.Height - 1800)): Image3(t).Top = ch(t, 4): ch(t, 5) = Int(Rnd(1) * (SR.Width - 2000)): Image3(t).Left = ch(t, 5)
                            'good guys position & destination
pb3:
    If lv < 2 And t < 2 Then GoTo pb5
    Image2(t).Visible = True
pb4:
    em(t, 4) = Int(Rnd(1) * (SR.Height - 1800)): em(t, 5) = Int(Rnd(1) * (SR.Width - 200)) ' bad guys destination
    em(t, 3) = h * lv: em(t, 2) = h ' bad guys armor and speed
pb5:
    Next t
    If strt = 0 Then Label17.Top = (SR.Height / 2): Label17.Left = (SR.Width / 2) - 1000
    Image3_Click (s)
    List1.AddItem ("Entered Level" + Str$(lv))
        Sndfx ("level.snd")
End Sub

Private Sub Timer1_Timer() ' Player Movement Engine
    For k = 0 To 6: uu = 0: If Image3(k).Visible = False Then GoTo pb2
    X = Image3(k).Top: Y = Image3(k).Left:
    If X <= (ch(k, 4) + ch(k, 2)) And X >= (ch(k, 4) - ch(k, 2)) Then uu = 1: GoTo pb
    If X < ch(k, 4) Then Image3(k).Top = Image3(k).Top + ch(k, 2)
    If X > ch(k, 4) Then Image3(k).Top = Image3(k).Top - ch(k, 2)
    If k = s Then Shape1.BorderColor = vbGreen: Shape1.Top = Image3(s).Top - 60: Shape1.Left = Image3(s).Left - 60:
pb:
    If Y <= (ch(k, 5) + ch(k, 2)) And Y >= (ch(k, 5) - ch(k, 2)) Then uu = uu + 1: GoTo pb1
    If Y < ch(k, 5) Then Image3(k).Left = Image3(k).Left + ch(k, 2)
    If Y > ch(k, 5) Then Image3(k).Left = Image3(k).Left - ch(k, 2)
    If k = s Then Shape1.BorderColor = vbGreen: Shape1.Top = Image3(s).Top - 60: Shape1.Left = Image3(s).Left - 60:
pb1:
    chkarea (k)
pb2:
    If ch(k, 6) = 1 And uu > 1 Then seekdest (k)
    If skd = 1 And uu > 1 Then seekdest (k)
    Next k
End Sub
Private Sub Timer2_Timer() ' Robot Movement Engine
    For k = 0 To 6: u = 0: If Image2(k).Visible = False Then GoTo pb2
    X = Image2(k).Top: Y = Image2(k).Left
    If X < (em(k, 4) + em(k, 2)) And X > (em(k, 4) - em(k, 2)) Then u = 1: GoTo pb
    If X < em(k, 4) Then Image2(k).Top = Image2(k).Top + em(k, 2): u = 1
    If X > em(k, 4) Then Image2(k).Top = Image2(k).Top - em(k, 2): u = 1
    If k = c Then Shape2.Top = Image2(c).Top - 60: Shape2.Left = Image2(c).Left - 60
pb:
    If Y < (em(k, 5) + em(k, 2)) And Y > (em(k, 5) - em(k, 2)) Then u = u + 1: GoTo pb2
    If Y < em(k, 5) Then Image2(k).Left = Image2(k).Left + em(k, 2): u = 1
    If Y > em(k, 5) Then Image2(k).Left = Image2(k).Left - em(k, 2): u = 1
    If k = c Then Shape2.Top = Image2(c).Top - 60: Shape2.Left = Image2(c).Left - 60
pb2:
    If Timer1.Enabled = True And (u = 0 And Image2(k).Visible = False) Then Timer1_Timer
    If u > 1 Then rchngdir (k): u = 0
    If Build < 1 Then GoTo pbb
    xp = Shape3.Top: yp = Shape3.Left
    If Shape3.Visible = False Then GoTo pbb
    If (X > xp - 480) And (X <= xp + Shape3.Height) Then GoTo paul
    GoTo pbb
paul:
     If (Y > yp - 480) And (Y <= yp + Shape3.Width) Then Label15.Visible = True: Image6(0).Visible = True: Image6(0).Top = X: Image6(0).Left = Y: uu = lv * 2: em(k, 0) = em(k, 0) - uu: bh = bh - ((Int(em(k, 1) / 700)) * lv): Label15.Caption = bh: Label15.Top = Shape3.Top + 100: Label15.Left = Shape3.Left + 100: rchngdir (k)
pbb:
    Next k
End Sub
Private Sub Timer3_Timer() ' Bonus Item Engine
    If Shape3.Visible = False Then GoTo yawn
    If bh < 0 Then bh = 0: Build = 0: Shape3.Visible = False: Label15.Caption = "Destroyed": List1.AddItem ("Building Destroyed"): ubx
yawn:
    If Shape1.Visible = True Then Shape1.Visible = False
    If Shape2.Visible = True Then Shape2.Visible = False
    For paul = 0 To 4
    If Image6(paul).Visible = True Then Image6(paul).Visible = False
    Next paul
    Label10.Caption = ""
    Label15.Caption = ""
    pb = Int(Rnd(1) * 10) + 1: If pb > 6 Then GoTo pbb
    pb = pb - 1
    If pb = 5 Then pb = Int(Rnd(1) * 5) + 1: If pb <> 5 Then pb = 0
    Image1(pb).Top = Int(Rnd(1) * (SR.Height - 1800)): Image1(pb).Left = Int(Rnd(1) * (SR.Width - 2000))
    Image1(pb).Visible = True
pbb:
    If (lv >= 5 And pb < 3) And Build < 1 Then Build = Build + 1: Image1(6).Top = Int(Rnd(1) * (SR.Height - 6000)): Image1(6).Left = Int(Rnd(1) * (SR.Width - 6000)): Image1(6).Visible = True
    
End Sub
Private Sub Timer4_Timer() ' see if any of your charcters are still alive
    If strt = 0 Then Label17.Visible = True
    For z = 0 To 6: If Image2(z).Visible = True Then GoTo pbb
    Next z: GoTo pb2
pbb:
    sc = 0
    For z = 0 To 6: If Image3(z).Visible = True Then GoTo pb
    sc = sc + (ch(z, 1) + ch(z, 3) + ch(z, 2))
    Next z: For z = 0 To 3: Label5(z) = ch(t, z): Next z
    If strt = 0 Then GoTo demo
    MsgBox "You've Been Dragged to the depths of Hell" + Chr$(13) + "With a Total Score of " + Str$(sc * lv) + " on Level" + Str$(lv)
    List1.AddItem ("Game Over-Score=" + Str$(sc * lv) + " on Level" + Str$(lv)): ubx
demo:
    For z = 0 To 6: Image3(z).Visible = True: Image2(z).Visible = False: ch(z, 6) = 0: Next z
    lv = 0: Build = 0: Shape3.Visible = False: skd = 0: strt = 0:
    Label17.Enabled = True: Label17.Visible = True
pb2:
   For paul = 0 To 6: Image2(paul).Top = Int(Rnd(1) * (SR.Height - 1800)): Image2(paul).Left = Int(Rnd(1) * (SR.Width - 2000))
   Next paul:
   Form_Load
pb:
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then GoTo rclk ' if right-click then rclk
    ch(s, 5) = X: ch(s, 4) = Y ' good guys destination = where form was clicked
    GoTo dn
rclk:
    Call Moveall(Y, X) ' Do the moveall routine
dn:
End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = Image1(Index).Top: Y = Image1(Index).Left
    If Button = 2 Then GoTo js 'if right-click then js
    ch(s, 5) = Y: ch(s, 4) = X: GoTo dn 'good guys destination = where the image that is clicked is located.
js:
    Call Moveall(X, Y) 'Do the moveall routine, to where the image that was click, is located
dn:
End Sub
Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    c = Index: If Image2(c).Visible = False Then GoTo dn ' if clicked on image thats not visible then nevermind
    If Button = 2 Then GoTo js 'if right-click then js
    GoTo pz
js:
    Call Moveall(Image2(c).Top, Image2(c).Left) 'Do the moveall routine
pz:
    If Shape2.Visible = False Then Shape2.Visible = True ' show outline
    Shape2.BorderColor = vbBlack
    Shape2.Top = Image2(c).Top - 60: Shape2.Left = Image2(c).Left - 60 ' position outline around seleted charcter
    Image5.Picture = Image2(c).Picture: Label6.Caption = en(c)
    For t = 0 To 3: Label5(t) = em(c, t): Next t ' update the status area
dn:
End Sub
Private Sub Image3_Click(Index As Integer)
    
    s = Index: Shape1.Visible = True: Shape1.BorderColor = vbGreen 'show outline
    Shape1.Top = Image3(s).Top - 60: Shape1.Left = Image3(s).Left - 60 ' position outline around seleted charcter
    Call status(s) ' update the status area
End Sub
Private Sub Text1_Change()
    paul = Val(Text1.Text): If paul < 1 Or paul > 4 Then paul = 1
    pb = Int(150 / paul): Timer2.Interval = pb: List1.AddItem ("Skill Level " & paul & " Selected")
    Picture2.SetFocus
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then GoTo rclk
    ch(s, 5) = X: ch(s, 4) = Y ' good guys destination = where label is located
    GoTo dn
rclk:
    Call Moveall(Y, X) 'Do the moveall routine
dn:
End Sub

Private Sub rchngdir(ByVal t As Integer) ' Robot destination may be set to either a bonus item or a good guy
    u = Int(Rnd(1) * 100): If u < 80 Then GoTo dn ' %20 chance of bad guy getting a new destination
    If em(t, 2) > 150 Then em(t, 2) = 150         ' ^ this causes a random delay before moving again.
    If u >= 98 Then GoTo home
    For paul = 0 To 6:  bryan = Int(Rnd(1) * 10) + 1: If Image3(paul).Visible = True And bryan > 5 Then em(t, 4) = Image1(paul).Top: em(t, 5) = Image1(paul).Left: em(t, 2) = em(t, 2) + lv: GoTo dn
    Next paul: em(t, 4) = Int(Rnd(1) * (SR.Height - 1800)): em(t, 5) = Int(Rnd(1) * (SR.Width - 2000)): em(t, 2) = em(t, 2) + 1
    GoTo dn
home:
    For paul = 0 To 5: bryan = Int(Rnd(1) * 10) + 1: If (Image3(paul).Visible = True And bryan > 5) And Build > 0 Then em(t, 4) = Image3(paul).Top: em(t, 5) = Image3(paul).Left: em(t, 2) = em(t, 2) + lv: GoTo dn
    Next paul: em(t, 4) = Int(Rnd(1) * (SR.Height - 1800)): em(t, 5) = Int(Rnd(1) * (SR.Width - 2000)): em(t, 2) = em(t, 2) + 1
dn:
End Sub
Private Sub chkarea(ByVal t As Integer) ' chk area around each charcter and respond to items within parameter accordingly
    For p = 0 To 6
    If em(p, 0) <= 0 Or Image2(p).Visible = False Then em(p, 0) = 0: Image2(p).Visible = False: GoTo no ' if Enemy is dead then skip him
    If Image2(p).Visible = False Then GoTo no
    If Image2(p).Top <= (Image3(t).Top + 300) And Image2(p).Top >= (Image3(t).Top - 780) Then GoTo wt
    GoTo no
wt:
    If Image2(p).Left <= (Image3(t).Left + 300) And Image2(p).Left >= (Image3(t).Left - 780) Then GoTo hit
    GoTo no
hit:
    If ch(t, 3) > 1 Then uu = Int(em(p, 3) / 20): ch(t, 1) = ch(t, 1) - uu: GoTo zt
    uu = Int(em(p, 1) / 10)
zt:
    ch(t, 0) = ch(t, 0) - uu: If ch(t, 0) < 0 Then ch(t, 0) = 0
    If ch(t, 1) < 0 Then ch(t, 1) = 0
   
pt:
    If strt = 0 Then Label17.Left = Image3(t).Left - 500: Label17.Top = Image3(t).Top - 500: Shape1.Top = Image3(t).Top - 60: _
                     Shape1.Left = Image3(t).Left - 60: Shape2.Top = Image2(p).Top - 60: Shape2.Left = Image2(p).Left - 60: _
                     Shape2.BorderColor = vbRed: Shape1.BorderColor = vbRed: Shape1.Visible = True: Shape2.Visible = True: GoTo dm
    
    Shape1.BorderColor = vbRed: Shape1.Visible = True: sd = Int(Rnd(1) * 6) + 1: Sndfx ("oo" & sd & ".snd")
    Shape1.Top = Image3(t).Top - 60: Shape1.Left = Image3(t).Left - 60
    Label10.ForeColor = vbRed: Label10.Top = Image3(t).Top - 100: Label10.Left = Image3(t).Left - 100: Label10.Caption = n(t) + " is being attacked by " + en(p) + " for" + Str$(uu) + " Points.": List1.AddItem (n(t) + " is being attacked by " + en(p)) + " for" + Str$(uu) + " Health Points.": ubx
    If Shape2.Visible = False Then Shape2.Visible = True
    Shape2.Top = Image2(p).Top - 60: Shape2.Left = Image2(p).Left - 60
    Shape2.BorderColor = vbRed
dm:
    If ch(t, 0) <= 0 Then ch(t, 0) = 0: Image3(t).Visible = False: Label10.ForeColor = vbRed: Label10.Top = Image3(t).Top - 100: Label10.Left = Image3(t).Left - 100: Label10.Caption = n(t) + " was KILLED by " + en(p): List1.AddItem (n(t) + " was KILLED by " + en(p)): fx = Int(Rnd(1) * 3) + 1: Sndfx ("dd" & fx & ".snd"): ubx
    If em(p, 3) > 1 Then em(p, 0) = em(p, 0) - Int(ch(t, 3) / 20): em(p, 1) = em(p, 1) - Int(ch(t, 3) / 20): GoTo yo
    em(p, 0) = em(p, 0) - Int(ch(t, 3) / 10)
yo:
    If em(p, 0) <= 0 Then em(p, 0) = 0: Image2(p).Visible = False: Label10.ForeColor = vbBlue: Label10.Top = Image3(t).Top - 10: Label10.Left = Image3(t).Left - 100: Label10.Caption = en(p) + " was KILLED by " + n(t): List1.AddItem (en(p) + " was KILLED by " + n(t)): ubx
no:
    If p > 6 Then GoTo wo
    If Image1(p).Visible = False Then GoTo wo
    If Image1(p).Top <= (Image3(t).Top + 300) And Image1(p).Top >= (Image3(t).Top - 780) Then GoTo wh
    GoTo wo
wh:
    If Image1(p).Left <= (Image3(t).Left + 300) And Image1(p).Left >= (Image3(t).Left - 780) Then GoTo got
    GoTo wo
got:
    If p > 3 Then GoTo clover
    If p = 2 And ch(t, p) > 140 Then Label10.Top = Image3(t).Top - 100: Label10.Left = Image3(t).Left - 100: Label10.ForeColor = vbRed: Label10.Caption = n(t) + " Is Already at Maximum Speed": GoTo wo
    fx = Int(Rnd(1) * 2) + 1: Sndfx ("gg" & fx & ".snd")
    Label10.ForeColor = vbBlack: Label10.Top = Image3(t).Top - 100: Label10.Left = Image3(t).Left - 100: Label10.Caption = n(t) + " Gets " + itn(p): List1.AddItem (n(t) + " Got " + itn(p)): ubx
    Image1(p).Visible = False
    ch(t, p) = ch(t, p) + it(p): Call status(t)
    GoTo wo
clover:
    If p = 5 Then GoTo undead
    If (p = 6 And Build < 2) Then Build = Build + 1: Image1(p).Visible = False: GoTo Bld
    Sndfx ("gg1.snd")
    Label10.Caption = n(t) + " Gets " + itn(p) + " !!!!": Label10.ForeColor = vbBlack: Label10.Top = Image3(t).Top - 100: Label10.Left = Image3(t).Left - 100
    List1.AddItem (n(t) + " Got " + itn(p) + " !!!!"): ubx
    Call status(t)
    For z = 0 To 3
    ch(t, z) = ch(t, z) + it(z): If z = 2 And ch(t, z) > 150 Then ch(t, z) = 160
    Next z: Image1(p).Visible = False
    GoTo wo
undead:
    
    For z = 0 To 6: If Image3(z).Visible = True Then GoTo foob
    Image3(z).Visible = True: ch(z, 0) = it(p): ch(z, 1) = it(p): ch(z, 3) = it(p)
    Label10.ForeColor = vbBlue: Label10.Top = Image3(t).Top - 10: Label10.Left = Image3(t).Left - 100: Label10.Caption = n(t) + itn(p) + n(z) + " Back to LIFE !!!": List1.AddItem (n(t) + itn(p) + n(z) + " Back to LIFE !!!"): ubx
    Image1(p).Visible = False: Sndfx ("gg2.snd")
    GoTo wo
foob:
    Next z: Label10.ForeColor = vbRed: Label10.Top = Image3(t).Top - 10: Label10.Left = Image3(t).Left - 100: Label10.Caption = n(t) + " Can't Use this item right now!!!": GoTo wo

Bld:
    bh = lv * (10 * ch(z, 2))
    Label10.Caption = n(t) + " Builds a Safety Compound": Label10.ForeColor = vbBlack: Label10.Top = Image3(t).Top - 100: Label10.Left = Image3(t).Left - 100: List1.AddItem ("Safety Compound Built"): ubx
     Shape3.Visible = True: Shape3.Top = Image3(t).Top: Shape3.Left = Image3(t).Left
    

wo:
    Next p
n:
    If ch(t, 3) < 0 Then ch(t, 3) = 0
    If em(p, 3) < 0 Then em(0, 3) = 0
    If Val(who.Caption) = t Then status (t)
End Sub
Private Sub status(ByVal pb As Integer) ' Update the charcter information area
Image5.Picture = Image3(pb).Picture: Label6.Caption = n(pb)
    For t = 0 To 3: Label5(t) = ch(pb, t): Next t: who.Caption = pb
    If ch(pb, 6) = 1 Then Option1.Value = True
    If ch(pb, 6) = 0 Then Option1.Value = False
End Sub
Private Sub Moveall(ByVal X As Single, ByVal Y As Single) ' do a formation move to x,y
    u = -1: For k = 0 To 6: If Image3(k).Visible = False Then GoTo pb ' if charater is dead then skip him
    u = u + 1 ' Offset Coordinates (keeps characters from stacking on top of each other)
    ch(k, 5) = (Y + (u * 200)): ch(k, 4) = (X + (u * 200)) ' change characters destination to x,y + offset
pb:
    Next k ' change next characters destination
End Sub
Private Sub Label11_Click() 'contact info
    Label9_Click
End Sub
Private Sub Command2_Click()
    List1.Clear
End Sub
Private Sub Label9_Click() 'contact info
    Shell "start mailto:pb2012@mad.scientist.com"
End Sub
Private Sub Label14_Click() 'contact info
    Shell "start http://pbryan.webjump.com"
End Sub
Public Sub Sndfx(fName As String) ' sound effects routine (the Only API Call)
    If Check1.Value = 0 Then GoTo pb ' muted
    PlaySndFx fName, 1
pb:
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer) ' Key control
    If KeyAscii = 13 Then GoTo Start
    If KeyAscii = 112 Then GoTo pause
    If KeyAscii = 99 Then Image1(5).Visible = True: Image1(5).Top = Int(Rnd(1) * (SR.Height - 1800)): Image1(5).Left = Int(Rnd(1) * (SR.Width - 2000)) ' cheat = 'c' key
    If KeyAscii = 115 Then Build = 1: Image1(6).Top = Int(Rnd(1) * (SR.Height - 6000)): Image1(6).Left = Int(Rnd(1) * (SR.Width - 6000)): Image1(6).Visible = True ' cheat = 's' Key for Safety Parameter
    If KeyAscii = 103 Then GoTo mad ' "g" Cheat key is Mad Mode
    If KeyAscii = 107 And skd = 0 Then skd = 1: GoTo dn ' "k" Kill 'Em All Cheat Key!!
    If KeyAscii = 107 And skd = 1 Then skd = 0 ' turn it off
    GoTo dn
pause:
    If Timer1.Enabled = True Then Timer1.Enabled = False: Label7.Caption = "Game Paused, When Ready, Hit 'P' to Start, 'P' = Pause": List1.AddItem ("Game Paused"): GoTo dn
    If Timer1.Enabled = False Then Timer1.Enabled = True: Label7.Caption = "": List1.AddItem ("Game Resumed"): GoTo dn
    GoTo dn:
Start:
    For z = 0 To 6: Image3(z).Visible = True: Image2(z).Visible = False: ch(z, 6) = 0: Next z
    lv = 0: Build = 0: Shape3.Visible = False: skd = 0: strt = 1
    Label17.Enabled = False: Label17.Visible = False: Form_Load: GoTo dn
mad:
    For paul = 0 To 6: If ch(paul, 0) > 50000 Then GoTo dn
    ch(paul, 0) = ch(paul, 0) + 10000: ch(paul, 1) = ch(paul, 1) + 500: ch(paul, 3) = ch(paul, 3) + 1000
    status (p): Next paul
    
dn:
End Sub
Public Sub ubx() 'Scroll Message Log
    List1.ListIndex = List1.NewIndex
End Sub
Private Sub Command1_Click() ' help
    MsgBox "The Bad Guys are Satan, Death, and Evil Damien Clones," + Chr$(13) + "Gang up on the Dark Forces before they Get your Characters." + _
    "'P' is to Pause the game,left-click on a character to view his status," + Chr$(13) + "then left-click on the screen" + _
    " and he'll move to that location. Right-click on the screen or an object," + Chr$(13) + "and all of your remaining charcters will do a formation move, to where you right-clicked." + _
    Chr$(13) + "The Object of the Game, is to Collect Enough Bonus Items, to Destroy The Evil Forces and Advance in Level." + _
    Chr$(13) + "Skill Level and Game Speed are the Same Level 1 is easiest(Slowest), 4 is Hardest(Fastest). GOOD LUCK !"
End Sub
Private Sub Label12_Click() ' easter egg
   Call Main
   For paul = 0 To 4: Image6(paul).Visible = True
   Image6(paul).Top = Int(Rnd(1) * (SR.Height - 1800)): Image6(paul).Left = Int(Rnd(1) * (SR.Width - 2000))
   Next paul
End Sub
Private Sub Option1_Click() ' seek and destroy
    pb = who.Caption: If Image3(pb).Visible = False Then GoTo dn
    If ch(pb, 6) = 1 Then ch(pb, 6) = 0: GoTo dn
    If ch(pb, 6) = 0 Then ch(pb, 6) = 1: GoTo dn
dn:
End Sub
Private Sub seekdest(pb As Integer)
wow:
    If stst = 1 Then GoTo wow2
    pz = Int(Rnd(1) * 100) + 1: If pz < 50 Then GoTo wow2
    paul = (Int(Rnd(1) * 7) + 1) - 1: If Image1(paul).Visible = False Then GoTo dn
    ch(pb, 4) = Image1(paul).Top: ch(pb, 5) = Image1(paul).Left: GoTo dn

wow2:
    paul = (Int(Rnd(1) * 7) + 1) - 1: If Image2(paul).Visible = False Then GoTo dn
    ch(pb, 4) = Image2(paul).Top: ch(pb, 5) = Image2(paul).Left: GoTo dn
    
dn:
End Sub

