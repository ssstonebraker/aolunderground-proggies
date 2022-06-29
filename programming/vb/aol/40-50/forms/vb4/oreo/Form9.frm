VERSION 4.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   855
   ClientLeft      =   3645
   ClientTop       =   75
   ClientWidth     =   3780
   Height          =   1260
   Left            =   3585
   LinkTopic       =   "Form9"
   ScaleHeight     =   855
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Top             =   -270
   Width           =   3900
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   1508
      _StockProps     =   15
      BackColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      BevelOuter      =   1
      BevelInner      =   1
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "   OREO 1.0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.Line Line2 
         X1              =   2640
         X2              =   2640
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   1200
         X2              =   1200
         Y1              =   240
         Y2              =   720
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Exit"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "IMz Off"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   6.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "IMz on"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   6.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Maximize"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   6.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub SSCommand1_Click()
Form2.Show
Unload Form9
End Sub


Private Sub SSCommand2_Click()
SendChat "<b><i><s><font Face= Arial>" & BlackRedBlack("ØRëO ¹·° • ÍMz Š†á†úš • ÍMz ÒÑ •")
Call IMKeyword("$IM_ON", "OreO Is Da Shit ")
End Sub

Private Sub SSCommand3_Click()
SendChat "<b><i><s><font Face= Arial>" & BlackRedBlack("ØRëO ¹·° • ÍMz Š†á†úš • ÍMz ÒÑ •")
Call IMKeyword("$IM_OFF", "OreO Is Da Shit ")
End Sub

Private Sub SSCommand4_Click()
Unload Form9
Unload Form2
End Sub


