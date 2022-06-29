VERSION 5.00
Begin VB.Form TrainLoadUp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "TrainLoadUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   420
      Left            =   225
      ScaleHeight     =   360
      ScaleWidth      =   4005
      TabIndex        =   0
      Top             =   765
      Width           =   4065
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         TabIndex        =   1
         Top             =   0
         Width           =   3930
      End
      Begin VB.Shape Bar 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   690
         Left            =   0
         Top             =   -90
         Width           =   2355
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1485
      Top             =   1305
   End
   Begin VB.Image IconImg 
      Height          =   420
      Left            =   3870
      Top             =   180
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Quad-Ball By Arvinder Sehmi 1999."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   1215
      Width           =   4290
   End
   Begin VB.Label CurrLoad 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   3
      Top             =   405
      Width           =   3930
   End
   Begin VB.Label Label2 
      Caption         =   "Loading Quad-Ball Training Mode..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   45
      Width           =   4065
   End
End
Attribute VB_Name = "TrainLoadUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Form Shows How Much of the Game Has loaded'
'________________________________________________'
Private Sub Form_Load()
 IconImg = Me.Icon
End Sub
Private Sub form_Paint()
 Bar.Width = Int((Picture1.Width / 14) * Val(Me.Caption))
 Percent.Caption = Int(100 / 14 * Val(Me.Caption)) & "%"
 colpercent = Int((255 / 14) * Val(Me.Caption))
 Bar.FillColor = RGB(0, 255 - colpercent, colpercent)
 Percent.ForeColor = RGB(0, colpercent, 255 - colpercent)
 Bar.BorderColor = Bar.FillColor
 Me.Caption = Percent.Caption
End Sub
