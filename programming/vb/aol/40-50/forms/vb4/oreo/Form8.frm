VERSION 4.00
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Who's Da PoSer?"
   ClientHeight    =   975
   ClientLeft      =   3645
   ClientTop       =   2595
   ClientWidth     =   3915
   Height          =   1380
   Icon            =   "Form8.frx":0000
   Left            =   3585
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Top             =   2250
   Width           =   4035
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Text            =   "-Screen Name-"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Text            =   "-Screen Name-"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FuSH's"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PoOH's"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "CHECK"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   11.25
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
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Verirify FuSH Yu MaNG"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   6.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "CHECK"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
End
Attribute VB_Name = "Form8"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form2.Show
Unload Form8
End Sub


Private Sub SSCommand1_Click()
If Text1.Text = "Screen Name" Then
Else
End If
If Text1.Text = "DJPooH143" Then
MsgBox "WoooHooo It's the real PoOH", 64, "Real"
Else
MsgBox "A Poser Kill this Basterd for Me!", vbinfermation, "Poser"
Unload Me
End If
End Sub


Private Sub SSCommand2_Click()
Do
Form8.Height = Form8.Height + 5
Loop Until Form8.Height > 2240
End Sub


Private Sub SSCommand3_Click()
If Text2.Text = "Screen Name" Then
Else
End If
If Text2.Text = "Ejoech27" Then
MsgBox "WoooHooo It's the real FuSH Yu MaNG", 64, "Real"
Else
MsgBox "A Poser Kill this Basterd for Me!", vbinfermation, "Poser"
Unload Me
End If
End Sub


