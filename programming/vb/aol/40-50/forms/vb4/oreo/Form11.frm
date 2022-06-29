VERSION 4.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OreO 1.0 : Room Buster"
   ClientHeight    =   735
   ClientLeft      =   3645
   ClientTop       =   2655
   ClientWidth     =   3915
   Height          =   1140
   Icon            =   "Form11.frx":0000
   Left            =   3585
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Top             =   2310
   Width           =   4035
   Begin VB.TextBox Text1 
      BackColor       =   &H000000C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "Aol 40"
      Top             =   120
      Width           =   2175
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Stop"
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Start"
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
Attribute VB_Name = "Form11"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form2.Show
Unload Form11
End Sub


Private Sub SSCommand1_Click()
Do
Call KeyWord("aol://2719:2-2-" & Text1.Text)
Loop
End Sub


Private Sub SSCommand2_Click()
Do
DoEvents
Loop
End Sub


