VERSION 4.00
Begin VB.Form Form14 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OreO 1.0 : Attention Bot"
   ClientHeight    =   855
   ClientLeft      =   3915
   ClientTop       =   3210
   ClientWidth     =   3420
   Height          =   1260
   Icon            =   "Form14.frx":0000
   Left            =   3855
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   Top             =   2865
   Width           =   3540
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
      Left            =   1440
      TabIndex        =   1
      Text            =   "I Need Some ass"
      Top             =   240
      Width           =   1935
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Send"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
End
Attribute VB_Name = "Form14"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Form14
Form.Show
End Sub


Private Sub SSCommand1_Click()
SendChat "<b><i><s>" & BlackRedBlack("(¯`·¸A.T.T.E.N.T.I.O.N¸·´¯)")
timeout 0.3
SendChat "<b><i>" & BlackRedBlack(Text1.Text)
timeout 0.3
SendChat "<b><i><s>" & BlackRedBlack("(¯`·¸A.T.T.E.N.T.I.O.N¸·´¯)")
End Sub


