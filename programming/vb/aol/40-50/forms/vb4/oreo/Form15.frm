VERSION 4.00
Begin VB.Form Form15 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OreO 1.0 : Request Bot"
   ClientHeight    =   1065
   ClientLeft      =   3990
   ClientTop       =   3225
   ClientWidth     =   3210
   Height          =   1470
   Icon            =   "Form15.frx":0000
   Left            =   3930
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Top             =   2880
   Width           =   3330
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "Form15.frx":030A
      Left            =   1440
      List            =   "Form15.frx":030C
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Some ass"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   1680
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Stop"
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Start"
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
Attribute VB_Name = "Form15"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form2.Show
Unload Form15
End Sub


Private Sub SSCommand1_Click()
Timer1.Enabled = True
SendChat "<b><i><s>" & BlackRedBlack("(¯`·¸ØRëO ¹·° Request BOT¸·´¯)")
timeout 0.3
SendChat "<b><i><s>" & BlackRedBlack("Requesting ""+ text1.text +""")
timeout 0.3
SendChat "<b><i><s>" & BlackRedBlack("type-/I got it- if you have it")
End Sub


Private Sub SSCommand2_Click()
Timer1.Enabled = False
SendChat "<b><i><s>" & BlueBlack("(¯`·¸ØRëO ¹·° Request bot off¸·´¯)")
End Sub


Private Sub Timer1_Timer()
If LastChatLine = "/I got it" Then
List1.AddItem SNFromLastChatLine
SendChat "" + SNFromLastChatLine + " Can you send!"
For X = 0 To List1.ListCount - 1
Call IMKeyword("" + List1.List(X) + "", "Sup " + List1.List(X) + ", ""Can you send me " + Text1.Text + "")
Next X
End If
End Sub


