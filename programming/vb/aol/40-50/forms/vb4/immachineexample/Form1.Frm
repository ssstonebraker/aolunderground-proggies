VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3330
   ClientLeft      =   2370
   ClientTop       =   2085
   ClientWidth     =   6465
   ControlBox      =   0   'False
   Height          =   3735
   Icon            =   "Form1.frx":0000
   Left            =   2310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   ShowInTaskbar   =   0   'False
   Top             =   1740
   Width           =   6585
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed Light"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0442
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed Light"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":04AE
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox text2 
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Text            =   "Person you want to ignore"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6720
      Top             =   2280
   End
   Begin VB.ListBox List1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      ItemData        =   "Form1.frx":04FA
      Left            =   120
      List            =   "Form1.frx":04FC
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your ignore message"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Machine Message"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   3960
      Picture         =   "Form1.frx":04FE
      Top             =   2880
      Width           =   2400
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   5640
      Picture         =   "Form1.frx":2AC0
      Top             =   2400
      Width           =   705
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   2640
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   4800
      Picture         =   "Form1.frx":39A2
      Top             =   2400
      Width           =   795
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   3960
      Picture         =   "Form1.frx":4A24
      Top             =   2400
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   3345
      Left            =   0
      Picture         =   "Form1.frx":5AA6
      Top             =   0
      Width           =   6480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False








Private Sub Form_Click()
FormDrag Me
End Sub

Private Sub Form_Load()
MsgBox "Hi, I found an example of an im machine on a site made by: Dr.Tru, I though that i could improve it a bit more and a little bit better look, so i decided to see what i could do.  It is an im answering machinge along with a im ignorer, so you don't have to be bothered by people you don't want to talk to. My name is aNGUs and if you have any questions feel free to email me at ooangusoo@aol.com"
Call FormOnTop(Me)
Timer1.Enabled = False
End Sub


Private Sub Image1_Click()
FormDrag Me
End Sub

Private Sub Image2_Click()
Timer1.Enabled = True
End Sub

Private Sub Image3_Click()
Timer1.Enabled = False
End Sub

Private Sub Image5_Click()
Form1.WindowState = 1
End Sub

Private Sub Image6_Click()
End
End Sub

Private Sub Timer1_Timer()
Dim Text As String
If FindIM <> 0 Then
GoTo Hello:

Hello:
If IMSender = text2 Then
Call InstantMessage(IMSender, text3)
TimeOut 1
Call CloseWindow(FindIM&)
Else: GoTo Programmers:

Programmers:
List1.AddItem IMSender & ": " & IMLastMsg
Call InstantMessage(IMSender, Text1)
TimeOut 1
Call CloseWindow(FindIM&)
End If
End If
End Sub




