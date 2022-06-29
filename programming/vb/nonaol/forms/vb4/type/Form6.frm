VERSION 4.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "From Pyro"
   ClientHeight    =   3870
   ClientLeft      =   1380
   ClientTop       =   4980
   ClientWidth     =   4245
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Westminster"
      Size            =   11.25
      Charset         =   0
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   4305
   Left            =   1320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Top             =   4605
   Width           =   4365
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Text            =   " "
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Westminster"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Text            =   " "
      Top             =   360
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   -120
      Top             =   3600
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Westminster"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Timer1_Timer()
text1.Text = text1.Text + "H"
TimeOut (0.5)
text1.Text = text1.Text + "i"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "P"
TimeOut (0.5)
text1.Text = text1.Text + "y"
TimeOut (0.5)
text1.Text = text1.Text + "r"
TimeOut (0.5)
text1.Text = text1.Text + "o"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "H"
TimeOut (0.5)
text1.Text = text1.Text + "e"
TimeOut (0.5)
text1.Text = text1.Text + "r"
TimeOut (0.5)
text1.Text = text1.Text + "e"
TimeOut (0.5)
text1.Text = text1.Text + "."
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "T"
TimeOut (0.5)
text1.Text = text1.Text + "h"
TimeOut (0.5)
text1.Text = text1.Text + "i"
TimeOut (0.5)
text1.Text = text1.Text + "s"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "i"
TimeOut (0.5)
text1.Text = text1.Text + "s"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "M"
TimeOut (0.5)
text1.Text = text1.Text + "y"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "F"
TimeOut (0.5)
text1.Text = text1.Text + "i"
TimeOut (0.5)
text1.Text = text1.Text + "g"
TimeOut (0.5)
text1.Text = text1.Text + "h"
TimeOut (0.5)
text1.Text = text1.Text + "t"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "B"
TimeOut (0.5)
text1.Text = text1.Text + "o"
TimeOut (0.5)
text1.Text = text1.Text + "t"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "T"
TimeOut (0.5)
text1.Text = text1.Text + "h"
TimeOut (0.5)
text1.Text = text1.Text + "a"
TimeOut (0.5)
text1.Text = text1.Text + "t"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "I"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text1.Text = text1.Text + "M"
TimeOut (0.5)
text1.Text = text1.Text + "a"
TimeOut (0.5)
text1.Text = text1.Text + "d"
TimeOut (0.5)
text1.Text = text1.Text + "e"
TimeOut (0.5)
text1.Text = text1.Text + " "
TimeOut (0.5)
text2.Text = text2.Text + "F"
TimeOut (0.5)
text2.Text = text2.Text + "o"
TimeOut (0.5)
text2.Text = text2.Text + "r"
TimeOut (0.5)
text2.Text = text2.Text + " "
TimeOut (0.5)
text2.Text = text2.Text + "T"
TimeOut (0.5)
text2.Text = text2.Text + "h"
TimeOut (0.5)
text2.Text = text2.Text + "e"
TimeOut (0.5)
text2.Text = text2.Text + " "
TimeOut (0.5)
text2.Text = text2.Text + "H"
TimeOut (0.5)
text2.Text = text2.Text + "e"
TimeOut (0.5)
text2.Text = text2.Text + "l"
TimeOut (0.5)
text2.Text = text2.Text + "l"
TimeOut (0.5)
text2.Text = text2.Text + " "
TimeOut (0.5)
text2.Text = text2.Text + "O"
TimeOut (0.5)
text2.Text = text2.Text + "f"
TimeOut (0.5)
text2.Text = text2.Text + " "
TimeOut (0.5)
text2.Text = text2.Text + "I"
TimeOut (0.5)
text2.Text = text2.Text + "t"
TimeOut (0.5)
text2.Text = text2.Text + " "
TimeOut (0.5)
text2.Text = text2.Text + "."
TimeOut (0.5)
text2.Text = text2.Text + ".."
TimeOut (0.5)
text2.Text = text2.Text + " "
TimeOut (0.5)
text2.Text = text2.Text + "L"
TimeOut (0.5)
text2.Text = text2.Text + "A"
TimeOut (0.5)
text2.Text = text2.Text + "T"
TimeOut (0.5)
text2.Text = text2.Text + "E"
TimeOut (0.5)
text2.Text = text2.Text + "R"
TimeOut (2)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (0.2)
Text3.Text = ""
text2.Text = ""
text1.Text = ""
TimeOut (0.7)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (0.5)
Text3.Text = ""
TimeOut (0.7)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (0.5)
Text3.Text = ""
TimeOut (0.7)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (0.5)
Text3.Text = ""
TimeOut (0.7)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (0.5)
Text3.Text = ""
TimeOut (0.7)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (0.5)
Text3.Text = ""
TimeOut (0.7)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (0.5)
Text3.Text = ""
TimeOut (0.7)
Text3.Text = Text3.Text + "              SUBJECT LOST"
TimeOut (3)

Form5.Show
Unload Me
TimeOut (9999999)
End Sub


