VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "triangle test"
   ClientHeight    =   1125
   ClientLeft      =   2910
   ClientTop       =   3585
   ClientWidth     =   1905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1530
   Left            =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   1905
   Top             =   3240
   Width           =   2025
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "check"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   -15
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Iz representing eXcel 2001
If Text1.Text = "" Then Exit Sub ' tells vb to quit if one of the textboxes are empty
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub
a = Text1.Text ' sets your textboxes to intigers...so you dont have to type text1, text 2 etc...
b = Text2.Text
c = Text3.Text
'rofl this has somthing to do with math..i cant remember wtf it does ;x
'i made this last year in school for my programming class....so..umm..you figure it out ;x
'ok i just remembered it...it checks to see if its an isocilies triangle...or htf ever you spell it
'well fuck i checked that and it wasnt that either...if you can figure out what i programemd
'email me at Sir_Iz@Hotmail.com ;x
If Val(a) + Val(b) > Val(c) And Val(a) + Val(c) > Val(b) And Val(c) + Val(b) > Val(a) Then
answer = "yes" ' uses "answer" as a variable
Else            'if it wasnt yes, "else" will make it go to no
answer = "no" ' sets answer to no
End If 'ends your if statement
Label1.Caption = answer ' put the answer (yes or no) in label1
End Sub

Private Sub Command2_Click()
Text1.Text = "" ' clears everything
Text2.Text = ""
Text3.Text = ""
Label1.Caption = ""
End Sub


