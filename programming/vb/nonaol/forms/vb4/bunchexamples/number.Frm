VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "guess the number"
   ClientHeight    =   900
   ClientLeft      =   3585
   ClientTop       =   4200
   ClientWidth     =   1560
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
   Height          =   1305
   Left            =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   1560
   Top             =   3855
   Width           =   1680
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   540
      TabIndex        =   5
      Top             =   1035
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   255
      Left            =   810
      TabIndex        =   4
      Top             =   600
      Width           =   540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "check"
      Height          =   255
      Left            =   270
      TabIndex        =   3
      Top             =   600
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      MaxLength       =   3
      TabIndex        =   2
      Top             =   45
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "genorate"
      Height          =   255
      Left            =   270
      TabIndex        =   0
      Top             =   345
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "your guess:"
      Height          =   255
      Left            =   15
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Iz representing eXcel 2001
Beep ' beeps cuz beeping is fun ;x
Text2.Text = "" ' clears text2(text2 is hidden)
Randomize ' radonmizes numbers
Number = Int((100 * Rnd) + 1) ' tells vb what numbers to randomize between
Text2.Text = Number 'puts the random number in the hidden textbox
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then Exit Sub ' if your textboxes are empty..it wont do nothing
If Text2.Text = "" Then Exit Sub
Beep ' beeping is fun, remember?
If Val(Text1.Text) < Val(Text2.Text) Then MsgBox "your guess was to low": Exit Sub
'/\ check to see if the guess is lower than the genoraten number
If Val(Text1.Text) > Val(Text2.Text) Then MsgBox "your guess was to high": Exit Sub
'/\ checks to see if your guess was too high
If Val(Text1.Text) = Val(Text2.Text) Then MsgBox "good guess": Text1.Text = "":  Text2.Text = "": Exit Sub
'/\ checks to see if your guess is perfect :)
End Sub


Private Sub Command3_Click()
Beep
End
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
'the way this program works is you have to press command2
If KeyAscii = 13 Then ' waits for the user to press enter
Command2_Click ' clicks label2 for them
Beep ' beeps
End If 'ends the if statement
End Sub


