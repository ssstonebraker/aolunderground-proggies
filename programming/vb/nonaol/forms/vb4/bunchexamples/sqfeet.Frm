VERSION 4.00
Begin VB.Form Form1 
   ClientHeight    =   1320
   ClientLeft      =   3240
   ClientTop       =   4260
   ClientWidth     =   4020
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1725
   Left            =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   4020
   Top             =   3915
   Width           =   4140
   Begin VB.CommandButton Command4 
      Caption         =   "help"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "calculate"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "sq. feet: 0"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "width"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "height"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "legnth"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line12 
      X1              =   3960
      X2              =   3960
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      X1              =   2640
      X2              =   2640
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line10 
      X1              =   2640
      X2              =   3960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line9 
      X1              =   2640
      X2              =   3960
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line8 
      X1              =   3600
      X2              =   3960
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FF00&
      X1              =   2280
      X2              =   2640
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Line Line6 
      X1              =   3600
      X2              =   3960
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   2280
      X2              =   2640
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line4 
      X1              =   3600
      X2              =   3600
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line3 
      X1              =   2280
      X2              =   3600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2280
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   3600
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label3 
      Caption         =   "height"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "width:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "legnth:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Iz representing eXcel 2001
a = Val(Text1.Text) ' \/ gets the values if your textbox "checks for the numbers in them :)
b = Val(Text2.Text)
c = Val(Text3.Text)
answer = (2 * a * c + 2 * l * b + 2 * b * c) 'stupid math formula to ge tthe size of a box
'in square feet
Label7.Caption = "sq. feet: " & answer ' displays the size of the box in square feet
End Sub

Private Sub Command2_Click()
Text1.Text = "" '\/ clears your textboxes
Text2.Text = ""
Text3.Text = ""
Label7.Caption = "sq. feet: 0" ' clears your label
End Sub


Private Sub Command3_Click()
End ' exits your program
End Sub

Private Sub Command4_Click()
MsgBox "this program was created to give you the number of square feet in a room by using your own mesurements"
'/\ mesage box
End Sub

