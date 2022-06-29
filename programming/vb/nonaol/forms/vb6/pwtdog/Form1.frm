VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "password help by: t-dog [www.t-dog.org]"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "Enter"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   1215
      Left            =   1560
      Max             =   99
      Min             =   1
      TabIndex        =   21
      Top             =   2400
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   1215
      Left            =   1200
      Max             =   99
      Min             =   1
      TabIndex        =   20
      Top             =   2400
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1215
      Left            =   840
      Max             =   99
      Min             =   1
      TabIndex        =   19
      Top             =   2400
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   480
      Max             =   99
      Min             =   1
      TabIndex        =   18
      Top             =   2400
      Value           =   1
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Enter"
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "2"
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "9"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "8"
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "6"
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "5"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "0"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "•"
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Combonation Scroll"
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Number type"
      Height          =   495
      Left            =   3000
      TabIndex        =   27
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   2280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   495
      Left            =   1560
      TabIndex        =   25
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   375
      Left            =   1200
      TabIndex        =   24
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   0
      Y2              =   4200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User and Pass"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "password:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "user:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "tdog" And Text2.Text = "password here" Then
Form2.Show
Else
MsgBox "Password wrong check and try again", 16, "t-dog's password help"
End If
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command12_Click()
If Label4.Caption = "1498" Then
Form2.Show
Else
MsgBox "Password was wrong try again!", 16, "t-dog's password help"
End If
Label4.Caption = ""
End Sub

Private Sub Command10_Click()
Label4.Caption = "149"
End Sub

Private Sub Command13_Click()
If Label6.Caption = "1" And Label7.Caption = "4" And Label8.Caption = "9" And Label9.Caption = "8" Then
Form2.Show
Else
MsgBox "Number password wrong. please try again", 16, "t-dog's password help"
End If
Label6.Caption = "0"
Label7.Caption = "0"
Label8.Caption = "0"
Label9.Caption = "0"
End Sub

Private Sub Command3_Click()
Label4.Caption = "1"
End Sub

Private Sub Command5_Click()
Label4.Caption = "14"
End Sub

Private Sub Command9_Click()
Label4.Caption = "1498"
End Sub

Private Sub VScroll1_Change()
Label6.Caption = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

Private Sub VScroll2_Change()
Label7.Caption = VScroll2.Value
End Sub

Private Sub VScroll3_Change()
Label8.Caption = VScroll3.Value
End Sub

Private Sub VScroll4_Change()
Label9.Caption = VScroll4.Value
End Sub
