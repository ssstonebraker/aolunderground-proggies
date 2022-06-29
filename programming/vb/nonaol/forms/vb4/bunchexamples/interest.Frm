VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Interest Program"
   ClientHeight    =   2055
   ClientLeft      =   2460
   ClientTop       =   2010
   ClientWidth     =   3450
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   2460
   Left            =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   3450
   Top             =   1665
   Width           =   3570
   Begin VB.CommandButton Command1 
      Caption         =   "calculate"
      Height          =   225
      Left            =   2640
      TabIndex        =   6
      Top             =   1815
      Width           =   720
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   705
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1755
      TabIndex        =   2
      Text            =   "."
      Top             =   240
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label9 
      Height          =   210
      Left            =   15
      TabIndex        =   12
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label8 
      Caption         =   "monthly payment:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1605
      Width           =   1365
   End
   Begin VB.Label Label7 
      Height          =   180
      Left            =   15
      TabIndex        =   10
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label Label6 
      Caption         =   "total due:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   30
      TabIndex        =   9
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label Label5 
      Enabled         =   0   'False
      Height          =   255
      Left            =   15
      TabIndex        =   8
      Top             =   1005
      Width           =   1350
   End
   Begin VB.Label Label4 
      Caption         =   "interest due:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   15
      TabIndex        =   7
      Top             =   810
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "time(in years):"
      Height          =   255
      Left            =   2415
      TabIndex        =   4
      Top             =   75
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "rate:"
      Height          =   255
      Left            =   1770
      TabIndex        =   3
      Top             =   60
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "ammount borrowed:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   75
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
''Iz representing eXcel 2001
'this is a good way to learn vb, and some math for interest hehe
If Text1.Text = "" Then Exit Sub 'checks to see if your textboxes are blank
If Text2.Text = "" Then Exit Sub 'if they are, it wont do nothing
If Text3.Text = "" Then Exit Sub ' blah blah /\
Label4.Enabled = True ' enabled your labels \/
Label5.Enabled = True
Label7.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label9.Enabled = True
i = Val((Text1.Text) * (Text2.Text) * (Text3.Text))  'multiplys the numbers in your textboxes
total = i + Val(Text1.Text) ' adds text1 to the total you just got
mp = total / (12 * Val(Text3.Text)) ' multiplys 12(months) by the number in text3
Label5.Caption = "$" & i 'puts your totals in the labels \/
Label7.Caption = "$" & total
Label9.Caption = "$" & mp
End Sub


Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label5.Caption = ""
Label5.Enabled = False
Label4.Enabled = False

End Sub


