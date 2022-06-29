VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "clock"
   ClientHeight    =   1560
   ClientLeft      =   2295
   ClientTop       =   3885
   ClientWidth     =   3390
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1965
   Left            =   2235
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   3390
   Top             =   3540
   Width           =   3510
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   2760
      TabIndex        =   6
      Text            =   "AM"
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   840
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "off"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "set"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   240
      Top             =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "alarm time:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer2.Enabled = True ' starts timer2 [i advise that you check the properties to timer2
'to learn more about timers
Command1.Visible = False 'hides the start button
Command2.Visible = True ' shows the stop button
'now the program is running and ready to be shut off
End Sub

Private Sub Command2_Click()
Timer2.Enabled = False ' stops timer2from doin anything
Command1.Visible = True ' hides the off button
Command2.Visible = False ' shows the start button
'now its ready to be set again ;D
End Sub


Private Sub Form_Load()
'Iz representing eXcel 2001
Timer1.Enabled = True ' starts timer1 when the program loads
Label2.Caption = Date ' displays the date in a label box
Combo1.AddItem ("PM") ' adds PM to a combo box
Combo1.AddItem ("AM") ' adds AM to a combo box
End Sub


Private Sub Timer1_Timer()
Label1.Caption = Time
'Time is a code in VB that will get the time from your Systray box way down there to the right
End Sub


Private Sub Timer2_Timer()
If Text1.Text = Label1.Caption & " " & Combo1.Text Then ' tells vb that if
'the time you set your alarm for comes up, then beep like a motherfucker
Beep ' beeps
Beep
Beep
Beep
Beep
Beep
Beep
Beep
Beep
Beep
Beep
Beep
End If ' ends your if statement
End Sub


