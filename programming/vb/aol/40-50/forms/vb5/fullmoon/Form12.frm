VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Color Coder"
   ClientHeight    =   1965
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   1965
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   11
      Text            =   "<FONT COLOR=#000000>"
      Top             =   480
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   135
      LargeChange     =   10
      Left            =   360
      Max             =   255
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      LargeChange     =   10
      Left            =   360
      Max             =   255
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   10
      Left            =   360
      Max             =   255
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "000"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "000"
      Top             =   1080
      Width           =   492
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "000"
      Top             =   720
      Width           =   492
   End
   Begin VB.PictureBox pbRGB 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   3240
      ScaleHeight     =   1155
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Color Coder"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
StayOnTop Me
Text1.Text = "<FONT COLOR=#" + Label1.Caption + Label2.Caption + Label3.Caption + ">"
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub HScroll1_Change()
Text2.Text = HScroll1.Value
pbRGB.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
If HScroll1.Value = 0 Then
Label3.Caption = "00"
Else
Label3.Caption = Hex(HScroll1.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label3.Caption + Label2.Caption + Label1.Caption + ">"
End Sub

Private Sub HScroll1_Scroll()
Text2.Text = HScroll1.Value
pbRGB.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
If HScroll1.Value = 0 Then
Label3.Caption = "00"
Else
Label3.Caption = Hex(HScroll1.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label3.Caption + Label2.Caption + Label1.Caption + ">"
End Sub


Private Sub HScroll2_Change()
Text3.Text = HScroll2.Value
pbRGB.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label1.Caption + Label2.Caption + Label3.Caption + ">"
If HScroll2.Value = 0 Then
Label2.Caption = "00"
Else
Label2.Caption = Hex(HScroll2.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label3.Caption + Label2.Caption + Label1.Caption + ">"
End Sub

Private Sub HScroll2_Scroll()
Text3.Text = HScroll2.Value
pbRGB.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label1.Caption + Label2.Caption + Label3.Caption + ">"
If HScroll2.Value = 0 Then
Label2.Caption = "00"
Else
Label2.Caption = Hex(HScroll2.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label3.Caption + Label2.Caption + Label1.Caption + ">"
End Sub


Private Sub HScroll3_Change()
Text4.Text = HScroll3.Value
pbRGB.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label1.Caption + Label2.Caption + Label3.Caption + ">"
If HScroll3.Value = 0 Then
Label1.Caption = "00"
Else
Label1.Caption = Hex(HScroll3.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label3.Caption + Label2.Caption + Label1.Caption + ">"
End Sub






Private Sub HScroll3_Scroll()
Text4.Text = HScroll3.Value
pbRGB.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label1.Caption + Label2.Caption + Label3.Caption + ">"
If HScroll3.Value = 0 Then
Label1.Caption = "00"
Else
Label1.Caption = Hex(HScroll3.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label3.Caption + Label2.Caption + Label1.Caption + ">"
End Sub

Private Sub Label1_Change()
If Len(Label1.Caption) = 1 Then
Label1.Caption = "0" + Label1.Caption
Else
End If
End Sub

Private Sub Label2_Change()
If Len(Label2.Caption) = 1 Then
Label2.Caption = "0" + Label2.Caption
Else
End If
End Sub

Private Sub Label3_Change()
If Len(Label3.Caption) = 1 Then
Label3.Caption = "0" + Label3.Caption
Else
End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub


Private Sub SSCommand1_Click()
Unload Me
End Sub

Private Sub SSCommand2_Click()
Me.WindowState = 1
Me.Caption = "CoLor Coder By UnderCross"
End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub Label5_Click()
Unload Me
End Sub

Private Sub Label6_Click()
Form12.WindowState = 1

End Sub

Private Sub Timer1_Timer()
Label7 = Time
End Sub
