VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "no scroll bar fader example by burden"
   ClientHeight    =   3165
   ClientLeft      =   2925
   ClientTop       =   2865
   ClientWidth     =   4275
   Height          =   3570
   Left            =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4275
   Top             =   2520
   Width           =   4395
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2160
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   4095
      TabIndex        =   10
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   4095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Fade Normal"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fade Wavy"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color Control"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command3 
         Caption         =   "Color 3"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Color 2"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Color 1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
'just in case.
On Error GoTo burden
    
    'call the ocx's 3rd action
    CommonDialog1.Action = 3
    'the color you chose becomes the back color
    Label1.BackColor = CommonDialog1.Color

'lets hope it didn't reach this part /=^)
burden:
    Exit Sub
End Sub

Private Sub Command2_Click()
'just in case.
On Error GoTo burden
    
    'call the ocx's 3rd action
    CommonDialog1.Action = 3
    'the color you chose becomes the back color
    Label2.BackColor = CommonDialog1.Color

'lets hope it didn't reach this part /=^)
burden:
    Exit Sub
End Sub


Private Sub Command3_Click()
'just in case.
On Error GoTo burden
    
    'call the ocx's 3rd action
    CommonDialog1.Action = 3
    'the color you chose becomes the back color
    Label3.BackColor = CommonDialog1.Color

'lets hope it didn't reach this part /=^)
burden:
    Exit Sub
End Sub

Private Sub Command4_Click()
'this is just a error message for the user
If Text1.Text = "" Then MsgBox "your gonna need to write something in order for it to fade", 64, "Error:" Else

'this tells it to fade the text in text1----------------------------------------------------[..]--this makes the text wavy
FadedText$ = FadeByColor3(Label1.BackColor, Label2.BackColor, Label3.BackColor, Text1.Text, True)
'this says that text2 is supposed to have the HTML
Text2 = FadedText$
'this part tells the picturebox to read the html in text2
Call FadePreview(Picture1, Text2.Text)
End Sub

Private Sub Command5_Click()
'this is just a error message for the user
If Text1.Text = "" Then MsgBox "your gonna need to write something in order for it to fade", 64, "Error:" Else

'this tells it to fade the text in text1----------------------------------------------------[..]--this makes the text normal
FadedText$ = FadeByColor3(Label1.BackColor, Label2.BackColor, Label3.BackColor, Text1.Text, False)
'this says that text2 is supposed to have the HTML
Text2 = FadedText$
'this part tells the picturebox to read the html in text2
Call FadePreview(Picture1, Text2.Text)
End Sub

Private Sub Command6_Click()
Call IMKeyword(" + Text3.Text + ", " + Text2.Text + ")
End Sub

