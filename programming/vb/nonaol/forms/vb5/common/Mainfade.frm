VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "                        Progressive Fade"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   Icon            =   "Mainfade.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Options"
      Height          =   285
      Left            =   4200
      TabIndex        =   11
      Top             =   1080
      Width           =   765
   End
   Begin VB.PictureBox Pic2 
      Height          =   315
      Left            =   0
      MousePointer    =   15  'Size All
      ScaleHeight     =   255
      ScaleWidth      =   4410
      TabIndex        =   10
      Top             =   0
      Width           =   4470
   End
   Begin VB.CheckBox Option1 
      Caption         =   "Wavy"
      Height          =   285
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   630
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6930
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6930
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color2"
      Height          =   270
      Left            =   1785
      TabIndex        =   5
      Top             =   1080
      Width           =   630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color1"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   630
   End
   Begin VB.TextBox Text2 
      Height          =   540
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2415
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   5070
   End
   Begin VB.PictureBox pic1 
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   4785
      TabIndex        =   1
      Top             =   735
      Width           =   4845
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   405
      Width           =   4845
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   345
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Example By: SuNGoD"
      Top             =   435
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6945
      Top             =   5085
   End
   Begin VB.Label Label3 
      Height          =   870
      Left            =   6690
      TabIndex        =   13
      Top             =   4755
      Width           =   1050
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4770
      TabIndex        =   9
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4470
      TabIndex        =   8
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Error_Event:
CommonDialog1.ShowColor
Label1.BackColor = CommonDialog1.Color
Error_Event:
    Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo Error_Event:
CommonDialog2.ShowColor
Label2.BackColor = CommonDialog2.Color
Error_Event:
    Exit Sub
End Sub

Private Sub Command3_Click()
Form2.PopupMenu Form2.Options, 1
End Sub

Private Sub Form_Paint()
FadeFormBlue Me
End Sub

Private Sub Label4_Click()
Form1.WindowState = 1
End Sub

Private Sub Label5_Click()
End
End Sub

Private Sub Pic2_Click()
Call MoveForm(Form1)
End Sub

Private Sub Timer1_Timer()
If Option1.Value = False Then
Text2 = "<B>" & FadeByColor2(Label1.BackColor, Label2.BackColor, Text1, False)
Call FadePreview(pic1, Text2)
Else
Text2 = "<B>" & FadeByColor2(Label1.BackColor, Label2.BackColor, Text1, True)
Call FadePreview(pic1, Text2)
End If
End Sub

Private Sub Timer2_Timer()
Text6 = "<B>" & FadeByColor2(Label1.BackColor, Label2.BackColor, "                       Progressive Fade", False)
Call FadePreview(Pic2, Text3)
End Sub

Private Sub Timer3_Timer()

End Sub
