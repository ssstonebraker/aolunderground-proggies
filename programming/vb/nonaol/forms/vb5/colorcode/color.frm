VERSION 5.00
Begin VB.Form Color 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Coder Example by Melon"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3225
   Icon            =   "color.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   -15.968
   ScaleMode       =   0  'User
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Text            =   "<FONT COLOR=#000000>"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   735
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
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This function was taken from Monkey-Fade Bas
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
End Function
Private Sub Form_Load()
'==========================================
'If you are going to put this on your prog please give me credit.
'Made with VB5 pro
'E-Mail = XMelon2000@Hotmail.com
'===========================================
End Sub

Private Sub HScroll1_Change()
Label1.Caption = HScroll1.Value
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
If HScroll1.Value = 0 Then
Label2.Caption = "00"
Else
Label2.Caption = Hex(HScroll1.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label5.Caption + Label4.Caption + Label2.Caption + ">"

End Sub

Private Sub HScroll2_Change()
Label3.Caption = HScroll2.Value
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label5.Caption + Label4.Caption + Label2.Caption + ">"
If HScroll2.Value = 0 Then
Label4.Caption = "00"
Else
Label4.Caption = Hex(HScroll2.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label5.Caption + Label4.Caption + Label2.Caption + ">"
End Sub

Private Sub HScroll3_Change()
Label6.Caption = HScroll3.Value
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label5.Caption + Label4.Caption + Label2.Caption + ">"
If HScroll3.Value = 0 Then
Label5.Caption = "00"
Else
Label5.Caption = Hex(HScroll3.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label5.Caption + Label4.Caption + Label2.Caption + ">"

End Sub

Private Sub Image1_Click()

End Sub
