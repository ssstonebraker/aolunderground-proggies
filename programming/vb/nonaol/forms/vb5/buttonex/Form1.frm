VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Option button example Made By Gorila"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Make RED"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Make BLUE"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "end"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visit my web page, fly.to/yippie"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Check1_Click()
Form1.BackColor = &HFF&
Option1.BackColor = &HFF&
Option2.BackColor = &HFF&
End Sub

Private Sub Label1_Click()
Call ShellExecute(hwnd, "Open", "http://members.tripod.com/origal/index.html/", "", App.Path, 1)
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Option1_Click()
'simply give this command to say commands, change the colors

Form1.BackColor = &HFF0000
Option1.BackColor = &HFF0000
Option2.BackColor = &HFF0000
Label1.BackColor = &HFF0000
Label2.BackColor = &HFF0000
End Sub

Private Sub Option2_Click()
'this says what back colors to asign, chang the 7hff& to any color you want
'try a checkbox, somewhat different
Form1.BackColor = &HFF&
Option1.BackColor = &HFF&
Option2.BackColor = &HFF&
Label1.BackColor = &HFF&
Label2.BackColor = &HFF&
End Sub
