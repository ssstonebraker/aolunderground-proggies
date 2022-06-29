VERSION 4.00
Begin VB.Form Form9 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Name"
   ClientHeight    =   840
   ClientLeft      =   3765
   ClientTop       =   4650
   ClientWidth     =   2490
   Height          =   1245
   Icon            =   "PROG10.frx":0000
   Left            =   3705
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   Top             =   4305
   Width           =   2610
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   195
      Index           =   1
      Left            =   1590
      TabIndex        =   2
      Top             =   360
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   360
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   600
      Width           =   2355
   End
End
Attribute VB_Name = "Form9"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
AOL% = FindWindow("AOL Frame25", vbNullString)
Call sendtext(AOL%, Text1.text)
End Sub


Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Changes AOL's Caption"
End Sub


Private Sub Command2_Click(Index As Integer)
Form9.Hide
Form4.Show
End Sub


Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Back to Main Form"
End Sub


Private Sub Form_Load()
Call StayOnTop(Form9)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Nutin here"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Yep, another box"
End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "What you wanna change it to"
End Sub


