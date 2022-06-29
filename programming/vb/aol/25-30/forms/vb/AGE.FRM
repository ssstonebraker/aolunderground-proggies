VERSION 2.00
Begin Form Form1 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "User Login Verification"
   ClientHeight    =   2985
   ClientLeft      =   2400
   ClientTop       =   2130
   ClientWidth     =   3495
   ControlBox      =   0   'False
   Height          =   3390
   Left            =   2340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3495
   Top             =   1785
   Width           =   3615
   Begin CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   465
      Left            =   165
      TabIndex        =   4
      Top             =   2325
      Width           =   3150
   End
   Begin PictureBox Picture1 
      BackColor       =   &H00FFFF80&
      Height          =   1185
      Left            =   150
      ScaleHeight     =   1155
      ScaleWidth      =   3150
      TabIndex        =   2
      Top             =   990
      Width           =   3180
      Begin TextBox txtAge 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   570
         TabIndex        =   0
         Top             =   510
         Width           =   1920
      End
      Begin Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter your age, below..."
         Height          =   255
         Left            =   135
         TabIndex        =   3
         Top             =   150
         Width           =   2760
      End
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Age Check"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   22.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   660
      Left            =   285
      TabIndex        =   1
      Top             =   120
      Width           =   2925
   End
End
Option Explicit

Sub cmdDone_Click ()

    If txtAge = "" Then
        MsgBox "You must enter your age"
        txtAge.SetFocus
        Exit Sub
    ElseIf Val(txtAge) = 0 Then
        MsgBox "Enter your age in digits!"
        txtAge.SetFocus
        Exit Sub
    ElseIf Val(txtAge) < 18 Then
        MsgBox "You are not old enough to run this program"
        End
    Else
        MsgBox "Ok, you can run the program, but there's " & Chr$(10) & "             nothing else to see"
    End If

End Sub

Sub Form_Load ()

    form1.Left = (screen.Width - form1.Width) / 2
    form1.Top = (screen.Height - form1.Height) / 2

End Sub

