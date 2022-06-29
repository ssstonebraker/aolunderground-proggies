VERSION 2.00
Begin Form Form2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1200
   ClientLeft      =   3390
   ClientTop       =   2070
   ClientWidth     =   2385
   ForeColor       =   &H00C0C0C0&
   Height          =   1605
   Left            =   3330
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2385
   Top             =   1725
   Width           =   2505
   Begin SSCommand Command3D1 
      Caption         =   "Enter Execution"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2100
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BevelWidth      =   3
      BorderWidth     =   4
      Caption         =   "Panel3D1"
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
      Begin TextBox Text1 
         BackColor       =   &H00000000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Wingdings"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   100
         MaxLength       =   15
         PasswordChar    =   "l"
         TabIndex        =   2
         Top             =   100
         Width           =   1950
      End
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
' PW.FRM
Option Explicit

Sub Command3D1_Click ()
If  Text1.Text = "" Then MsgBox "Type in a Password", 16, "Error": Exit Sub
If  UCase$(Text1) = "YUBAN OWNS ME" Then
Unload Me
TOSer.Show
Else
MsgBox "Sorry You Bastard.  Maybe Next Time.", 16, "Execution": Exit Sub
Text1.Text = ""
End If
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub Text1_KeyPress (KeyAscii As Integer)
If  KeyAscii = 13 Then
If  Text1.Text = "" Then MsgBox "Type in a Password", 16, "Error": Exit Sub
If  UCase$(Text1) = "YUBAN OWNS ME" Then
Unload Me
TOSer.Show
Else
MsgBox "Sorry You Bastard.  Maybe Next Time.", 16, "Execution": Exit Sub
Text1.Text = ""
End If
End If
End Sub
