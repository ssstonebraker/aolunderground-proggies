VERSION 2.00
Begin Form formModal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modal Form"
   ClientHeight    =   2160
   ClientLeft      =   555
   ClientTop       =   2115
   ClientWidth     =   4335
   ControlBox      =   0   'False
   Height          =   2595
   Left            =   480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4335
   Top             =   1755
   Width           =   4485
   Begin TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   912
      Left            =   660
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "This is a modal form. Try clicking on any other form in this project... the focus does not shift. Also notice that the MsgBox hasn't popped up yet."
      Top             =   120
      Width           =   3072
   End
   Begin CommandButton cmndOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   492
      Left            =   1320
      TabIndex        =   0
      Top             =   1260
      Width           =   1752
   End
End

Sub cmndOK_Click ()
     Hide
End Sub

