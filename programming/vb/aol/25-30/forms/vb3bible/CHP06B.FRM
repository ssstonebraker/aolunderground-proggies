VERSION 2.00
Begin Form formLetter 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6765
   ClientLeft      =   315
   ClientTop       =   300
   ClientWidth     =   6255
   ControlBox      =   0   'False
   Height          =   7200
   Left            =   240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6255
   Top             =   -60
   Width           =   6405
   Begin CommandButton cmndClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   315
      Left            =   4980
      TabIndex        =   3
      Top             =   6420
      Width           =   1095
   End
   Begin TextBox textPrintOut 
      BorderStyle     =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   5475
      Left            =   660
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   420
      Width           =   4935
   End
   Begin TextBox textPage 
      Height          =   6075
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   5775
   End
   Begin TextBox textShadow 
      BackColor       =   &H00000000&
      Height          =   6075
      Left            =   300
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   5775
   End
End

Sub cmndClose_Click ()
     formLetter.Hide              'Hide the letter display form
End Sub

