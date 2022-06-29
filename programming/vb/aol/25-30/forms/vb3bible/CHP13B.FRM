VERSION 2.00
Begin Form formPassword 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Enter Password"
   ClientHeight    =   624
   ClientLeft      =   1824
   ClientTop       =   3360
   ClientWidth     =   3612
   Height          =   1068
   Left            =   1764
   LinkTopic       =   "Form2"
   ScaleHeight     =   624
   ScaleWidth      =   3612
   Top             =   2976
   Width           =   3732
   Begin CommandButton cmndOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   432
      Left            =   3108
      TabIndex        =   1
      Top             =   120
      Width           =   372
   End
   Begin TextBox textPassword 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   420
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3012
   End
End
Sub cmndOK_Click ()
     formPassword.Hide
End Sub

