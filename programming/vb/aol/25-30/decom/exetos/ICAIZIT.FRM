VERSION 2.00
Begin Form Form102 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "YuBaN Checker:"
   ClientHeight    =   1425
   ClientLeft      =   3735
   ClientTop       =   2220
   ClientWidth     =   2640
   ControlBox      =   0   'False
   Height          =   1830
   Left            =   3675
   LinkTopic       =   "Form10"
   ScaleHeight     =   1425
   ScaleWidth      =   2640
   Top             =   1875
   Width           =   2760
   Begin CommandButton Command2 
      Caption         =   "Nah...."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin CommandButton Command1 
      Caption         =   "Check String"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin TextBox Text1 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "String Here!"
      Top             =   360
      Width           =   2055
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check if it is Realy YuBaN!:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
' ICAIZIT.FRM
Option Explicit

Sub Command1_Click ()
If  Text1.Text = "" Then MsgBox "Type in Da String!", 16, "Error": Exit Sub
If  UCase$(Text1.Text) = UCase$("alphazulu136420") Then
MsgBox "Thats YuBaN!!!", 77, "YippY!"
Call subBF0("†     ïñ^v[§A.T.T.E.N.T.I.O.N§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§YuﬂaN is Online Right Now!§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Oh My God!..I Gotta Go Suck on his Dick§]v^ñï")
subD08 1E-100#
Call subBF0("†     ïñ^v[§A.T.T.E.N.T.I.O.N§]v^ñï")
subD08 1E-100#
subD08 .3#
Unload Me
Else
MsgBox "Wrong!", 16, "TOS THAT POZER ASS!!!"
End If
End Sub

Sub Command2_Click ()
Unload Me
End Sub

Sub Form_Load ()
subCD0 Me
End Sub
