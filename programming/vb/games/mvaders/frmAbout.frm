VERSION 2.00
Begin Form frmAbout 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "MVaders - About the program."
   ClientHeight    =   4455
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
   Height          =   4860
   Icon            =   FRMABOUT.FRX:0000
   Left            =   1035
   LinkTopic       =   "Form2"
   ScaleHeight     =   4455
   ScaleWidth      =   7365
   Top             =   1140
   Width           =   7485
   Begin CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   435
      Left            =   3060
      TabIndex        =   1
      Top             =   3900
      Width           =   1155
   End
   Begin Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "MVaders"
      Height          =   3495
      Left            =   420
      TabIndex        =   0
      Top             =   240
      Width           =   6435
      WordWrap        =   -1  'True
   End
End
Option Explicit

Sub cmdOk_Click ()

Unload frmAbout

End Sub

Sub Form_Activate ()

Dim sMsg As String

'Build and display programming information
sMsg = "MVaders"
sMsg = sMsg + Chr$(10) + ""
sMsg = sMsg + Chr$(10) + "Version: v1.0"
sMsg = sMsg + Chr$(10) + "Date: August 1997"
sMsg = sMsg + Chr$(10) + "Programmer: Mark Meany"
sMsg = sMsg + Chr$(10) + "Language: Visual Basic v3"
sMsg = sMsg + Chr$(10) + "Graphics: Ari Feldman, from SpriteLib"
sMsg = sMsg + Chr$(10) + ""
sMsg = sMsg + Chr$(10) + "The source code for this and others games along with"
sMsg = sMsg + Chr$(10) + "tips and tutorials can be found on my web"
sMsg = sMsg + Chr$(10) + "site:"
sMsg = sMsg + Chr$(10) + ""
sMsg = sMsg + Chr$(10) + "http://www.geocities.com/SiliconValley/Bay/9520"
sMsg = sMsg + Chr$(10) + ""
sMsg = sMsg + Chr$(10) + "Enjoy!"
lblAbout = sMsg

End Sub

Sub Form_Load ()

'Make sure we are central
CenterForm Me

End Sub

