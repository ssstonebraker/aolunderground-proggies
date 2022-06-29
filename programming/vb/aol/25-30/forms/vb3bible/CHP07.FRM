VERSION 2.00
Begin Form Form1 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Change Project"
   ClientHeight    =   1920
   ClientLeft      =   1170
   ClientTop       =   1965
   ClientWidth     =   3690
   Height          =   2355
   Icon            =   CHP07.FRX:0000
   Left            =   1095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3690
   Top             =   1605
   Width           =   3840
   Begin TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   1692
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   3696
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   3732
   End
End

Sub Text1_Change ()
    Dim CharCount As Integer
    Dim countMsg As String
    CharCount = Len(Text1.Text)
    countMsg = "characters"
    Select Case CharCount
    Case 0
        countMsg = " no characters"
    Case 1
        countMsg = " one character"
    Case Else
        countMsg = Format$(CharCount, " ### ")
        countMsg = countMsg + "characters"
    End Select
    Label1.Caption = countMsg
End Sub

