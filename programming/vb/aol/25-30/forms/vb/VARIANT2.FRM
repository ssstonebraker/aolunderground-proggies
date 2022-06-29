VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Checking Variants"
   ClientHeight    =   1680
   ClientLeft      =   2175
   ClientTop       =   2280
   ClientWidth     =   6405
   Height          =   2085
   Left            =   2115
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   6405
   Top             =   1935
   Width           =   6525
   Begin PictureBox Picture1 
      Height          =   1410
      Left            =   210
      ScaleHeight     =   1380
      ScaleWidth      =   6030
      TabIndex        =   1
      Top             =   105
      Width           =   6060
      Begin TextBox txtValue 
         Height          =   300
         Left            =   2910
         TabIndex        =   0
         Top             =   165
         Width           =   2925
      End
      Begin Label lblType 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   420
         Left            =   180
         TabIndex        =   3
         Top             =   690
         Width           =   5640
      End
      Begin Label Label1 
         Caption         =   "Enter your value here"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   195
         TabIndex        =   2
         Top             =   135
         Width           =   2775
      End
   End
End
Option Explicit

Sub txtValue_Change ()

    Dim varValue As Variant

    If IsDate(txtValue.Text) Then
        varValue = CVDate(txtValue.Text)
    
    ElseIf IsNumeric(txtValue.Text) Then
        varValue = Val(txtValue.Text)

    Else
        varValue = txtValue.Text

    End If

    Select Case VarType(varValue)

        Case 5
            lblType.Caption = "Type 5 - Double"
        Case 7
            lblType.Caption = "Type 7 - Date"
        Case 8
            lblType.Caption = "Type 8 - String"

    End Select

End Sub

