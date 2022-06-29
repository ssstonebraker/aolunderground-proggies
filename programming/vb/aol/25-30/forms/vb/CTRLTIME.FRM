VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Timing explicit and implicit object variables"
   ClientHeight    =   2400
   ClientLeft      =   1290
   ClientTop       =   2535
   ClientWidth     =   6525
   Height          =   2805
   Left            =   1230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   6525
   Top             =   2190
   Width           =   6645
   Begin PictureBox Picture1 
      Height          =   1905
      Left            =   480
      ScaleHeight     =   1875
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   210
      Width           =   5565
      Begin CommandButton cmdImplicit 
         Caption         =   "&Implicit"
         Height          =   660
         Left            =   3225
         TabIndex        =   4
         Top             =   1020
         Width           =   2010
      End
      Begin CommandButton cmdExplicit 
         Caption         =   "&Explicit"
         Height          =   660
         Left            =   135
         TabIndex        =   3
         Top             =   1005
         Width           =   2010
      End
      Begin Label lblTime 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   15
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   585
         Left            =   2055
         TabIndex        =   2
         Top             =   180
         Width           =   3180
      End
      Begin Label Label1 
         Caption         =   "Time Taken"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   210
         TabIndex        =   1
         Top             =   255
         Width           =   1815
      End
   End
End
Option Explicit

Sub cmdExplicit_Click ()

    Dim varTime As Variant
    Dim nIndex As Integer

    varTime = Now

    For nIndex = 1 To 1000

        Time_Explicit cmdExplicit, nIndex

    Next

    cmdExplicit.Caption = "&Explicit"

    varTime = Second(Now - varTime)

   lblTime.Caption = varTime & " Seconds"

End Sub

Sub cmdImplicit_Click ()
    
    Dim varTime As Variant
    Dim nIndex As Integer

    varTime = Now

    For nIndex = 1 To 1000

        Time_Implicit cmdImplicit, nIndex

    Next

    cmdImplicit.Caption = "&Implicit"

    varTime = Second(Now - varTime)

   lblTime.Caption = varTime & " Seconds"

End Sub

Sub Time_Explicit (cmdCommand As CommandButton, nNumber As Integer)

    cmdCommand.Caption = nNumber

End Sub

Sub Time_Implicit (cmdCommand As Control, nNumber As Integer)

    cmdCommand.Caption = nNumber

End Sub

