VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Using the RGB() function"
   ClientHeight    =   4320
   ClientLeft      =   1740
   ClientTop       =   2265
   ClientWidth     =   5625
   Height          =   4725
   Left            =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5625
   Top             =   1920
   Width           =   5745
   Begin PictureBox Picture1 
      Height          =   3750
      Left            =   225
      ScaleHeight     =   3720
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   225
      Width           =   5145
      Begin PictureBox picColour 
         BackColor       =   &H00000000&
         Height          =   2775
         Left            =   1740
         ScaleHeight     =   2745
         ScaleWidth      =   3105
         TabIndex        =   5
         Top             =   405
         Width           =   3135
      End
      Begin VScrollBar vscrColours 
         Height          =   2760
         Index           =   2
         Left            =   1125
         Max             =   255
         TabIndex        =   3
         Top             =   390
         Width           =   270
      End
      Begin VScrollBar vscrColours 
         Height          =   2760
         Index           =   1
         Left            =   690
         Max             =   255
         TabIndex        =   2
         Top             =   390
         Width           =   270
      End
      Begin VScrollBar vscrColours 
         Height          =   2760
         Index           =   0
         Left            =   225
         Max             =   255
         TabIndex        =   1
         Top             =   390
         Width           =   270
      End
      Begin Label lblCode 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   255
         TabIndex        =   6
         Top             =   3300
         Width           =   4590
      End
      Begin Label Label1 
         Caption         =   "Red Green Blue"
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   90
         Width           =   1410
      End
   End
End
Option Explicit

Sub vscrColours_Change (Index As Integer)

    ' This code runs whenever any of the scrollbars changes. Since all the scrollbars belong to
    ' a control array (all have the same name, only the index property differs) the same piece
    ' of event code is used to respond to a change in all three.

    Dim sCode As String

    
    ' First the Values of the three scrollbars are fed into the RGB function to get a colour
    ' which is assigned to the picture box's BackColor property.
    picColour.BackColor = RGB(vscrColours(0).Value, vscrColours(1).Value, vscrColours(2).Value)


    ' A string variable called sCode is then built up to show the user the code that would
    ' actually be used in the program.
    sCode = "picColour.BackColor = RGB(" & vscrColours(0).Value & ", "
    sCode = sCode & vscrColours(1).Value & ", " & vscrColours(2).Value & ")"

    
    ' Finally this string is fed into the caption of the label at the foot of the form, to
    ' show the user the code that has just been run.
    lblCode.Caption = sCode

End Sub

