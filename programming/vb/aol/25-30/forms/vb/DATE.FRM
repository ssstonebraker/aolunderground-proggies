VERSION 2.00
Begin Form frmDate 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Choose A Date"
   ClientHeight    =   4740
   ClientLeft      =   1335
   ClientTop       =   2265
   ClientWidth     =   6585
   Height          =   5145
   Left            =   1275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6585
   Top             =   1920
   Width           =   6705
   Begin PictureBox Picture1 
      Height          =   4560
      Left            =   60
      ScaleHeight     =   4530
      ScaleWidth      =   6390
      TabIndex        =   0
      Top             =   75
      Width           =   6420
      Begin CommandButton cmdAccept 
         Caption         =   "Accept the chosen date"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   540
         Left            =   180
         TabIndex        =   4
         Top             =   3840
         Width           =   6105
      End
      Begin ListBox lstDates 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   17.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   3510
         Left            =   165
         TabIndex        =   1
         Top             =   210
         Width           =   3810
      End
      Begin Label lblDate 
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   510
         Left            =   4215
         TabIndex        =   3
         Top             =   1785
         Width           =   1950
      End
      Begin Label Label1 
         Caption         =   "Selected Date"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   4200
         TabIndex        =   2
         Top             =   1440
         Width           =   2025
      End
   End
End
Option Explicit

Sub cmdAccept_Click ()

    frmDate!Label1.Caption = lblDate.Caption
    Unload frmDate

End Sub

Sub Form_Load ()

    Dim nDays As Integer
    Dim varDate As Variant

    varDate = #1994-01-01#

    For nDays = 0 To 364

        varDate = DateAdd("d", 1, varDate)
        lstDates.AddItem Format(varDate, "Long Date")

    Next

End Sub

Sub lstDates_Click ()

    lblDate = Format(lstDates.Text, "Short Date")

End Sub

