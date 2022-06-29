VERSION 2.00
Begin Form frmMain 
   Caption         =   "Clock"
   ClientHeight    =   1065
   ClientLeft      =   1125
   ClientTop       =   2145
   ClientWidth     =   5535
   Height          =   1470
   Left            =   1065
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   5535
   Top             =   1800
   Width           =   5655
   Begin Timer timCLock 
      Interval        =   100
      Left            =   1695
      Top             =   435
   End
   Begin Label lblHour 
      Caption         =   "Label2"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   630
      Left            =   30
      TabIndex        =   1
      Top             =   360
      Width           =   5325
   End
   Begin Label lblMinute 
      Caption         =   "Label1"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   600
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5460
   End
End
Option Explicit

Sub timCLock_Timer ()

    lblMinute.Caption = "We are " & Minute(Now) & " minutes, " & Second(Now) & " seconds"
    lblHour.Caption = "into hour " & Hour(Now) & " of day " & Day(Now) & " of the month"

End Sub

