VERSION 2.00
Begin Form frmDetail 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Selected Contact"
   ClientHeight    =   2160
   ClientLeft      =   4575
   ClientTop       =   2565
   ClientWidth     =   3570
   Height          =   2565
   Left            =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   3570
   Top             =   2220
   Width           =   3690
   Begin PictureBox Picture1 
      Height          =   2040
      Left            =   75
      ScaleHeight     =   2010
      ScaleWidth      =   3405
      TabIndex        =   0
      Top             =   45
      Width           =   3435
      Begin Label lblSurname 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   180
         TabIndex        =   4
         Top             =   1380
         Width           =   3120
      End
      Begin Label Label3 
         Caption         =   "Surname"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   150
         TabIndex        =   3
         Top             =   1020
         Width           =   2460
      End
      Begin Label lblForename 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   150
         TabIndex        =   2
         Top             =   465
         Width           =   3135
      End
      Begin Label Label1 
         Caption         =   "Forename"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Top             =   120
         Width           =   2175
      End
   End
End
Option Explicit

