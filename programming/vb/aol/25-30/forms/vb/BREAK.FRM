VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Left$ and Right$"
   ClientHeight    =   3330
   ClientLeft      =   2580
   ClientTop       =   2655
   ClientWidth     =   7365
   Height          =   3735
   Left            =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   7365
   Top             =   2310
   Width           =   7485
   Begin PictureBox Picture1 
      Height          =   3045
      Left            =   180
      ScaleHeight     =   3015
      ScaleWidth      =   6960
      TabIndex        =   1
      Top             =   165
      Width           =   6990
      Begin CommandButton cmdBreak 
         Caption         =   "Break"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   510
         Left            =   2415
         TabIndex        =   7
         Top             =   2325
         Width           =   2010
      End
      Begin TextBox txtName 
         BackColor       =   &H00FFFF00&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   270
         TabIndex        =   0
         Top             =   540
         Width           =   6090
      End
      Begin Label lblSurname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   1875
         TabIndex        =   6
         Top             =   1695
         Width           =   4485
      End
      Begin Label lblForename 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   1860
         TabIndex        =   5
         Top             =   1170
         Width           =   4485
      End
      Begin Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Surname"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   345
         TabIndex        =   4
         Top             =   1740
         Width           =   1410
      End
      Begin Label Label2 
         Caption         =   "Forename"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   315
         TabIndex        =   3
         Top             =   1275
         Width           =   1410
      End
      Begin Label Label1 
         Caption         =   "Please enter your full name (forename and surname) below"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   270
         TabIndex        =   2
         Top             =   225
         Width           =   6420
      End
   End
End
Option Explicit

