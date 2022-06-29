VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Using Instr"
   ClientHeight    =   4110
   ClientLeft      =   2655
   ClientTop       =   2400
   ClientWidth     =   5280
   Height          =   4515
   Left            =   2595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5280
   Top             =   2055
   Width           =   5400
   Begin PictureBox Picture1 
      Height          =   3675
      Left            =   210
      ScaleHeight     =   3645
      ScaleWidth      =   4785
      TabIndex        =   3
      Top             =   240
      Width           =   4815
      Begin CommandButton cmdSearch 
         Caption         =   "Search"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   585
         Left            =   390
         TabIndex        =   2
         Top             =   2685
         Width           =   4005
      End
      Begin TextBox txtSearch 
         BackColor       =   &H00FFFF00&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   375
         TabIndex        =   1
         Top             =   1800
         Width           =   4050
      End
      Begin TextBox txtSource 
         BackColor       =   &H00FFFF00&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   360
         TabIndex        =   0
         Top             =   615
         Width           =   4050
      End
      Begin Label Label2 
         Caption         =   "String to search for"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Left            =   390
         TabIndex        =   5
         Top             =   1335
         Width           =   2865
      End
      Begin Label Label1 
         Caption         =   "String to search"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   390
         TabIndex        =   4
         Top             =   180
         Width           =   2445
      End
   End
End
Option Explicit

