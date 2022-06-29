VERSION 2.00
Begin Form mmoptions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Mass Mail Options"
   ClientHeight    =   2040
   ClientLeft      =   1200
   ClientTop       =   2040
   ClientWidth     =   3720
   ControlBox      =   0   'False
   Height          =   2376
   Left            =   1152
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   3720
   Top             =   1752
   Width           =   3816
   Begin SSFrame Frame3D1 
      Alignment       =   2  'Center
      Caption         =   "Options"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin CommandButton Command1 
         Caption         =   "Save"
         FontBold        =   0   'False
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin CheckBox Check11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close Chatroom"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin CheckBox Check10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "45 Minute Kill"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2760
         TabIndex        =   10
         Text            =   "•UPS•"
         Top             =   240
         Width           =   495
      End
      Begin CheckBox Check9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start At Bottom"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin CheckBox Check8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BCC Names"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin CheckBox Check7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prefix:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin CheckBox Check6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "mIRC Stayontop"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin CheckBox Check5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alphabetical"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin CheckBox Check4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sign Off After"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin CheckBox Check3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kill Dupes"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remove Fwd:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Keep As New"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
   End
End
' MMOPTION.FRM
Option Explicit

Sub Command1_Click ()
Me.Hide
MMMain.Show
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
End Sub
