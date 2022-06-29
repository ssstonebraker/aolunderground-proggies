VERSION 2.00
Begin Form Form5 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "EXeCuTioN Away Bot"
   ClientHeight    =   1095
   ClientLeft      =   3120
   ClientTop       =   1890
   ClientWidth     =   2550
   ControlBox      =   0   'False
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Arial"
   FontSize        =   12
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   1500
   Left            =   3060
   LinkTopic       =   "Form5"
   ScaleHeight     =   1095
   ScaleWidth      =   2550
   Top             =   1545
   Width           =   2670
   Begin SSPanel Panel3D4 
      BevelOuter      =   0  'None
      BevelWidth      =   2
      BorderWidth     =   1
      Caption         =   "Seconds:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin SSPanel Panel3D3 
      BevelOuter      =   0  'None
      BevelWidth      =   2
      BorderWidth     =   1
      Caption         =   "60"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   1320
   End
   Begin SSCommand Command3D2 
      Caption         =   "Exit"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin SSCommand Command3D1 
      Caption         =   "Start aF|ã"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin SSPanel Panel3D2 
      BevelOuter      =   0  'None
      BevelWidth      =   2
      BorderWidth     =   1
      Caption         =   "Gone For:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   610
      Width           =   975
   End
   Begin SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelOuter      =   0  'None
      BevelWidth      =   2
      BorderWidth     =   1
      Caption         =   "0"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   615
      Width           =   735
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Reason Gone:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin TextBox Text1 
         BackColor       =   &H00000000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Having Phone Sex"
         Top             =   240
         Width           =   2295
      End
   End
End
' AWAY.FRM
Option Explicit

Sub Command3D1_Click ()
If  Text1.Text = "" Then MsgBox "Whats the Reason?", 16, "Error": Exit Sub
If  Command3D1.Caption = "Start aF|ã" Then
Command3D1.Caption = "Stop"
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Currently aF|ã - Gone Less Than a Minute§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Reason: " + Text1 + "§]v^ñï")
Text1.Enabled = False
Panel3D1.Caption = "0"
Timer1.Enabled = True
Exit Sub
End If
If  Command3D1.Caption = "Stop" Then
Command3D1.Caption = "Start aF|ã"
Text1.Enabled = True
Timer1.Enabled = False
Panel3D1.Caption = "0"
Panel3D3.Caption = "0"
Call subBF0("ïñ^v[§I am Not aF|ã AnyMore§]v^ñï")
End If
End Sub

Sub Command3D2_Click ()
Me.Hide
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub Timer1_Timer ()
If  Panel3D3.Caption = "0" Then
Panel3D3.Caption = "60"
Panel3D1.Caption = Panel3D1.Caption + 1
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Currently aF|ã - Gone " + Panel3D1 + " Min(s)§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Reason: " + Text1 + "§]v^ñï")
End If
Panel3D3.Caption = Panel3D3.Caption - 1
If  Text1.Enabled = True Then
Timer1.Enabled = False
Exit Sub
End If
subD08 1
End Sub
