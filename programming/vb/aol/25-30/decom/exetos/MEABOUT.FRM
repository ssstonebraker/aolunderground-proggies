VERSION 2.00
Begin Form Form7 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "About EXeCuTioN:"
   ClientHeight    =   2340
   ClientLeft      =   2580
   ClientTop       =   1995
   ClientWidth     =   4020
   ControlBox      =   0   'False
   Height          =   2745
   Left            =   2520
   LinkTopic       =   "Form7"
   ScaleHeight     =   2340
   ScaleWidth      =   4020
   Top             =   1650
   Width           =   4140
   Begin SSCommand Command3D2 
      BevelWidth      =   3
      Caption         =   "Let Me See Da Greets"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin SSPanel Panel3D3 
      Alignment       =   6  'Center - TOP
      BevelInner      =   1  'Inset
      BevelWidth      =   2
      BorderWidth     =   1
      Caption         =   "© YuﬂaN Enterprises, 1997"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   400
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3855
   End
   Begin SSPanel Panel3D2 
      Alignment       =   6  'Center - TOP
      BevelInner      =   1  'Inset
      BevelWidth      =   2
      BorderWidth     =   1
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
      Begin SSCommand Command3D1 
         BevelWidth      =   3
         Caption         =   "Okay Thats Nice"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
   End
   Begin SSPanel Panel3D1 
      Alignment       =   6  'Center - TOP
      BevelInner      =   2  'Raised
      BevelWidth      =   3
      BorderWidth     =   4
      Caption         =   "    ïñ^v[§EXeCuTioNπ∑∫ TOSer§]v^ñï                   ïñ^v[§Version 1.o§]v^ñï             ïñv^[§ﬂy YuﬂaN§]v^ñï"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3735
   End
   Begin Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
' MEABOUT.FRM
Option Explicit

Sub Command3D1_Click ()
subD08 .3#
Unload Me
End Sub

Sub Command3D2_Click ()
subD08 .5#
Unload Me
Form72.Show
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub Panel3D1_DragDrop (Source As Control, X As Single, Y As Single)
TOSer.mnuespec.Enabled = True
MsgBox "You Can Use the E-Special TOS Method", 77, "Yay!"
End Sub
