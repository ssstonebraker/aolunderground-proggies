VERSION 2.00
Begin Form Creating 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   1995
   ClientTop       =   720
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Height          =   1140
   Icon            =   CREATING.FRX:0000
   Left            =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   7215
   Top             =   375
   Width           =   7335
   Begin SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      Caption         =   "  Entering registration cert..."
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   0
      ShadowColor     =   1  'Black
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin SSCommand Command3D1 
         Caption         =   "Cancel"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   18
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5400
         RoundedCorners  =   0   'False
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
End
' CREATING.FRM
Option Explicit
Const c0026 =-1 ' &HFFFF%
Const c0028 = 3 ' &H3%

Sub Command3D1_Click ()
gv0072 = 1
Command3D1.Caption = "Wait..."
End Sub

Sub Form_Load ()
Dim l002C As Variant
l002C = extfn4120(Creating.hWnd, c0026, 0, 0, 0, 0, c0028)
Creating.Left = (Screen.Width - Creating.Width) / 2
Creating.Top = 0
End Sub
