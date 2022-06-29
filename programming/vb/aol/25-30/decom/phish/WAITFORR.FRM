VERSION 2.00
Begin Form WaitForReset 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1215
   ClientLeft      =   4275
   ClientTop       =   3120
   ClientWidth     =   3135
   ControlBox      =   0   'False
   Height          =   1620
   Icon            =   WAITFORR.FRX:0000
   Left            =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3135
   Top             =   2775
   Width           =   3255
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      Caption         =   "Resetting SN... Please Wait!"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      ShadowColor     =   1  'Black
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
' WAITFORR.FRM
Option Explicit
Const c001E =-1 ' &HFFFF%
Const c0020 = 3 ' &H3%

Sub Form_Load ()
Dim l0024 As Variant
l0024 = extfn4120(WaitForReset.hWnd, c001E, 0, 0, 0, 0, c0020)
WaitForReset.Left = (Screen.Width - WaitForReset.Width) / 2
WaitForReset.Top = (Screen.Height - WaitForReset.Height) / 2
End Sub
