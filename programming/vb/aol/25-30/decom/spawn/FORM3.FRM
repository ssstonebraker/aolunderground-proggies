VERSION 2.00
Begin Form Form3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   " Password..."
   ClientHeight    =   615
   ClientLeft      =   3660
   ClientTop       =   3225
   ClientWidth     =   1710
   ControlBox      =   0   'False
   Height          =   1020
   Left            =   3600
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   1710
   Top             =   2880
   Width           =   1830
   Begin SSCommand Command3D1 
      Caption         =   "Enter Spawn"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin SSPanel Panel3D1 
      AutoSize        =   3  'AutoSize Child To Panel
      BackColor       =   &H00FF0000&
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      Caption         =   "Panel3D1"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ShadowColor     =   1  'Black
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin TextBox Text1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   45
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   45
         Width           =   1605
      End
   End
End
