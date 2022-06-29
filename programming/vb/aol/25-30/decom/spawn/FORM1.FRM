VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   " Spawn Punter 1.0"
   ClientHeight    =   2175
   ClientLeft      =   2910
   ClientTop       =   2640
   ClientWidth     =   3015
   ControlBox      =   0   'False
   Height          =   2865
   Icon            =   FORM1.FRX:0000
   Left            =   2850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   FORM1.FRX:0302
   ScaleHeight     =   2175
   ScaleWidth      =   3015
   Top             =   2010
   Width           =   3135
   Begin SSPanel Panel3D1 
      AutoSize        =   3  'AutoSize Child To Panel
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin SSCommand Command3D1 
         Caption         =   "Start Punt"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin SSCommand Command3D9 
         Caption         =   "STOP Punt"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin SSCommand Command3D8 
         Caption         =   "On"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
      Begin SSPanel Panel3D1 
         AutoSize        =   3  'AutoSize Child To Panel
         BackColor       =   &H00FF0000&
         BevelInner      =   1  'Inset
         BorderWidth     =   1
         Caption         =   "Panel3D1"
         Height          =   420
         Index           =   1
         Left            =   120
         ShadowColor     =   1  'Black
         TabIndex        =   9
         Top             =   120
         Width           =   1335
         Begin ComboBox Combo1 
            BackColor       =   &H00000000&
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   45
            TabIndex        =   10
            Text            =   "Person"
            Top             =   45
            Width           =   1245
         End
      End
      Begin SSPanel Panel3D2 
         AutoSize        =   3  'AutoSize Child To Panel
         BackColor       =   &H00FF0000&
         BevelInner      =   1  'Inset
         BorderWidth     =   1
         Caption         =   "Panel3D2"
         Height          =   1170
         Left            =   1560
         ShadowColor     =   1  'Black
         TabIndex        =   7
         Top             =   120
         Width           =   1335
         Begin ListBox List1 
            BackColor       =   &H00000000&
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   1080
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   1245
         End
      End
      Begin SSCommand Command3D2 
         Caption         =   "Clear"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin SSCommand Command3D7 
         Caption         =   "Spawn Status"
         FontBold        =   0   'False
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   2775
      End
      Begin SSCommand Command3D4 
         Caption         =   "Fake Punt"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin SSCommand Command3D5 
         Caption         =   "Minimize"
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
         Top             =   1080
         Width           =   1335
      End
      Begin SSCommand Command3D6 
         Caption         =   "Exit Punter"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin CommonDialog CMDialog1 
         Left            =   1680
         Top             =   2160
      End
      Begin SSCommand Command3D3 
         Caption         =   "-"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   1320
         Width           =   255
      End
   End
   Begin Menu menu1 
      Caption         =   "&IMs"
      Begin Menu im1 
         Caption         =   "IMs On"
      End
      Begin Menu im2 
         Caption         =   "IMs Off"
      End
   End
   Begin Menu menu2 
      Caption         =   "&Punted"
      Begin Menu punt9 
         Caption         =   "Online?"
      End
      Begin Menu puntbar 
         Caption         =   "-"
      End
      Begin Menu Punt4 
         Caption         =   "Load List"
      End
      Begin Menu Punt3 
         Caption         =   "Save List"
      End
      Begin Menu Punt5 
         Caption         =   "Clear List"
      End
      Begin Menu punt6 
         Caption         =   "-"
      End
      Begin Menu Punt7 
         Caption         =   "Add Room"
      End
      Begin Menu Punt8 
         Caption         =   "Add Name"
      End
   End
   Begin Menu menu3 
      Caption         =   "&Brags"
      Begin Menu Brag1 
         Caption         =   "Brag All"
      End
      Begin Menu brag4 
         Caption         =   "-"
      End
      Begin Menu brag2 
         Caption         =   "Brag One"
      End
      Begin Menu brag3 
         Caption         =   "Brag Num"
      End
   End
   Begin Menu menu4 
      Caption         =   "&Other"
      Begin Menu Other9 
         Caption         =   "Kill Wait"
      End
      Begin Menu other4 
         Caption         =   "Sign Off"
      End
      Begin Menu other3 
         Caption         =   "-"
      End
      Begin Menu other5 
         Caption         =   "Kill Modal"
      End
   End
End
