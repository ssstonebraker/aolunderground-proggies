VERSION 2.00
Begin Form MMstatus 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "MM Status"
   ClientHeight    =   2688
   ClientLeft      =   1800
   ClientTop       =   1812
   ClientWidth     =   4428
   ControlBox      =   0   'False
   Height          =   3024
   Left            =   1752
   LinkTopic       =   "Form1"
   ScaleHeight     =   2688
   ScaleWidth      =   4428
   Top             =   1524
   Width           =   4524
   Begin SSFrame Frame3D2 
      Alignment       =   2  'Center
      Caption         =   "Mail Information"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Width           =   2415
      Begin Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   9.6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   720
         Width           =   615
      End
      Begin Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   9.6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   9.6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Mails:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   9.6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mails Left:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   9.6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mails Sent:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   9.6
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin SSFrame Frame3D1 
      Alignment       =   2  'Center
      Caption         =   "Dead Boxes"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1575
      Begin ListBox List1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   1290
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin SSCommand Command3D1 
      Caption         =   "Emergency Stop!"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.6
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4215
   End
   Begin Label Label9 
      BackColor       =   &H00800000&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   15
   End
   Begin Label Label8 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   4215
   End
   Begin Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   4215
   End
End
' MMSTATUS.FRM
Option Explicit

Sub Command3D1_Click ()
gv1210 = 1
MMMain.Show
Unload MMstatus
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
End Sub
