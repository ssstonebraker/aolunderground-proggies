VERSION 2.00
Begin Form greetz 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Greetz"
   ClientHeight    =   4020
   ClientLeft      =   1092
   ClientTop       =   1512
   ClientWidth     =   7404
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   4356
   Left            =   1044
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7404
   Top             =   1224
   Width           =   7500
   Begin CommandButton Command1 
      Caption         =   "Start"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Now Presenting..."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00004080&
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   9495
   End
   Begin Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MaD"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   95
      TabIndex        =   1
      Top             =   2050
      Width           =   9500
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MaD"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   100
      TabIndex        =   0
      Top             =   2040
      Width           =   9500
   End
End
' GREETZ.FRM
Option Explicit

Sub Command1_Click ()
sub05BC Me
sub02A9 Me
Command1.Visible = False
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Me.WindowState = 2
Label1.Visible = False
Label2.Visible = False
Label3.FontSize = "1"
Do: DoEvents
Label3.FontSize = Label3.FontSize + 1
sub060A .01#
Loop Until Label3.FontSize >= 50
sub060A .7#
Label3.FontSize = "50"
Do: DoEvents
Label3.FontSize = Label3.FontSize - 1
sub060A .01#
Loop Until Label3.FontSize <= 1
Label3.Visible = False
Label1.Visible = True
Label2.Visible = True
Label1.FontSize = 100
Label1.Caption = "MaD"
Label2.FontSize = 100
Label2.Caption = "MaD"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "ProGGeR"
Label2.FontSize = 100
Label2.Caption = "ProGGeR"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "MisterX"
Label2.FontSize = 100
Label2.Caption = "MisterX"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "ZaR"
Label2.FontSize = 100
Label2.Caption = "ZaR"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "RoT"
Label2.FontSize = 100
Label2.Caption = "RoT"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "EViL"
Label2.FontSize = 100
Label2.Caption = "EViL"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "SaS"
Label2.FontSize = 100
Label2.Caption = "SaS"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "Lite"
Label2.FontSize = 100
Label2.Caption = "Lite"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "Ruler"
Label2.FontSize = 100
Label2.Caption = "Ruler"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "Dark"
Label2.FontSize = 100
Label2.Caption = "Dark"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "Genius"
Label2.FontSize = 100
Label2.Caption = "Genius"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "FiX"
Label2.FontSize = 100
Label2.Caption = "FiX"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "drizzt"
Label2.FontSize = 100
Label2.Caption = "drizzt"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "Code"
Label2.FontSize = 100
Label2.Caption = "Code"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "DRaFt"
Label2.FontSize = 100
Label2.Caption = "DRaFt"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "RaSH"
Label2.FontSize = 100
Label2.Caption = "RaSH"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "SaInT"
Label2.FontSize = 100
Label2.Caption = "SaInT"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
Label1.Top = 2040
Label1.FontSize = 100
Label1.Caption = "Others"
Label2.FontSize = 100
Label2.Caption = "Others"
Label1.Left = -5000
Do: DoEvents
Label1.Left = Label1.Left + 100
Loop Until Label1.Left >= 100
sub060A 2
Do: DoEvents
Label1.Top = Label1.Top - 50
Loop Until Label1.Top <= -5000
MMMain.Show
Unload Me
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
End Sub
