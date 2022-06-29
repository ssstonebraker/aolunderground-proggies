VERSION 2.00
Begin Form Form72 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "GreeTz To..."
   ClientHeight    =   600
   ClientLeft      =   3570
   ClientTop       =   2010
   ClientWidth     =   1920
   ControlBox      =   0   'False
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Arial"
   FontSize        =   8.25
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   1005
   Left            =   3510
   LinkTopic       =   "Form7"
   ScaleHeight     =   600
   ScaleWidth      =   1920
   Top             =   1665
   Width           =   2040
   Begin CommandButton Command1 
      Caption         =   "Okay!"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   2520
   End
   Begin Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   570
   End
   Begin Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
End
' ICEGRT.FRM
Option Explicit

Sub Command1_Click ()
Unload Me
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub Timer1_Timer ()
Label1.Caption = "iNTuiT"
Label2.Caption = "iNTuiT"
Label3.Caption = "iNTuiT"
Label4.Caption = "iNTuiT"
Label5.Caption = "iNTuiT"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "iCa"
Label2.Caption = "iCa"
Label3.Caption = "iCa"
Label4.Caption = "iCa"
Label5.Caption = "iCa"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "FerGus"
Label2.Caption = "FerGus"
Label3.Caption = "FerGus"
Label4.Caption = "FerGus"
Label5.Caption = "FerGus"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "ePoD"
Label2.Caption = "ePoD"
Label3.Caption = "ePoD"
Label4.Caption = "ePoD"
Label5.Caption = "ePoD"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "GuS"
Label2.Caption = "GuS"
Label3.Caption = "GuS"
Label4.Caption = "GuS"
Label5.Caption = "GuS"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "LaZy"
Label2.Caption = "LaZy"
Label3.Caption = "LaZy"
Label4.Caption = "LaZy"
Label5.Caption = "LaZy"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "CeD"
Label2.Caption = "CeD"
Label3.Caption = "CeD"
Label4.Caption = "CeD"
Label5.Caption = "CeD"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "WAVe"
Label2.Caption = "WAVe"
Label3.Caption = "WAVe"
Label4.Caption = "WAVe"
Label5.Caption = "WAVe"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "NeK"
Label2.Caption = "NeK"
Label3.Caption = "NeK"
Label4.Caption = "NeK"
Label5.Caption = "NeK"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "JaGo"
Label2.Caption = "JaGo"
Label3.Caption = "JaGo"
Label4.Caption = "JaGo"
Label5.Caption = "JaGo"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "TeeDaYz"
Label2.Caption = "TeeDaYz"
Label3.Caption = "TeeDaYz"
Label4.Caption = "TeeDaYz"
Label5.Caption = "TeeDaYz"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "KiRbY"
Label2.Caption = "KiRbY"
Label3.Caption = "KiRbY"
Label4.Caption = "KiRbY"
Label5.Caption = "KiRbY"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "BoNe"
Label2.Caption = "BoNe"
Label3.Caption = "BoNe"
Label4.Caption = "BoNe"
Label5.Caption = "BoNe"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "RPM"
Label2.Caption = "RPM"
Label3.Caption = "RPM"
Label4.Caption = "RPM"
Label5.Caption = "RPM"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "HuGz"
Label2.Caption = "HuGz"
Label3.Caption = "HuGz"
Label4.Caption = "HuGz"
Label5.Caption = "HuGz"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "FuSiON"
Label2.Caption = "FuSiON"
Label3.Caption = "FuSiON"
Label4.Caption = "FuSiON"
Label5.Caption = "FuSiON"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "GoRDo"
Label2.Caption = "GoRDo"
Label3.Caption = "GoRDo"
Label4.Caption = "GoRDo"
Label5.Caption = "GoRDo"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Buck"
Label2.Caption = "Buck"
Label3.Caption = "Buck"
Label4.Caption = "Buck"
Label5.Caption = "Buck"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Lee"
Label2.Caption = "Lee"
Label3.Caption = "Lee"
Label4.Caption = "Lee"
Label5.Caption = "Lee"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Malik"
Label2.Caption = "Malik"
Label3.Caption = "Malik"
Label4.Caption = "Malik"
Label5.Caption = "Malik"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "CiL"
Label2.Caption = "CiL"
Label3.Caption = "CiL"
Label4.Caption = "CiL"
Label5.Caption = "CiL"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Pher"
Label2.Caption = "Pher"
Label3.Caption = "Pher"
Label4.Caption = "Pher"
Label5.Caption = "Pher"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "TasK"
Label2.Caption = "TasK"
Label3.Caption = "TasK"
Label4.Caption = "TasK"
Label5.Caption = "TasK"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Pizza"
Label2.Caption = "Pizza"
Label3.Caption = "Pizza"
Label4.Caption = "Pizza"
Label5.Caption = "Pizza"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "KeRSeD"
Label2.Caption = "KeRSeD"
Label3.Caption = "KeRSeD"
Label4.Caption = "KeRSeD"
Label5.Caption = "KeRSeD"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Slayer"
Label2.Caption = "Slayer"
Label3.Caption = "Slayer"
Label4.Caption = "Slayer"
Label5.Caption = "Slayer"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Napalm"
Label2.Caption = "Napalm"
Label3.Caption = "Napalm"
Label4.Caption = "Napalm"
Label5.Caption = "Napalm"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "ProGG"
Label2.Caption = "ProGG"
Label3.Caption = "ProGG"
Label4.Caption = "ProGG"
Label5.Caption = "ProGG"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "SPuL"
Label2.Caption = "SPuL"
Label3.Caption = "SPuL"
Label4.Caption = "SPuL"
Label5.Caption = "SPuL"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "FuDD"
Label2.Caption = "FuDD"
Label3.Caption = "FuDD"
Label4.Caption = "FuDD"
Label5.Caption = "FuDD"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "AsTrO"
Label2.Caption = "AsTrO"
Label3.Caption = "AsTrO"
Label4.Caption = "AsTrO"
Label5.Caption = "AsTrO"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "GraVitY"
Label2.Caption = "GraVitY"
Label3.Caption = "GraVitY"
Label4.Caption = "GraVitY"
Label5.Caption = "GraVitY"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "CLouD"
Label2.Caption = "CLouD"
Label3.Caption = "CLouD"
Label4.Caption = "CLouD"
Label5.Caption = "CLouD"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "ReMMi"
Label2.Caption = "ReMMi"
Label3.Caption = "ReMMi"
Label4.Caption = "ReMMi"
Label5.Caption = "ReMMi"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Bombr"
Label2.Caption = "Bombr"
Label3.Caption = "Bombr"
Label4.Caption = "Bombr"
Label5.Caption = "Bombr"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "The Dorks"
Label2.Caption = "The Dorks"
Label3.Caption = "The Dorks"
Label4.Caption = "The Dorks"
Label5.Caption = "The Dorks"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "Everyone Else"
Label2.Caption = "Everyone Else"
Label3.Caption = "Everyone Else"
Label4.Caption = "Everyone Else"
Label5.Caption = "Everyone Else"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Label1.Caption = "YuBaN Enterprises"
Label2.Caption = "YuBaN Enterprises"
Label3.Caption = "YuBaN Enterprises"
Label4.Caption = "YuBaN Enterprises"
Label5.Caption = "YuBaN Enterprises"
Label1.Visible = True
subD08 .1#
Label1.Visible = False
subD08 .2#
Label2.Visible = True
subD08 .1#
Label2.Visible = False
subD08 .2#
Label3.Visible = True
subD08 .1#
Label3.Visible = False
subD08 .2#
Label4.Visible = True
subD08 .1#
Label4.Visible = False
subD08 .2#
Label5.Visible = True
subD08 .2#
Label5.Visible = False
Command1.Visible = True
Timer1.Enabled = False
Unload Form7
Exit Sub
End Sub
