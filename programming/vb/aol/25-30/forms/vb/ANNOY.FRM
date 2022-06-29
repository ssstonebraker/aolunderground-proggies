VERSION 2.00
Begin Form frmDialog 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Wake those users up!"
   ClientHeight    =   1785
   ClientLeft      =   255
   ClientTop       =   2715
   ClientWidth     =   5235
   ControlBox      =   0   'False
   Height          =   2190
   Left            =   195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5235
   Top             =   2370
   Width           =   5355
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3810
      TabIndex        =   1
      Top             =   1170
      Width           =   1215
   End
   Begin Timer timTimer 
      Interval        =   10000
      Left            =   360
      Top             =   1260
   End
   Begin Label lblLabel 
      Caption         =   "Time for a wake up call. Dialog will get upset if you don't hit OK within 10 seconds - wait and see."
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   690
      Left            =   375
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Option Explicit

Sub cmdOK_Click ()

    Unload Me


End Sub

Sub timTimer_Timer ()

    lblLabel.Caption = "At this point you should be getting really annoyed with the noise. Probably a good idea to hit OK now!"
    Beep
    timTimer.Interval = 500
    timTimer.Enabled = True

End Sub

