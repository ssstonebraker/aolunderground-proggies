VERSION 2.00
Begin Form welcome 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Welcome to FrostBite Mass Mailer v2.o By DeCK"
   ClientHeight    =   4368
   ClientLeft      =   1356
   ClientTop       =   1500
   ClientWidth     =   4872
   ControlBox      =   0   'False
   Height          =   4704
   Left            =   1308
   LinkTopic       =   "Form1"
   ScaleHeight     =   4368
   ScaleWidth      =   4872
   Top             =   1212
   Width           =   4968
   Begin CommandButton Command2 
      Caption         =   "View Greetz"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin CommandButton Command1 
      Caption         =   "To The MMer..."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   120
      ScaleHeight     =   828
      ScaleWidth      =   4668
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin Line Line1 
         BorderColor     =   &H0000FF00&
         Visible         =   0   'False
         X1              =   0
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "FrostBite Mass Mailer v2.o"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   13.2
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
   Begin Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "           Welcome to FrostBite Mass Mailer v2.o By DeCK.  Now you are probably asking yourself what is so special about this MMer.  Well I'll tell ya.  This MMer is the fastest on the scene, you dont believe me? Test it and tell your buddies about it.  I didn't make this MMer for any peticular group, I made so everyone can enjoy it so please spread it around, if you like please email me at deck97@usa.net and give me any comments, questions, or suggestions for my next version.  Please dont be shy, use it and give me your feedback.  Thank you and I hope you enjoy this.                                                                                                                                                                                                                               -DeCK"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4575
      WordWrap        =   -1  'True
   End
End
' WELCOME.FRM
Option Explicit

Sub Command1_Click ()
MMMain.Show
Unload Me
End Sub

Sub Command2_Click ()
greetz.Show
Me.Hide
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
Label1.Left = -5000
Label2.Left = 7000
Do: DoEvents
Label1.Left = Label1.Left + 25
Loop Until Label1.Left >= 360
Line1.Visible = True
Do: DoEvents
Label2.Left = Label2.Left - 175
Loop Until Label2.Left <= 240
Command1.Visible = True
Command2.Visible = True
End Sub
