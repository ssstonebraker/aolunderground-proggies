VERSION 5.00
Object = "{CC01A07E-1AD2-11D4-8730-40C84FC10000}#2.0#0"; "nkPercent.ocx"
Begin VB.Form Form1 
   Caption         =   "nkPercent.ocx help forms - nk"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin Project1.nkPercent nkPercent1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start percent"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Dim PercentCount As Long
'name this ^ whatever you please

For PercentCount = 1 To 100 'starting number 1, ending percent 100%. you can change this too

    Call nkPercent1.DrawPercent(Picture1, PercentCount) 'draw percent in picture box
    
    Call nkPercent1.Pause(0.001) 'pause sub to slow down the percent bar

Next PercentCount

End Sub

