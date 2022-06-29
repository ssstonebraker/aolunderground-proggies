VERSION 5.00
Object = "{3B00B10A-6EF0-11D1-A6AA-0020AFE4DE54}#1.0#0"; "MP3PLAY.OCX"
Begin VB.Form Form1 
   Caption         =   "Mp3 Player Example by YaNg GoOn"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin MPEGPLAYLib.Mp3Play Mp3Play1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "<==  you should hide this!"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Dim err As Long
    'to load a .mp3 file do:
 err = Mp3Play1.Open("C:\whatever.mp3", "") 'opens the file c:\whatever.mp3
 'C:\whatever.mp3 is the file name of the .mp3 that you want to play
 
    'to load a .wav file do:
 err = Mp3Play1.Open("", "C:\whatever.wav") 'opens the file c:\whatever.wav
 'C:\whatever.wav is the file name of the .wav that you want to play

'****you have to delete one of the black lines for this to work!****'

 Mp3Play1.Play 'plays the file

End Sub

Private Sub Command2_Click()

Mp3Play1.Stop 'stops the .mp3 or .wav file

End Sub

Private Sub Mp3Play1_ActFrame(ByVal ActFrame As Long)

End Sub
