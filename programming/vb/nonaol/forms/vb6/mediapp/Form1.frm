VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load and Play a file"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim MultiMedia As New mMedia

    With CommonDialog1
    
    .Filter = "WaveAudio (*.wav)|*.wav|Midi (*.mid)|*.mid|Video Files(*.avi)|*.avi"
    .FilterIndex = 0
    .ShowOpen
    
    End With
If CommonDialog1.Filename <> "" Then
    MultiMedia.mmOpen CommonDialog1.Filename
    MultiMedia.mmPlay
End If

End Sub
