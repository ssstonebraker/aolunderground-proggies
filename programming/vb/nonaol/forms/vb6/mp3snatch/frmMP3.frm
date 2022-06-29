VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMP3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3Snatch v2.0"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgMP3 
      Left            =   3120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open MP3"
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   285
      Left            =   3300
      TabIndex        =   2
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txtMP3 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\SCD.MP3"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "View MP3 Info"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MP3 As clsMP3
Private Sub cmdInfo_Click()

    MP3.Filename = RTrim(txtMP3.Text)
    MsgBox "Track: " & MP3.Title & Chr(10) & _
        "Artist: " & MP3.Artist & Chr(10) & _
        "Album: " & MP3.Album & Chr(10) & _
        "Year: " & MP3.Year & Chr(10) & _
        "Comment: " & MP3.Comment & Chr(10) & _
        "Genre: " & MP3.Genre

End Sub


Private Sub cmdOpen_Click()

    ' *Very* basic dialogue handling ;-)

    dlgMP3.Filter = "MP3s|*.MP3"
    dlgMP3.ShowOpen
    
    If Right(UCase(dlgMP3.Filename), 4) = ".MP3" Then
        txtMP3.Text = dlgMP3.Filename
    End If

End Sub

Private Sub Form_Load()

    Set MP3 = New clsMP3

End Sub


