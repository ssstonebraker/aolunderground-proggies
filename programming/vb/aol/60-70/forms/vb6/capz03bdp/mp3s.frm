VERSION 5.00
Begin VB.Form mp3s 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   1455
      Left            =   240
      Pattern         =   "*.mp3"
      ReadOnly        =   0   'False
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "  X"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "mp3s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_DblClick()
ChatSend ".play " & (File1.FileName)
End Sub

Private Sub Form_Load()
File1.Path = GetFromINI("Settings", "dir", "c:\windows\system\cap.set")
FormTop Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
