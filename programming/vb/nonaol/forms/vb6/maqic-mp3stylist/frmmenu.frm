VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   900
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   900
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu about 
         Caption         =   "&about"
      End
      Begin VB.Menu lmp3 
         Caption         =   "&load mp3"
      End
      Begin VB.Menu r3 
         Caption         =   "&remove mp3"
      End
      Begin VB.Menu fplay 
         Caption         =   "&file playing"
      End
      Begin VB.Menu stw 
         Caption         =   "&search the web"
      End
      Begin VB.Menu min 
         Caption         =   "&minimize"
      End
      Begin VB.Menu linegr 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&exit"
      End
   End
   Begin VB.Menu webrowser 
      Caption         =   "webrowser"
      Begin VB.Menu gtlm 
         Caption         =   "&goto lycos music"
      End
      Begin VB.Menu gmc 
         Caption         =   "&goto mp3.com"
      End
      Begin VB.Menu linecs 
         Caption         =   "-"
      End
      Begin VB.Menu wmm 
         Caption         =   "&want more mp3s?"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
dos32.FormNotOnTop Form1
MsgBox "mp3 stylist by maqic" + Chr$(10) + "mp3 player for windows 98 and up" + Chr$(10) + "send emails to maqic@maqicnet.tk" + Chr(10) + "website: http://www.maqicnet.tk", vbInformation, "mp3 stylist : about"
dos32.FormOnTop Form1
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub fplay_Click()
dos32.FormNotOnTop Form1
MsgBox "File now playing is :" & Form1.List1.Text, vbInformation, "mp3 stylist : file plaing"
dos32.FormOnTop Form1
End Sub

Private Sub gmc_Click()
Form3.WebBrowser1.Navigate ("http://www.mp3.com")
End Sub

Private Sub gtlm_Click()
Form3.WebBrowser1.Navigate ("http://www.lycos.com/music")
End Sub

Private Sub lmp3_Click()
Form1.CommonDialog1.Filter = ("*.mp3")
Form1.CommonDialog1.ShowOpen
If Form1.CommonDialog1.FileName = "" Then
Exit Sub
End If
mp32play = Form1.CommonDialog1.FileName
Form1.List1.AddItem mp32play
Form1.mp3lab.Caption = "status : mp3(s) loaded"
End Sub

Private Sub min_Click()
Form1.WindowState = 1
End Sub

Private Sub r3_Click()
If Form1.List1.Text = "" Then
Exit Sub
End If
Form1.List1.RemoveItem (List1)
End Sub

Private Sub stw_Click()
Form3.Show
End Sub

Private Sub wmm_Click()
dos32.FormNotOnTop Form3
MsgBox "want more mp3s?" + Chr$(10) + "Join togermano.ath.cx port 8888 on winmx or an opennap program.", vbInformation, "mp3 stylist : more mp3s"
dos32.FormOnTop Form3
End Sub
