VERSION 4.00
Begin VB.Form Form7 
   Caption         =   "NoTePaD"
   ClientHeight    =   3705
   ClientLeft      =   1875
   ClientTop       =   2520
   ClientWidth     =   3315
   Height          =   4395
   Left            =   1815
   LinkTopic       =   "Form7"
   ScaleHeight     =   3705
   ScaleWidth      =   3315
   Top             =   1890
   Width           =   3435
   Begin RichtextLib.RichTextBox RichTextBox1 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   9585
      _Version        =   65536
      _ExtentX        =   16907
      _ExtentY        =   11033
      _StockProps     =   69
      BackColor       =   -2147483643
      ScrollBars      =   3
      TextRTF         =   $"PROG7.frx":0000
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   2865
      _Version        =   65536
      _ExtentX        =   5054
      _ExtentY        =   556
      _StockProps     =   192
      Appearance      =   1
   End
   Begin VB.Menu mnuFILE 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu2Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Mnu2Encrypt 
         Caption         =   "En&crypt"
      End
      Begin VB.Menu mnu2Unen 
         Caption         =   "&Unencrypt"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub asdf_Click()

End Sub

Private Sub adsf_Click()

End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Call StayOnTop(Form7)
mypath = CurDir
ProgressBar1.Visible = False
End Sub


Private Sub Gauge1_Change()

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Mnu2Encrypt_Click()
Counter = 0
Encry = RichTextBox1.text
Let Lenth% = Len(Encry)
If Lenth% > 0 Then
ProgressBar1.Max = Lenth%
Else
Exit Sub
End If
Screen.MousePointer = 11
ProgressBar1.Visible = True
ProgressBar1.Min = 0
ProgressBar1.Value = Min
Do: DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(Encry, numspc%, 1)
If nextchr$ = "A" Then
Let nextchr$ = "I"
ElseIf nextchr$ = "a" Then: Let nextchr$ = "i"
ElseIf nextchr$ = "B" Then: Let nextchr$ = "z"
ElseIf nextchr$ = "b" Then: Let nextchr$ = "Z"
ElseIf nextchr$ = "C" Then: Let nextchr$ = "L"
ElseIf nextchr$ = "c" Then: Let nextchr$ = "l"
ElseIf nextchr$ = "D" Then: Let nextchr$ = "r"
ElseIf nextchr$ = "d" Then: Let nextchr$ = "R"
ElseIf nextchr$ = "E" Then: Let nextchr$ = "n"
ElseIf nextchr$ = "e" Then: Let nextchr$ = "N"
ElseIf nextchr$ = "F" Then: Let nextchr$ = "M"
ElseIf nextchr$ = "f" Then: Let nextchr$ = "m"
ElseIf nextchr$ = "G" Then: Let nextchr$ = "y"
ElseIf nextchr$ = "g" Then: Let nextchr$ = "Y"
ElseIf nextchr$ = "H" Then: Let nextchr$ = "x"
ElseIf nextchr$ = "h" Then: Let nextchr$ = "X"
ElseIf nextchr$ = "I" Then: Let nextchr$ = "o"
ElseIf nextchr$ = "i" Then: Let nextchr$ = "O"
ElseIf nextchr$ = "J" Then: Let nextchr$ = "W"
ElseIf nextchr$ = "j" Then: Let nextchr$ = "w"
ElseIf nextchr$ = "K" Then: Let nextchr$ = "e"
ElseIf nextchr$ = "k" Then: Let nextchr$ = "E"
ElseIf nextchr$ = "L" Then: Let nextchr$ = "a"
ElseIf nextchr$ = "l" Then: Let nextchr$ = "A"
ElseIf nextchr$ = "M" Then: Let nextchr$ = "J"
ElseIf nextchr$ = "m" Then: Let nextchr$ = "j"
ElseIf nextchr$ = "N" Then: Let nextchr$ = "K"
ElseIf nextchr$ = "n" Then: Let nextchr$ = "k"
ElseIf nextchr$ = "O" Then: Let nextchr$ = "t"
ElseIf nextchr$ = "o" Then: Let nextchr$ = "T"
ElseIf nextchr$ = "P" Then: Let nextchr$ = "d"
ElseIf nextchr$ = "p" Then: Let nextchr$ = "D"
ElseIf nextchr$ = "Q" Then: Let nextchr$ = "b"
ElseIf nextchr$ = "q" Then: Let nextchr$ = "B"
ElseIf nextchr$ = "R" Then: Let nextchr$ = "Q"
ElseIf nextchr$ = "r" Then: Let nextchr$ = "q"
ElseIf nextchr$ = "S" Then: Let nextchr$ = "f"
ElseIf nextchr$ = "s" Then: Let nextchr$ = "F"
ElseIf nextchr$ = "T" Then: Let nextchr$ = "c"
ElseIf nextchr$ = "t" Then: Let nextchr$ = "C"
ElseIf nextchr$ = "U" Then: Let nextchr$ = "g"
ElseIf nextchr$ = "u" Then: Let nextchr$ = "G"
ElseIf nextchr$ = "V" Then: Let nextchr$ = "H"
ElseIf nextchr$ = "v" Then: Let nextchr$ = "h"
ElseIf nextchr$ = "W" Then: Let nextchr$ = "P"
ElseIf nextchr$ = "w" Then: Let nextchr$ = "p"
ElseIf nextchr$ = "X" Then: Let nextchr$ = "S"
ElseIf nextchr$ = "x" Then: Let nextchr$ = "s"
ElseIf nextchr$ = "Y" Then: Let nextchr$ = "u"
ElseIf nextchr$ = "y" Then: Let nextchr$ = "U"
ElseIf nextchr$ = "Z" Then: Let nextchr$ = "V"
ElseIf nextchr$ = "z" Then: Let nextchr$ = "v"
End If
Counter = Counter + 1
nextchr$ = nextchr$
Let newsent$ = newsent$ + nextchr$
ProgressBar1.Value = Counter
Loop Until numspc% = Lenth%
RichTextBox1.text = newsent$
ProgressBar1.Value = 0
ProgressBar1.Visible = False
Screen.MousePointer = 1
End Sub

Private Sub mnu2Unen_Click()
Screen.MousePointer = 11
UEncry = RichTextBox1.text
Let Lenth% = Len(UEncry)
ProgressBar1.Visible = True
ProgressBar1.Min = 0
ProgressBar1.Max = Lenth%
ProgressBar1.Value = Min
Counter = 0
Do: DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(UEncry, numspc%, 1)
If nextchr$ = "I" Then
Let nextchr$ = "A"
ElseIf nextchr$ = "i" Then: Let nextchr$ = "a"
ElseIf nextchr$ = "z" Then: Let nextchr$ = "B"
ElseIf nextchr$ = "Z" Then: Let nextchr$ = "b"
ElseIf nextchr$ = "r" Then: Let nextchr$ = "D"
ElseIf nextchr$ = "R" Then: Let nextchr$ = "d"
ElseIf nextchr$ = "n" Then: Let nextchr$ = "E"
ElseIf nextchr$ = "N" Then: Let nextchr$ = "e"
ElseIf nextchr$ = "M" Then: Let nextchr$ = "F"
ElseIf nextchr$ = "m" Then: Let nextchr$ = "f"
ElseIf nextchr$ = "y" Then: Let nextchr$ = "G"
ElseIf nextchr$ = "Y" Then: Let nextchr$ = "g"
ElseIf nextchr$ = "x" Then: Let nextchr$ = "H"
ElseIf nextchr$ = "X" Then: Let nextchr$ = "h"
ElseIf nextchr$ = "o" Then: Let nextchr$ = "I"
ElseIf nextchr$ = "O" Then: Let nextchr$ = "i"
ElseIf nextchr$ = "W" Then: Let nextchr$ = "J"
ElseIf nextchr$ = "w" Then: Let nextchr$ = "j"
ElseIf nextchr$ = "e" Then: Let nextchr$ = "K"
ElseIf nextchr$ = "E" Then: Let nextchr$ = "k"
ElseIf nextchr$ = "a" Then: Let nextchr$ = "L"
ElseIf nextchr$ = "A" Then: Let nextchr$ = "l"
ElseIf nextchr$ = "J" Then: Let nextchr$ = "M"
ElseIf nextchr$ = "j" Then: Let nextchr$ = "m"
ElseIf nextchr$ = "K" Then: Let nextchr$ = "N"
ElseIf nextchr$ = "k" Then: Let nextchr$ = "n"
ElseIf nextchr$ = "t" Then: Let nextchr$ = "O"
ElseIf nextchr$ = "T" Then: Let nextchr$ = "o"
ElseIf nextchr$ = "d" Then: Let nextchr$ = "P"
ElseIf nextchr$ = "D" Then: Let nextchr$ = "p"
ElseIf nextchr$ = "b" Then: Let nextchr$ = "Q"
ElseIf nextchr$ = "B" Then: Let nextchr$ = "q"
ElseIf nextchr$ = "Q" Then: Let nextchr$ = "R"
ElseIf nextchr$ = "q" Then: Let nextchr$ = "r"
ElseIf nextchr$ = "f" Then: Let nextchr$ = "S"
ElseIf nextchr$ = "F" Then: Let nextchr$ = "s"
ElseIf nextchr$ = "c" Then: Let nextchr$ = "T"
ElseIf nextchr$ = "C" Then: Let nextchr$ = "t"
ElseIf nextchr$ = "L" Then: Let nextchr$ = "C"
ElseIf nextchr$ = "l" Then: Let nextchr$ = "c"
ElseIf nextchr$ = "g" Then: Let nextchr$ = "U"
ElseIf nextchr$ = "G" Then: Let nextchr$ = "u"
ElseIf nextchr$ = "H" Then: Let nextchr$ = "V"
ElseIf nextchr$ = "h" Then: Let nextchr$ = "v"
ElseIf nextchr$ = "P" Then: Let nextchr$ = "W"
ElseIf nextchr$ = "p" Then: Let nextchr$ = "w"
ElseIf nextchr$ = "S" Then: Let nextchr$ = "X"
ElseIf nextchr$ = "s" Then: Let nextchr$ = "x"
ElseIf nextchr$ = "u" Then: Let nextchr$ = "Y"
ElseIf nextchr$ = "U" Then: Let nextchr$ = "y"
ElseIf nextchr$ = "V" Then: Let nextchr$ = "Z"
ElseIf nextchr$ = "v" Then: Let nextchr$ = "z"
End If
Counter = Counter + 1
Let newsent$ = newsent$ + nextchr$
ProgressBar1.Value = Counter
Loop Until numspc% = Lenth%
RichTextBox1.text = newsent$
ProgressBar1.Value = 0
ProgressBar1.Visible = False
Screen.MousePointer = 1
End Sub


Private Sub mnuExit_Click()
mypath = CurDir
Open mypath & "\Text.opp" For Input As #1
Input #1, thing$
Close #1
If thing$ <> RichTextBox1.text Then
Msg = "Do you want to Save current text?"   ' Define message.
Style = vbYesNoCancel
Title = "Save?"
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then
Form13.Show
Form7.Hide
End If
If Response = vbNo Then
RichTextBox1.text = ""
Form7.Hide
Form4.Show
End If
If Response = vbCancel Then
End If
Else
Form4.Show
Unload Form7
End If
End Sub


Private Sub mnuLoad_Click()
Form7.Hide
Form14.Show
End Sub

Private Sub mnuSave_Click()
Form13.Show
Form7.Hide
End Sub


Private Sub RichTextBox1_Change()
Let nextchr$ = " "
End Sub

