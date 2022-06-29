VERSION 2.00
Begin Form ELiTE 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "-< Ë£ï†ë †@£{]}<ë•R• >-"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   1425
   ClientWidth     =   5670
   Height          =   1515
   HelpContextID   =   70
   Left            =   0
   LinkTopic       =   "Form3"
   ScaleHeight     =   1110
   ScaleWidth      =   5670
   Top             =   1080
   Width           =   5790
   Begin SSPanel Panel3D3 
      BackColor       =   &H00000000&
      BevelInner      =   1  'Inset
      BevelWidth      =   3
      Caption         =   "Panel3D3"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   5655
      Begin SSCommand Command3D3 
         BevelWidth      =   1
         Caption         =   "Clear"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   0
         Top             =   120
         Width           =   1815
      End
      Begin SSCommand Command3D2 
         BevelWidth      =   1
         Caption         =   "Close"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin SSCommand Command3D1 
         BevelWidth      =   1
         Caption         =   "Send"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
   End
   Begin TextBox Text2 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2280
      Width           =   1335
   End
   Begin SSPanel Panel3D2 
      BackColor       =   &H00000000&
      BevelInner      =   1  'Inset
      BevelWidth      =   3
      Caption         =   "Panel3D2"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2775
      Begin TextBox Text3 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2535
      End
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00000000&
      BevelInner      =   1  'Inset
      BevelWidth      =   3
      Caption         =   "Panel3D1"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      Begin TextBox Text1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "Type Here"
         Top             =   120
         Width           =   2655
      End
   End
End

Sub Command3D1_Click ()
aol% = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol%, "MDIClient")
welc% = FindChildByTitle(mdi%, "Welcome, ")
If welc% = 0 Then
MsgBox "You must sign on first.", 16, "Not signed on"
Exit Sub
Else
aol% = FindWindow("AOL Frame25", 0&)
x = SetFocusAPI(aol%)
x10% = FindCht()
edt% = FindChildByClass(x10%, "_AOL_Edit")
DoEvents
Call SendText(edt%, Space(0) & Text3.Text)
entr = SendMessageByNum(edt%, WM_CHAR, 13, 0)
End If
End Sub

Sub Command3D2_Click ()
Unload Elite
End Sub

Sub Command3D3_Click ()
Text3.Text = " "
End Sub

Sub Form_Resize ()
MsgBox "Using this in no way will make you elite. It may even cause you to be made fun of by others. Use at your own risk, hehehe.", 16, "Disclaimer"
Dim success%
success% = SetWindowPos(Elite.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
Command3D1.Font3D = 3
Command3D2.Font3D = 3
Command3D3.Font3D = 3
End Sub

Sub Input1_Change ()
End Sub

Sub Text1_Change ()
If Text1.Text = "a" Then
    text2.Text = "@"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "b" Then
    text2.Text = "þ"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "c" Then
    text2.Text = "©"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "d" Then
    text2.Text = "d"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "e" Then
    text2.Text = "ë"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "f" Then
    text2.Text = "ƒ"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "g" Then
    text2.Text = "g"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "h" Then
    text2.Text = "h"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "i" Then
    text2.Text = "ï"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "j" Then
    text2.Text = "j"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "k" Then
    text2.Text = "k"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "l" Then
    text2.Text = "l"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "m" Then
    text2.Text = "m"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "n" Then
    text2.Text = "ñ"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "o" Then
    text2.Text = "ø"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "p" Then
    text2.Text = "p"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "q" Then
    text2.Text = "q"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "r" Then
    text2.Text = "®"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "s" Then
    text2.Text = "$"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "t" Then
    text2.Text = "†"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "u" Then
    text2.Text = "ü"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "v" Then
    text2.Text = "v"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "w" Then
    text2.Text = "vv"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "x" Then
    text2.Text = "×"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "y" Then
    text2.Text = "ý"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "z" Then
    text2.Text = "z"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = " " Then
    text2.Text = " "
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "!" Then
    text2.Text = "¡"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "1" Then
    text2.Text = "1"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "2" Then
    text2.Text = "2"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "3" Then
    text2.Text = "3"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "4" Then
    text2.Text = "4"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "5" Then
    text2.Text = "5"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "6" Then
    text2.Text = "6"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "7" Then
    text2.Text = "7"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "8" Then
    text2.Text = "8"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "9" Then
    text2.Text = "9"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "0" Then
    text2.Text = "0"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "A" Then
    text2.Text = "Å"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "B" Then
    text2.Text = "ß"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "C" Then
    text2.Text = "Ç"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "D" Then
    text2.Text = "Ð"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "E" Then
    text2.Text = "Ë"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "F" Then
    text2.Text = "ƒ"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "G" Then
    text2.Text = "G"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "H" Then
    text2.Text = "{)}-{(}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "I" Then
    text2.Text = "‡"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "J" Then
    text2.Text = "J"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "K" Then
    text2.Text = "{]}<"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "L" Then
    text2.Text = "£"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "M" Then
    text2.Text = "{(}V{)}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "N" Then
    text2.Text = "{]}\{[}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "O" Then
    text2.Text = "Õ"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "P" Then
    text2.Text = "¶"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "Q" Then
    text2.Text = "`Q"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "R" Then
    text2.Text = "•R•"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "S" Then
    text2.Text = "§"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "T" Then
    text2.Text = "†"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "U" Then
    text2.Text = "Ü"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "V" Then
    text2.Text = "\/"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "W" Then
    text2.Text = "\X/"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "X" Then
    text2.Text = "><"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "Y" Then
    text2.Text = "¥"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "Z" Then
    text2.Text = "Z"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "~" Then
    text2.Text = "~"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "`" Then
    text2.Text = "`"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "!" Then
    text2.Text = "¡"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "@" Then
    text2.Text = "ä"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "#" Then
    text2.Text = "#"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "$" Then
    text2.Text = "$"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "%" Then
    text2.Text = "%"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "^" Then
    text2.Text = "{^}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "&" Then
    text2.Text = "&"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "*" Then
    text2.Text = "™"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "(" Then
    text2.Text = "{(}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = ")" Then
    text2.Text = "{)}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "-" Then
    text2.Text = "-"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "_" Then
    text2.Text = "_"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "+" Then
    text2.Text = "+"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "=" Then
    text2.Text = "="
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "[" Then
    text2.Text = "{[}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "]" Then
    text2.Text = "{]}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "{" Then
    text2.Text = "{{}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "}" Then
    text2.Text = "{}}"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = ":" Then
    text2.Text = ":"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = ";" Then
    text2.Text = ";"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "'" Then
    text2.Text = "'"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "," Then
    text2.Text = ","
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "." Then
    text2.Text = "."
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "<" Then
    text2.Text = "<"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = ">" Then
    text2.Text = ">"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
If Text1.Text = "?" Then
    text2.Text = "¿"
    Text1.Text = ""
    Text3.Text = Text3.Text + text2.Text
    Text1.SetFocus
End If
    If Text1.Text = "_" Then
        Text1.Text = ""
        Text3.SetFocus
            SendKeys "{RIGHT 90}"
            SendKeys "{Backspace}"
        DoEvents
        Text1.SetFocus
   End If

End Sub

Sub Text1_GotFocus ()
    Text1.SelStart = 0
    Text1.SelLength = 65000

End Sub

Sub Text1_KeyPress (Keyascii As Integer)
If (Keyascii = 8) And Len(Text3.Text) > 0 Then
    Text3.Text = Left$(Text3.Text, Len(Text3.Text) - 1)
End If
If (Keyascii = 13) Then
    SendKeys Text3
End If

End Sub

