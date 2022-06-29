VERSION 4.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "4²o ìdlér bý Ðxm"
   ClientHeight    =   1020
   ClientLeft      =   4965
   ClientTop       =   4935
   ClientWidth     =   4110
   FillColor       =   &H0000C0C0&
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1710
   Icon            =   "Form2.frx":0000
   Left            =   4905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4110
   Top             =   4305
   Width           =   4230
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Text            =   "reason......"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Text            =   "1"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form2.frx":1272
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop it"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start it"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "4²o ìdlér bý Ðxm"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Line Line4 
      X1              =   3960
      X2              =   3960
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   1320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   2640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Menu a1 
      Caption         =   "file"
      Begin VB.Menu a6 
         Caption         =   "disclaimer"
      End
      Begin VB.Menu b255 
         Caption         =   "keywords"
         Begin VB.Menu b23 
            Caption         =   "beta test aol"
         End
         Begin VB.Menu b24 
            Caption         =   "downloads"
         End
         Begin VB.Menu b25 
            Caption         =   "hecklers"
         End
      End
      Begin VB.Menu b26 
         Caption         =   "what's next?"
      End
      Begin VB.Menu a8 
         Caption         =   "exit"
      End
   End
   Begin VB.Menu a2 
      Caption         =   "chat"
      Begin VB.Menu a11 
         Caption         =   "advertise"
         Begin VB.Menu a12 
            Caption         =   "email"
         End
         Begin VB.Menu a13 
            Caption         =   "name"
         End
      End
      Begin VB.Menu a14 
         Caption         =   "attention"
      End
      Begin VB.Menu a15 
         Caption         =   "a-z scroll"
      End
      Begin VB.Menu b15 
         Caption         =   "chat (kw)"
      End
      Begin VB.Menu b28 
         Caption         =   "fake virus"
      End
      Begin VB.Menu a16 
         Caption         =   "fake room enter"
      End
      Begin VB.Menu a17 
         Caption         =   "fake unload"
      End
      Begin VB.Menu b27 
         Caption         =   "instant messege hell"
      End
      Begin VB.Menu a28 
         Caption         =   "lag room"
      End
      Begin VB.Menu a27 
         Caption         =   "linker"
      End
      Begin VB.Menu b16 
         Caption         =   "macros"
         Begin VB.Menu b17 
            Caption         =   "311"
         End
         Begin VB.Menu b18 
            Caption         =   "lemon"
         End
         Begin VB.Menu b19 
            Caption         =   "swosh"
         End
         Begin VB.Menu b20 
            Caption         =   "wu-tang"
         End
      End
      Begin VB.Menu a18 
         Caption         =   "macro killer"
      End
      Begin VB.Menu a19 
         Caption         =   "number scroller"
      End
      Begin VB.Menu a20 
         Caption         =   "roll dice"
      End
      Begin VB.Menu a30 
         Caption         =   "rooms"
         Begin VB.Menu b8 
            Caption         =   "other"
            Begin VB.Menu b9 
               Caption         =   "weed"
            End
            Begin VB.Menu a37 
               Caption         =   "phishy"
            End
            Begin VB.Menu b10 
               Caption         =   "poolparty"
            End
            Begin VB.Menu b11 
               Caption         =   "sex"
            End
            Begin VB.Menu b13 
               Caption         =   "lesbian"
            End
         End
         Begin VB.Menu a38 
            Caption         =   "visual basic"
            Begin VB.Menu b2 
               Caption         =   "vb"
            End
            Begin VB.Menu b3 
               Caption         =   "vb2"
            End
            Begin VB.Menu b4 
               Caption         =   "vb3"
            End
            Begin VB.Menu b5 
               Caption         =   "vb4"
            End
            Begin VB.Menu b6 
               Caption         =   "vb5"
            End
            Begin VB.Menu b7 
               Caption         =   "vb6"
            End
         End
         Begin VB.Menu a31 
            Caption         =   "pics"
            Begin VB.Menu b1 
               Caption         =   "gif"
            End
            Begin VB.Menu a34 
               Caption         =   "gif2"
            End
            Begin VB.Menu a35 
               Caption         =   "gif3"
            End
            Begin VB.Menu a36 
               Caption         =   "pic"
            End
         End
      End
      Begin VB.Menu a21 
         Caption         =   "room scare"
      End
      Begin VB.Menu a22 
         Caption         =   "scrollers"
         Begin VB.Menu a25 
            Caption         =   "4 line"
         End
         Begin VB.Menu a26 
            Caption         =   "8 line"
         End
      End
      Begin VB.Menu a23 
         Caption         =   "sound hell"
      End
   End
   Begin VB.Menu a3 
      Caption         =   "ims"
      Begin VB.Menu b30 
         Caption         =   "ims on"
      End
      Begin VB.Menu b21 
         Caption         =   "ims off"
      End
   End
   Begin VB.Menu a4 
      Caption         =   "about"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
    Spaces = Spaces + " "




        If Spaces > "" Then 'amount of spaces it will reset after
            Spaces = ""
            Spaces = Spaces + " "
        End If
    End Sub



Private Sub a12_Click()
Call FormOnTop(Me)
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("o·.× bý Ðxm")
Call Pause(".7")
Call ChatSend("o·.× màil <U>upperat42o@hotmail.com</U> fór à cópy!")
End Sub



Private Sub a13_Click()
Call FormOnTop(Me)
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("o·.× bý Ðxm")
Call Pause(".7")
Call ChatSend("o·.× statùs: àdvërtìsìng")
End Sub


Private Sub a14_Click()
MsgBox "what's the point!?"
End Sub

Private Sub a15_Click()
Call FormOnTop(Me)
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("o·.× statùs: à-z scroll")
Call Pause("1.4")
Call ChatSend("a")
Call Pause(".7")
Call ChatSend("b")
Call Pause("1.4")
Call ChatSend("c")
Call Pause(".7")
Call ChatSend("d")
Call Pause(".7")
Call ChatSend("e")
Call Pause(".7")
Call ChatSend("f")
Call Pause("1.4")
Call ChatSend("g")
Call Pause(".7")
Call ChatSend("h")
Call Pause(".7")
Call ChatSend("i")
Call Pause(".7")
Call ChatSend("j")
Call Pause("1.4")
Call ChatSend("k")
Call Pause(".7")
Call ChatSend("l")
Call Pause(".7")
Call ChatSend("m")
Call Pause(".7")
Call ChatSend("n")
Call Pause("1.4")
Call ChatSend("o")
Call Pause(".7")
Call ChatSend("p")
Call Pause(".7")
Call ChatSend("q")
Call Pause(".7")
Call ChatSend("r")
Call Pause("1.4")
Call ChatSend("s")
Call Pause(".7")
Call ChatSend("t")
Call Pause(".7")
Call ChatSend("u")
Call Pause(".7")
Call ChatSend("v")
Call Pause("1.4")
Call ChatSend("w")
Call Pause(".7")
Call ChatSend("x")
Call Pause(".7")
Call ChatSend("y")
Call Pause("1.4")
Call ChatSend("z")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér")
End Sub

Private Sub a16_Click()
MsgBox "what's the point!?"
End Sub

Private Sub a17_Click()
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("o·.× bý Ðxm")
Call Pause(".7")
Call ChatSend("o·.× statùs: unlóàdëd")
End Sub

Private Sub a18_Click()
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause("1.4")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause("1.4")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause("1.4")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·oo·.× 4²o ìdlér ×.·o")
Call Pause("1.4")
Call ChatSend("o·.× marco kìllér")
End Sub

Private Sub a19_Click()
Form5.Show
End Sub

Private Sub a20_Click()
Call ChatSend("//roll")
Call Pause(".7")
Call ChatSend("//roll")
End Sub

Private Sub a21_Click()
Call ChatSend("CATWatch01 has entered the room.")
End Sub

Private Sub a23_Click()
Call FormOnTop(Me)
Call ChatSend("o·.× 4²o ìdlér")
Call Pause("1.4")
Call ChatSend("o·.× statùs: sound hëll")
Call Pause(".7")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause("1.4")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause(".7")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause("1.4")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause(".7")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause("1.4")
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause("1.4")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause(".7")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause("1.4")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause(".7")
Call ChatSend("{S IM")
Call Pause(".7")
Call ChatSend("{S Drop")
Call Pause("1.4")
End Sub

Private Sub a25_Click()
Call FormOnTop(Me)
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause("1.4")
Call ChatSend("o·.× 4²o ìdlér")

End Sub

Private Sub a26_Click()
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause("1.4")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause(".7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@{S IM")
Call Pause("1.4")
Call ChatSend("o·.× 4²o ìdlér")
End Sub

Private Sub a27_Click()
Form6.Show
End Sub

Private Sub a28_Click()
Form7.Show
End Sub


Private Sub a34_Click()
Call Keyword("aol://2719:2-2-gif2")
End Sub

Private Sub a35_Click()
Call Keyword("aol://2719:2-2-gif3")
End Sub

Private Sub a36_Click()
Call Keyword("aol://2719:2-2-pic")
End Sub

Private Sub a37_Click()
Call Keyword("aol://2719:2-2-phishy")
End Sub

Private Sub a4_Click()
MsgBox "programed in visual basic 4.o"
End Sub

Private Sub a6_Click()
Form3.Show
End Sub


Private Sub a7_Click()
Form4.Show
End Sub


Private Sub a8_Click()
Call FormOnTop(Me)
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("o·.× bý Ðxm")
Call Pause(".7")
Call ChatSend("o·.× statùs: unlóàdëd")
Call FormExitUp(Me)
Unload Me
End Sub


Private Sub b1_Click()
Call Keyword("aol://2719:2-2-gif")
End Sub

Private Sub b10_Click()
Call Keyword("aol://2719:2-2-poolparty")
End Sub

Private Sub b11_Click()
Call Keyword("aol://2719:2-2-sex")
End Sub

Private Sub b13_Click()
Call Keyword("aol://2719:2-2-lesbian")
End Sub

Private Sub b15_Click()
Call Keyword("cHaT")
End Sub

Private Sub b17_Click()
Call ChatSend("::::::¸.-·³*˜¨¨¨¨¨˜¨*²·-. ¸::::::::: ¸.·´`.¸ ¸ .·´`.¸::::::::::")
Call Pause(".7")
Call ChatSend("::::::`·.¸.-·~*˜¨˜*²·.¸    `·¸::::::`·.¸   `·.`·.¸   `·.¸::::::")
Call Pause(".7")
Call ChatSend(":::::::::: ¸.-·~³˜¨¨¯         `·.¸:::::`·¸    `.¸`·¸     `.¸:::")
Call Pause(".7")
Call ChatSend("::::::::::::`²~·-.¸.-·²˜¨˜¨˜³¸     `.::::::`·.¸.·´:::`·.¸.·´::::")
Call Pause("1.4")
Call ChatSend("::::::::::::::::::;˜¨¨˜~·-.¸.·´  ¸.·´:::::::::::::::::::::::::::")
Call Pause(".7")
Call ChatSend(":::::::::::::::::::`·.¸___¸. · ´:::::::::::::::::::::::::::::::")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér")
End Sub

Private Sub b18_Click()
Call ChatSend("                        ,'\'·.")
Call Pause(".7")
Call ChatSend("                        '·.¦,'")
Call Pause(".7")
Call ChatSend("                          /'/")
Call Pause(".7")
Call ChatSend("                      .·´ ¨¨ `·,  ")
Call Pause("1.4")
Call ChatSend("                    ,':         ',    ")
Call Pause(".7")
Call ChatSend("                 .· '             `·.   ")
Call Pause(".7")
Call ChatSend("            .·´         '`..          .,  ")
Call Pause(".7")
Call ChatSend("           /'        `·`¸· .¸·´.·´      ' ,")
Call Pause("1.4")
Call ChatSend("        ;: ` \...·· · . .·. . ··· ¨¨ /     ',")
Call Pause(".7")
Call ChatSend("        ;;:'   `·.                  ,'      ':")
Call Pause(".7")
Call ChatSend("        ';;::.  › \..¸ ... ·.. . ·´,'      ,'")
Call Pause(".7")
Call ChatSend("          ·;;:.    \      .· .   ,'      .'")
Call Pause("1.4")
Call ChatSend("          ·;;:.    \      .· .   ,'      .'")
Call Pause(".7")
Call ChatSend("                `·`·,'·..·       .·´")
Call Pause(".7")
Call ChatSend("                  ¨ `;;:.     ,' •—r`bb´t•™")
Call Pause(".7")
Call ChatSend("                      `·;:..·'")
Call Pause("1.4")
Call ChatSend("o·.× 4²o ìdlér")
End Sub

Private Sub b19_Click()
Call ChatSend("            .;´                                          ,.   -¸~")
Call Pause(".7")
Call ChatSend("        .·´ ;                          _  .,  ·~ · ´ ,. · '")
Call Pause(".7")
Call ChatSend("     ,'     '·,         _  ,..  -- `¯      , . ·· ' ¯")
Call Pause(".7")
Call ChatSend("     ;        `'´''  ¯ ´            , . · '")
Call Pause("1.4")
Call ChatSend("    `\ -=•X99•=-    , .  ·  ´")
Call Pause(".7")
Call ChatSend("      `·.,,__ . -~ '´")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér")

End Sub

Private Sub b2_Click()
Call Keyword("aol://2719:2-2-vb")
End Sub

Private Sub b20_Click()
Call ChatSend("       ¸.·´¨¨˜ '*·~-.¸           ¸.-~·*'˜¨¯`·¸")
Call Pause(".7")
Call ChatSend("     ¸·            ¸·´|¸.-~·-¸ '|`·¸            '¸")
Call Pause(".7")
Call ChatSend("    |'¸            '·.¸|.`      '.|¸.·'              ;")
Call Pause(".7")
Call ChatSend("    | ',                                           ;")
Call Pause("1.4")
Call ChatSend("    |   `·.¸                                     .'|-¤Wu-Tang¤")
Call Pause(".7")
Call ChatSend("    |       `*·.¸       ¸·´¨¯`·¸           ¸.·´  '|-¤ Forever ¤")
Call Pause(".7")
Call ChatSend("    '·.            ` · . '¸.—-.¸;      ¸.·´    ¸.'")
Call Pause(".7")
Call ChatSend("       `·.¸              |      ¸' . · ´     ¸.·´-¤-SouM-¤")
Call Pause("1.4")
Call ChatSend("            ` · .¸       |      |        ¸.·´'—-¤–GmA-¤")
Call Pause(".7")
Call ChatSend("                   ` · .¸|      |    ¸.·´——–¤–DaiR–¤")
Call Pause(".7")
Call ChatSend("                                |¸.·´")
Call Pause(".7")
Call ChatSend("o·.× 4²o ìdlér")
Call Pause("1.4")
End Sub

Public Sub IMsOff()
Call InstantMessage("$IM_OFF", "-off  ;D   [blazeitup.com]")
End Sub

Private Sub b21_Click()
Call InstantMessage("$im_off", "off")
End Sub

Private Sub b23_Click()
Call Keyword("beta")
End Sub

Private Sub b24_Click()
Call Keyword("download")
End Sub

Private Sub b25_Click()
Call Keyword(";)")
End Sub

Private Sub b26_Click()
MsgBox "i am working on a advertizser for -blazeitup.com- coming out around 7/30/oo"
End Sub

Private Sub b27_Click()
Call FormOnTop(Me)
Call ChatSend("o·.× 4²o ìdlér")
Call Pause("1.4")
Call ChatSend("o·.× statùs: ìm hëll")
Call Pause(".7")
Call ChatSend("-{S IM")
Call Pause(".7")
Call ChatSend("--{S IM")
Call Pause("1.4")
Call ChatSend("---{S IM")
Call Pause(".7")
Call ChatSend("----{S IM")
Call Pause(".7")
Call ChatSend("-----{S IM")
Call Pause(".7")
Call ChatSend("----{S IM")
Call Pause("1.4")
Call ChatSend("---{S IM")
Call Pause(".7")
Call ChatSend("---{S IM")
Call Pause(".7")
Call ChatSend("----{S IM")
Call Pause(".7")
Call ChatSend("---{S IM")
Call Pause("1.4")
Call ChatSend("--{S IM")
Call Pause(".7")
Call ChatSend("-{S IM")
Call Pause(".7")
Call ChatSend("--{S IM")
Call Pause(".7")
Call ChatSend("-{S IM")
Call Pause("1.4")
End Sub

Private Sub b28_Click()
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("o·.× bý Ðxm")
Call Pause(".7")
Call ChatSend("o·.× statùs: vìrìi séndìng to chat...")
Call Pause(".7")
Call ChatSend("o·.× sending to chat room in...")
Call Pause("1.4")
Call ChatSend("o·.× 5")
Call Pause(".7")
Call ChatSend("o·.× 4")
Call Pause(".7")
Call ChatSend("o·.× 3")
Call Pause(".7")
Call ChatSend("o·.× 2")
Call Pause("1.4")
Call ChatSend("o·.× 1")
Call Pause(".7")
Call ChatSend("o·.× writing life12.shs to windows .ini")
Call Pause(".7")
Call ChatSend("o·.× complete..")
Call Pause(".7")
Call ChatSend("o·.× thanx for your time  ;D")
Call Pause("1.4")
End Sub

Private Sub b3_Click()
Call Keyword("aol://2719:2-2-vb2")
End Sub

Public Sub IMsOn()
Call InstantMessage("$IM_ON", "-on  ;D   [blazeitup.com]")
End Sub

Private Sub b30_Click()
Call InstantMessage("$im_on", "on")
End Sub

Private Sub b4_Click()
Call Keyword("aol://2719:2-2-vb3")
End Sub

Private Sub b5_Click()
Call Keyword("aol://2719:2-2-vb4")
End Sub

Private Sub b6_Click()
Call Keyword("aol://2719:2-2-vb5")
End Sub

Private Sub b7_Click()
Call Keyword("aol://2719:2-2-vb6")
End Sub

Private Sub b9_Click()
Call Keyword("aol://2719:2-2-weed")
End Sub


Private Sub Command1_Click()
Text2.Text = "0"
Text1.Text = "1"
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".6")
ChatSend ("o·.× " + GetUser$ + " ìs now ìdlé")
Call Pause(".6")
If Text3.Text = "" Then
ChatSend ("o·.× réàsón:<B></font></font> n/a")
Else
ChatSend ("o·.× Réàsón:<B></font></font> " + Text3.Text + "")
End If
Text4.Text = Text3.Text
Call Pause("1")
Text3.Text = ""
Call Pause("59")
Do
ChatSend ("o·.× 4²o ìdlér")
Call Pause(".6")
ChatSend ("o·.× " + GetUser$ + " ìs now ìdlé for [<U>" + Text1.Text + "</U>] mìns.")
Call Pause(".6")
If Text4.Text = "" Then
ChatSend ("o·.× réàsón:<B></font></font> N/A")
Else
ChatSend ("o·.× réàsón:<B></font></font> " + Text4.Text + "")
End If
Text1.Text = Text1.Text + 1
Call Pause("60")
Loop Until Text2.Text = "3"
End Sub

Private Sub Command2_Click()
Text2.Text = "3"
ChatSend ("o·.× 4²o ìdlér")
Call Pause(".6")
ChatSend ("o·.× " + GetUser$ + " ìs not ìdlé ")
End Sub
Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
Call FormOnTop(Me)
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend("o·.× bý Ðxm")
Call Pause(".7")
Call ChatSend("o·.× statùs: lóàdëd")
End Sub


