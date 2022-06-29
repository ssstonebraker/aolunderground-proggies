VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3480
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   6090
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Label 
      Caption         =   "Label"
      Begin VB.Menu About 
         Caption         =   "About..."
      End
      Begin VB.Menu Help 
         Caption         =   "Help "
      End
      Begin VB.Menu asdfg 
         Caption         =   "-"
      End
      Begin VB.Menu isits 
         Caption         =   "Is it us?"
      End
      Begin VB.Menu MailUs 
         Caption         =   "Mail Us"
      End
      Begin VB.Menu Web_Page 
         Caption         =   "Web Page"
      End
      Begin VB.Menu asdfgh 
         Caption         =   "-"
      End
      Begin VB.Menu Greetz 
         Caption         =   "Greetz"
      End
      Begin VB.Menu Credits 
         Caption         =   "Credits"
      End
      Begin VB.Menu spave 
         Caption         =   "-"
      End
      Begin VB.Menu Minimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu Mnuexitme 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu label1 
      Caption         =   "Label1"
      Begin VB.Menu AddRoom 
         Caption         =   "Add Room"
      End
      Begin VB.Menu Mail_Room 
         Caption         =   "Mail Room"
      End
      Begin VB.Menu Aol_uhh 
         Caption         =   "-"
      End
      Begin VB.Menu Chat_Clear 
         Caption         =   "Chat Clear"
      End
      Begin VB.Menu I_Talker 
         Caption         =   "Invisible Talker"
      End
      Begin VB.Menu Room_Anoy 
         Caption         =   "Room Anoy"
      End
      Begin VB.Menu Link_Sender 
         Caption         =   "Link Sender"
      End
      Begin VB.Menu RoomWav 
         Caption         =   "Room Wav Player"
      End
      Begin VB.Menu T_T 
         Caption         =   "-"
      End
      Begin VB.Menu Advertise 
         Caption         =   "Advertise"
         Begin VB.Menu GetFull 
            Caption         =   "Get Full Moon"
         End
         Begin VB.Menu FullMacro 
            Caption         =   "Full Moon Macro"
         End
         Begin VB.Menu LinkToPage 
            Caption         =   "Link To Page"
         End
         Begin VB.Menu NormalAdvertise 
            Caption         =   "Normal Adverise"
         End
      End
      Begin VB.Menu Talkers 
         Caption         =   "Talkers"
         Begin VB.Menu EliterTalker 
            Caption         =   "Elite Talker"
         End
         Begin VB.Menu Bak 
            Caption         =   "BackWord Talker"
         End
         Begin VB.Menu HackerTalker 
            Caption         =   "Hacker Talker"
         End
         Begin VB.Menu EncryptDe 
            Caption         =   "Encrypt/Decrypt"
         End
         Begin VB.Menu MChat 
            Caption         =   "M-Chat"
         End
      End
      Begin VB.Menu Bot_s 
         Caption         =   "Bots"
         Begin VB.Menu Echo_Bot 
            Caption         =   "Echo Bot"
         End
         Begin VB.Menu Voter_Bot 
            Caption         =   "Voter Bot"
         End
         Begin VB.Menu WarezRequest 
            Caption         =   "Warez Request"
         End
         Begin VB.Menu Ball 
            Caption         =   "8Ball"
         End
         Begin VB.Menu afk 
            Caption         =   "Afk"
         End
         Begin VB.Menu Attention 
            Caption         =   "Attention"
         End
         Begin VB.Menu Livid 
            Caption         =   "Livid Ebonics"
         End
         Begin VB.Menu Sushi 
            Caption         =   "Sushi Ebonics"
         End
      End
      Begin VB.Menu Room_Buster 
         Caption         =   "Room Buster"
      End
      Begin VB.Menu o 
         Caption         =   "-"
      End
      Begin VB.Menu Mail_server 
         Caption         =   "Mail Server"
      End
      Begin VB.Menu Mass_Mailer 
         Caption         =   "Mass Mailer"
      End
      Begin VB.Menu MassForward 
         Caption         =   "Mass Forward"
      End
      Begin VB.Menu ListMaker 
         Caption         =   "List Maker"
      End
      Begin VB.Menu errrrrrrr 
         Caption         =   "-"
      End
      Begin VB.Menu MailPref 
         Caption         =   "Mail Prefrences"
      End
      Begin VB.Menu MailTools 
         Caption         =   "Mail Tools"
      End
      Begin VB.Menu CountMail 
         Caption         =   "Count Mail"
      End
      Begin VB.Menu Server_Hlp 
         Caption         =   "Server Helper"
      End
      Begin VB.Menu k 
         Caption         =   "-"
      End
      Begin VB.Menu Fader 
         Caption         =   "Fader"
      End
   End
   Begin VB.Menu Label2 
      Caption         =   "Label2"
      Begin VB.Menu Color_Coder 
         Caption         =   "Color Coder"
      End
      Begin VB.Menu Macro_Shop 
         Caption         =   "Macro Shop"
      End
      Begin VB.Menu AsciiShop 
         Caption         =   "Ascii Shop"
      End
      Begin VB.Menu asdf 
         Caption         =   "-"
      End
      Begin VB.Menu Pw_Gen 
         Caption         =   "Pw Generator"
      End
      Begin VB.Menu Reset_Yaa 
         Caption         =   "SN Reset"
      End
      Begin VB.Menu SnDecoder 
         Caption         =   "Sn Decoder"
      End
      Begin VB.Menu Phish_Manager 
         Caption         =   "Phish Manager"
      End
      Begin VB.Menu AOl_ok 
         Caption         =   "-"
      End
      Begin VB.Menu Hide_Aol 
         Caption         =   "Hide AOL"
      End
      Begin VB.Menu KillTopWindow 
         Caption         =   "Kill Top Window"
      End
      Begin VB.Menu Killwait 
         Caption         =   "Kill Wait "
      End
      Begin VB.Menu Up_Chat 
         Caption         =   "Up-Chat"
      End
      Begin VB.Menu killer 
         Caption         =   "45min Killer"
      End
      Begin VB.Menu ahhhhhhhhhhhhhhhh 
         Caption         =   "-"
      End
      Begin VB.Menu QuickKW 
         Caption         =   "Quick Keyword"
      End
      Begin VB.Menu Tetris_Here 
         Caption         =   "Tetris"
      End
      Begin VB.Menu PWSD 
         Caption         =   "PWSD"
      End
      Begin VB.Menu WavPlayer 
         Caption         =   "Wav Player"
      End
      Begin VB.Menu er 
         Caption         =   "-"
      End
      Begin VB.Menu Auto 
         Caption         =   "Auto SignOff"
      End
      Begin VB.Menu SignOff 
         Caption         =   "Sign Off"
      End
   End
   Begin VB.Menu Label3 
      Caption         =   "Label3"
      Begin VB.Menu Im_on 
         Caption         =   "IMs On"
      End
      Begin VB.Menu Im_Off 
         Caption         =   "IMs  Off"
      End
      Begin VB.Menu ImIgnore 
         Caption         =   "IMs Ignore"
      End
      Begin VB.Menu Im_Anwser 
         Caption         =   "IM Anwser"
      End
      Begin VB.Menu Duh_er 
         Caption         =   "-"
      End
      Begin VB.Menu AimMassIM 
         Caption         =   "Aim Mass IM"
      End
      Begin VB.Menu Aimphisher 
         Caption         =   "Aim Phisher"
      End
      Begin VB.Menu darn 
         Caption         =   "-"
      End
      Begin VB.Menu MassIM 
         Caption         =   "Mass IM"
      End
      Begin VB.Menu Phisher 
         Caption         =   "Phisher"
      End
   End
   Begin VB.Menu Label4 
      Caption         =   "Label4"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEndGame 
         Caption         =   "&End Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu MM 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "&High Scores"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MnuOptions 
         Caption         =   "&Options....."
      End
      Begin VB.Menu MnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Label5 
      Caption         =   "Label5"
      Begin VB.Menu MnuInstructions 
         Caption         =   "&Instructions"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu Label6 
      Caption         =   "&Label6 "
      Begin VB.Menu Scroll 
         Caption         =   "Scroll "
      End
      Begin VB.Menu Print 
         Caption         =   "Print"
      End
      Begin VB.Menu Colors 
         Caption         =   "Colors"
         Begin VB.Menu BackColor 
            Caption         =   "Back Color"
         End
         Begin VB.Menu TextColor 
            Caption         =   "Text Color"
         End
      End
      Begin VB.Menu Load 
         Caption         =   "Load"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu Clear 
         Caption         =   "Clear "
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit "
      End
   End
   Begin VB.Menu Label7 
      Caption         =   "&Label7"
      Begin VB.Menu Undo 
         Caption         =   "Undo"
      End
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu Fonts 
         Caption         =   "Fonts"
         Begin VB.Menu Font1 
            Caption         =   "Font1"
         End
         Begin VB.Menu Font2 
            Caption         =   "Font2"
         End
      End
      Begin VB.Menu Macros 
         Caption         =   "Macros"
      End
   End
   Begin VB.Menu Label8 
      Caption         =   "Label8"
      Begin VB.Menu ViewList 
         Caption         =   "View List"
      End
      Begin VB.Menu trigger 
         Caption         =   "Trigger"
      End
      Begin VB.Menu mnuok 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOff 
         Caption         =   "IMs Off"
      End
      Begin VB.Menu imoffa 
         Caption         =   "IMs On"
      End
      Begin VB.Menu stuff 
         Caption         =   "-"
      End
      Begin VB.Menu Exit_Mnu 
         Caption         =   "Exit "
      End
   End
   Begin VB.Menu Label9 
      Caption         =   "Label9"
      Begin VB.Menu IMsOn 
         Caption         =   "IMs On"
      End
      Begin VB.Menu IMsOff 
         Caption         =   "IMs Off"
      End
      Begin VB.Menu asdajshdfkjhsa 
         Caption         =   "-"
      End
      Begin VB.Menu Enter 
         Caption         =   "Enter PR"
      End
      Begin VB.Menu RoomBuster 
         Caption         =   "Room Buster"
      End
      Begin VB.Menu okthen 
         Caption         =   "-"
      End
      Begin VB.Menu Exitonthe 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_a_Click()

End Sub

Private Sub afk_Click()
Form24.Show
End Sub

Private Sub AsciiShop_Click()
Form30.Show
End Sub

Private Sub Attention_Click()
Form25.Show
End Sub

Private Sub Ball_Click()
Form23.Show
End Sub

Private Sub Chat_Clear_Click()
Call Chat_Clear
End Sub

Private Sub Clear_Click()
Text1 = ""
End Sub

Private Sub Color_Coder_Click()
Form12.Show
End Sub

Private Sub Copy_Click()
Clipboard.Clear

Timeout (1.15)

End Sub

Private Sub Credits_Click()
Form21.Show
End Sub

Private Sub Cut_Click()
Clipboard.Clear
Clipboard.SetText Code.Text
Code.Text = ""
Timeout (1.15)

End Sub

Private Sub Echo_Bot_Click()
Form8.Show
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Exit_Mnu_Click()
Unload Me
End Sub

Private Sub Exitonthe_Click()
Unload Me
End Sub

Private Sub Font1_Click()
Form6.Show
End Sub

Private Sub Hide_Aol_Click()
If Hide_Aol.Caption = "Hide AOL" Then
Call HideAOL
Hide_Aol.Caption = "Show AOL"
Else
Call ShowAOL
Hide_Aol.Caption = "Hide AOL"
End If
End Sub

Private Sub I_Talker_Click()
Form14.Show
End Sub

Private Sub Im_Off_Click()
Call IMKeyword("$IM_OFF", "Full Moon,Imz Off ")
End Sub

Private Sub Im_on_Click()
Call IMKeyword("$IM_ON", "Full Moon,Imz On")
End Sub

Private Sub killer_Click()
Form32.Show
End Sub

Private Sub Link_Sender_Click()
Form15.Show
End Sub

Private Sub Macro_Shop_Click()
Form4.Show
End Sub

Private Sub Mail_Room_Click()
Form3.Show

End Sub

Private Sub New_Game_Click()
'-------------------------------------------------------
'Disable the new game and high score menus while
'enabling the end game and pause menus.
'-------------------------------------------------------
mnuNewGame.Enabled = False
mnuEndGame.Enabled = True
mnuHighScore.Enabled = False
mnuPause.Enabled = True
'-------------------------------------------------------
'Start the game
'-------------------------------------------------------
NewGame
End Sub

Private Sub Mail_server_Click()
Form9.Show

End Sub

Private Sub MailUs_Click()
Form27.Show
End Sub

Private Sub Mass_Mailer_Click()
Form26.Show

End Sub

Private Sub MChat_Click()
Form31.Show

End Sub

Private Sub Minimize_Click()

Form22.Show
Unload Form1
End Sub

Private Sub mnuAbout_Click()
'-------------------------------------------------------
'Display the about form and pause the game if one is in
'progress
'-------------------------------------------------------
If Board.Game Then
    PauseTheGame = True
End If
frmAbout.Show 1

End Sub

Private Sub mnuEndGame_Click()
'-------------------------------------------------------
'End the current game
'-------------------------------------------------------
GameOver = True

End Sub

Private Sub mnuExit_Click()
On Error Resume Next
'-------------------------------------------------------
'Unload the form to call the Form_Unload procedure.
'Attempting to end any other way has resulted in many
'a fatal error in VB32.EXE
'-------------------------------------------------------
Unload Me

End Sub

Private Sub Mnuexitme_Click()
End
End Sub

Private Sub mnuHighScore_Click()
'-------------------------------------------------------
'Display the high scores
'-------------------------------------------------------
DisplayHighScores

End Sub

Private Sub MnuInstructions_Click()
'-------------------------------------------------------
'Display the instructions and pause the game if one is
'in progress
'-------------------------------------------------------
If Board.Game Then
    PauseTheGame = True
End If
frmInstruct.Show 1

End Sub

Private Sub mnuNewGame_Click()
'-------------------------------------------------------
'Disable the new game and high score menus while
'enabling the end game and pause menus.
'-------------------------------------------------------
mnuNewGame.Enabled = False
mnuEndGame.Enabled = True
mnuHighScore.Enabled = False
mnuPause.Enabled = True
'-------------------------------------------------------
'Start the game
'-------------------------------------------------------
NewGame
End Sub

Private Sub mnuOptions_Click()
'-------------------------------------------------------
'Display the options form, fills in the options
'accordingly, and pause the game if one is in progress
'-------------------------------------------------------
If Board.Game Then
    PauseTheGame = True
End If
frmOptions.txtStartingLevel = StartingLevel
If FillLines Then
    frmOptions.chkFillLines.Value = 1
Else
    frmOptions.chkFillLines.Value = 0
End If
If PlaySounds Then
    frmOptions.chkPlaySounds.Value = 1
Else
    frmOptions.chkPlaySounds.Value = 0
End If
If HideSplash Then
    frmOptions.chkSkipIntro.Value = 1
Else
    frmOptions.chkSkipIntro.Value = 0
End If
frmOptions.Show 1

End Sub

Private Sub mnuPause_Click()
'-------------------------------------------------------
'Pause or unpause the game
'-------------------------------------------------------
PauseTheGame = Not (PauseTheGame)

End Sub

Private Sub Options_Yas_Click()

End Sub

Private Sub Pause_me_Click()
'-------------------------------------------------------
'Pause or unpause the game
'-------------------------------------------------------
PauseTheGame = Not (PauseTheGame)

End Sub

Private Sub Phish_Manager_Click()
Form19.Show
End Sub


Private Sub cmdPrint_Click()
  Dim TotalPages As Integer
  Dim PageCount As String
  ' Specifies the text position to print
  Printer.CurrentX = 100
  Printer.CurrentY = 100
  Printer.Print "This appears at the top of the page"

  ' Specifies a line width and location
  Printer.DrawWidth = 3
  Printer.Line (100, 100)-(10000, 100)
  Printer.Line (100, 350)-(10000, 350)
  TotalPages = Printer.Page

  ' Specifies where to print the page count
  Printer.CurrentX = 1000
  Printer.CurrentY = 400
  PageCount = "TotalPages = " & Str$(TotalPages)
  Printer.Print PageCount
  Printer.EndDoc
End Sub


Private Sub Print_Click()

  Dim TotalPages As Integer
  Dim PageCount As String
  ' Specifies the text position to print
  Printer.CurrentX = 100
  Printer.CurrentY = 100
  Printer.Print "This appears at the top of the page"

  ' Specifies a line width and location
  Printer.DrawWidth = 3
  Printer.Line (100, 100)-(10000, 100)
  Printer.Line (100, 350)-(10000, 350)
  TotalPages = Printer.Page

  ' Specifies where to print the page count
  Printer.CurrentX = 1000
  Printer.CurrentY = 400
  PageCount = "TotalPages = " & Str$(TotalPages)
  Printer.Print PageCount
  Printer.EndDoc

End Sub

Private Sub Pw_Gen_Click()
Form5.Show
End Sub

Private Sub QuickKW_Click()
Form29.Show
End Sub

Private Sub Reset_Yaa_Click()
Form7.Show
End Sub

Private Sub Room_Anoy_Click()
Form16.Show
End Sub

Private Sub Room_Buster_Click()
Form17.Show
End Sub

Private Sub RoomBuster_Click()
Form17.Show
End Sub

Private Sub Scroll_Click()
SendChat "" + Text1 + ""
Text1 = ""
End Sub

Private Sub Server_Hlp_Click()
Form10.Show
End Sub

Private Sub SnDecoder_Click()
Form18.Show
End Sub

Private Sub Tetris_Here_Click()
frmSplash.Show
End Sub

Private Sub TextColor_Click()
CommonDialog1.FLAGS = cdlCCRGBInit
  CommonDialog1.ShowColor
  Text1.ForeColor = CommonDialog1.Color

End Sub

Private Sub Up_Chat_Click()
If upchat1.Caption = "Upchat" Then
Call Upchat
upchat1.Caption = "Un-Upchat"
Else
Call UnUpchat
upchat1.Caption = "Upchat"
End If
End Sub

Private Sub ViewList_Click()
Form28.Show

End Sub

Private Sub Voter_Bot_Click()
Form13.Show
End Sub

Private Sub WarezRequest_Click()
Form20.Show
End Sub

Private Sub Web_Page_Click()
Call XAOL4_Keyword("http://www.rip-inc.com")
End Sub
