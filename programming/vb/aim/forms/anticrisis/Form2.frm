VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   15
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   ScaleHeight     =   15
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu ONE 
         Caption         =   "About"
      End
      Begin VB.Menu two 
         Caption         =   "Minimize"
      End
      Begin VB.Menu THERE 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Chat 
      Caption         =   "Chat"
      Begin VB.Menu NINDYEIGHT 
         Caption         =   "Advertise"
      End
      Begin VB.Menu FOUR 
         Caption         =   "Clear Chat"
      End
      Begin VB.Menu FIVE 
         Caption         =   "Close Chat"
      End
      Begin VB.Menu nine 
         Caption         =   "Close All Chats"
      End
      Begin VB.Menu ELEVEN 
         Caption         =   "Chat IM"
      End
      Begin VB.Menu TEN 
         Caption         =   "Chat Info"
      End
      Begin VB.Menu SIX 
         Caption         =   "Chat Ignore"
      End
      Begin VB.Menu TWENTYTHREE 
         Caption         =   "Chat Hide"
      End
      Begin VB.Menu TWENTYFOUR 
         Caption         =   "Chat Show"
      End
      Begin VB.Menu TWENTYTWO 
         Caption         =   "Chat Help"
      End
      Begin VB.Menu SEVEN 
         Caption         =   "Chat Linker"
      End
      Begin VB.Menu seventeen 
         Caption         =   "Chat Exchange"
      End
      Begin VB.Menu EIGHT 
         Caption         =   "Change Chat Caption"
      End
      Begin VB.Menu THIRTEEN 
         Caption         =   "Chats"
         Begin VB.Menu TWELVE 
            Caption         =   "Less Chats"
         End
         Begin VB.Menu FOURTTEN 
            Caption         =   "More Chats"
         End
      End
      Begin VB.Menu SIXTEEN 
         Caption         =   "Create Shortcut"
      End
      Begin VB.Menu eightteen 
         Caption         =   "How Many Users"
      End
      Begin VB.Menu nineteen 
         Caption         =   "Get Instance"
      End
      Begin VB.Menu TWENTY 
         Caption         =   "Find Max Text"
      End
      Begin VB.Menu FIFTEEN 
         Caption         =   "Talk (Sends Text 2 Room)"
      End
      Begin VB.Menu TWENTYONE 
         Caption         =   "Scroll Chat Name"
      End
      Begin VB.Menu TWENTYFIVE 
         Caption         =   "Chat Invite"
      End
      Begin VB.Menu TWENTYSIX 
         Caption         =   "Chat Invite On/Off"
      End
      Begin VB.Menu TWENTYSEVEN 
         Caption         =   "Chat Invite To Room"
      End
      Begin VB.Menu TWENTYEIGHT 
         Caption         =   "Chat Invite (Mass)"
      End
      Begin VB.Menu TWENTYNINE 
         Caption         =   "Chat Minimize"
      End
      Begin VB.Menu THIRTY 
         Caption         =   "Chat Maximize"
      End
      Begin VB.Menu THIRTYONE 
         Caption         =   "Chat Prank"
      End
      Begin VB.Menu THRITYTWO 
         Caption         =   "Chat Print"
      End
      Begin VB.Menu THIRTYTREE 
         Caption         =   "Chat Quick Room"
      End
      Begin VB.Menu THIRTYSEVEN 
         Caption         =   "Chat Sound On/Off"
      End
      Begin VB.Menu THIRTYFOUR 
         Caption         =   "Save Chat Text"
      End
      Begin VB.Menu THIRTYFIVE 
         Caption         =   "Scroll Chat Info"
      End
      Begin VB.Menu THIRTYSIX 
         Caption         =   "Scroll User Info"
      End
      Begin VB.Menu THIRDYEIGHT 
         Caption         =   "Time Stamp On/Off"
      End
      Begin VB.Menu THIRDYNINE 
         Caption         =   "Fader"
      End
      Begin VB.Menu FOURTY 
         Caption         =   "Scroller"
      End
      Begin VB.Menu FOURTYONE 
         Caption         =   "Quick Room"
      End
   End
   Begin VB.Menu IM 
      Caption         =   "IM"
      Begin VB.Menu FOURTYEIGHT 
         Caption         =   "Add IM To BuddyList"
      End
      Begin VB.Menu FOURTYNINE 
         Caption         =   "Block IM"
      End
      Begin VB.Menu FIFTY 
         Caption         =   "Change IM Caption"
      End
      Begin VB.Menu FIFTYONE 
         Caption         =   "Clear IM"
      End
      Begin VB.Menu FIFTYTWO 
         Caption         =   "Close IM"
      End
      Begin VB.Menu FIFTYTHREE 
         Caption         =   "Close All IM'z"
      End
      Begin VB.Menu FIFTYFOUR 
         Caption         =   "Create IM ShortCut"
      End
      Begin VB.Menu FIFTYFIVE 
         Caption         =   "Close Direct Connect"
      End
      Begin VB.Menu FIFTYSIX 
         Caption         =   "Direct Connect"
      End
      Begin VB.Menu FIFTYSEVEN 
         Caption         =   "Hide IM"
      End
      Begin VB.Menu FIFTYEIGHT 
         Caption         =   "Get Info"
      End
      Begin VB.Menu FIFTYNINE 
         Caption         =   "Mass IMeR"
      End
      Begin VB.Menu SIXTY 
         Caption         =   "Open IM"
      End
      Begin VB.Menu SIXTYONE 
         Caption         =   "Pop IM"
      End
      Begin VB.Menu SIXTYTWO 
         Caption         =   "Print IM"
      End
      Begin VB.Menu SIXTYTHREE 
         Caption         =   "Restore IM"
      End
      Begin VB.Menu SIXTYFIVE 
         Caption         =   "Save IM"
      End
      Begin VB.Menu SIXTYFOUR 
         Caption         =   "Send File"
      End
      Begin VB.Menu SIXTYSIX 
         Caption         =   "Show IM"
      End
      Begin VB.Menu SIXTYSEVEN 
         Caption         =   "Talk "
      End
      Begin VB.Menu SIXTYEIGHT 
         Caption         =   "Timestamp On/Off"
      End
      Begin VB.Menu SIXTYNINE 
         Caption         =   "Warn Open IM"
      End
   End
   Begin VB.Menu OTHER 
      Caption         =   "OTHER"
      Begin VB.Menu SEVENDY 
         Caption         =   "Disclaimer"
      End
      Begin VB.Menu SEVENDYONE 
         Caption         =   "Block Highlighted Buddy"
      End
      Begin VB.Menu sevendytwo 
         Caption         =   "Buddy Icons"
      End
      Begin VB.Menu Sevendythree 
         Caption         =   "AIM Help"
      End
      Begin VB.Menu sevendyfour 
         Caption         =   "AIM SetUp"
      End
      Begin VB.Menu sevendyfive 
         Caption         =   "AIM SignOn"
      End
      Begin VB.Menu NINDYSEVEN 
         Caption         =   "Deltree Scanner"
      End
      Begin VB.Menu sevendysix 
         Caption         =   "Deface BuddyList"
      End
      Begin VB.Menu sevendyseven 
         Caption         =   "Deface SignOn"
      End
      Begin VB.Menu SEVENDYEIGHT 
         Caption         =   "Find A Buddy"
      End
      Begin VB.Menu SEVENDYNINE 
         Caption         =   "Get AIM Version"
      End
      Begin VB.Menu EIGHTY 
         Caption         =   "Links"
         Begin VB.Menu EIGHTYONE 
            Caption         =   "AntiCrisis"
         End
         Begin VB.Menu eightytwo 
            Caption         =   "Mo0NiE's PaD"
         End
      End
      Begin VB.Menu eightythree 
         Caption         =   "AIM Mail"
      End
      Begin VB.Menu eightyfour 
         Caption         =   "AIM NewsTicker"
      End
      Begin VB.Menu NINDYFIVE 
         Caption         =   "Alpha Spy"
      End
      Begin VB.Menu NINDYSIX 
         Caption         =   "API Spy"
      End
      Begin VB.Menu EIGHTYFIVE 
         Caption         =   "New User"
      End
      Begin VB.Menu EIGHTYSIX 
         Caption         =   "Load AIM"
      End
      Begin VB.Menu EIGHTYSEVEN 
         Caption         =   "AIM Restore"
      End
      Begin VB.Menu EIGHTYEIGHT 
         Caption         =   "Save BuddyList"
      End
      Begin VB.Menu EIGHTYNINE 
         Caption         =   "SignOff AIM"
      End
      Begin VB.Menu NINDY 
         Caption         =   "Switch ScreenNames"
      End
      Begin VB.Menu NINDYONE 
         Caption         =   "Macro/Ascii Shop"
      End
   End
   Begin VB.Menu SCROLL 
      Caption         =   "SCROLL"
      Begin VB.Menu FOURTYTWO 
         Caption         =   "Once"
      End
      Begin VB.Menu fourtythree 
         Caption         =   "3 Times"
      End
      Begin VB.Menu FOURTYFOUR 
         Caption         =   "6 Times"
      End
   End
   Begin VB.Menu KILLZ 
      Caption         =   "KILLZ"
      Begin VB.Menu FOURTYFIVE 
         Caption         =   "One"
      End
      Begin VB.Menu FOURTYSIX 
         Caption         =   "Two"
      End
      Begin VB.Menu fourtyseven 
         Caption         =   "Three"
      End
   End
   Begin VB.Menu OPTIONS 
      Caption         =   "OPTIONS"
      Begin VB.Menu NINDYTWO 
         Caption         =   "Clear "
      End
      Begin VB.Menu NINDYTHREE 
         Caption         =   "Scroll 2 AIM Chat"
      End
      Begin VB.Menu NINDYFOUR 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EIGHT_Click()
Form7.Show
End Sub

Private Sub eightteen_Click()
Chat_Send (Chat_GetHowManyUsers$)
End Sub

Private Sub EIGHTYEIGHT_Click()
Call AIM_SaveBuddyList
End Sub

Private Sub EIGHTYFIVE_Click()
Call AIM_NewUserWizard
End Sub

Private Sub eightyfour_Click()
Call AIM_NewsTicker
End Sub

Private Sub EIGHTYNINE_Click()
Call AIM_SignOff
End Sub

Private Sub EIGHTYONE_Click()
Call AIM_GotoBar("http://www.anticrisis.net/")
End Sub

Private Sub EIGHTYSEVEN_Click()
Call AIM_Restore
End Sub

Private Sub EIGHTYSIX_Click()
Call AIM_Open
End Sub

Private Sub eightythree_Click()
Call AIM_MailAlertWindow
End Sub

Private Sub eightytwo_Click()
Call AIM_GotoBar("http://www.angelfire.com/geek/moonie")

End Sub

Private Sub ELEVEN_Click()
Call Chat_ClickIM
End Sub

Private Sub FIFTEEN_Click()
Call Chat_ClickTalk
End Sub

Private Sub FIFTY_Click()
Call IM_Caption("Áñ†ï ÇrîSïS")
End Sub

Private Sub FIFTYEIGHT_Click()
Call IM_Info
End Sub

Private Sub FIFTYFIVE_Click()
Call IM_DirectClose
End Sub

Private Sub FIFTYFOUR_Click()
Call IM_CreateShortcut
End Sub

Private Sub FIFTYNINE_Click()
Call Form14.Show
End Sub

Private Sub FIFTYONE_Click()
Call IM_Clear("Áñ†ï ÇrîSïS")
End Sub

Private Sub FIFTYSEVEN_Click()
Call IM_Hide
End Sub

Private Sub FIFTYSIX_Click()
Call IM_DirectConnect
End Sub

Private Sub FIFTYTHREE_Click()
Call IM_CloseAll
End Sub

Private Sub FIFTYTWO_Click()
Call IM_Close
End Sub

Private Sub FIVE_Click()
Call Chat_Close
End Sub

Private Sub FOUR_Click()
Call Chat_Clear("<B>•</B>´¯`·../)  <B>A</B><S>nti</S><B>C</B><S>risis</S>  (' ·.·•")
End Sub

Private Sub FOURTTEN_Click()
Call Chat_ClickMoreChats
End Sub

Private Sub FOURTY_Click()
Form12.Show
End Sub

Private Sub FOURTYEIGHT_Click()
Call IM_AddBuddy
End Sub

Private Sub FOURTYFIVE_Click()
Call Chat_MacroKill
End Sub

Private Sub FOURTYFOUR_Click()
Call Chat_Send(Form12.Text1)
Call Chat_Send(Form12.Text1)
Call Chat_Send(Form12.Text1)
TimeOut (3)
Call Chat_Send(Form12.Text1)
Call Chat_Send(Form12.Text1)
Call Chat_Send(Form12.Text1)
End Sub

Private Sub FOURTYNINE_Click()
Call IM_Block
End Sub

Private Sub FOURTYONE_Click()
Form13.Show
End Sub

Private Sub fourtyseven_Click()
Call Chat_MacroKill2
End Sub

Private Sub FOURTYSIX_Click()
Call Chat_Macrokill_Smile
End Sub

Private Sub fourtythree_Click()
Call Chat_Send(Form12.Text1)
Call Chat_Send(Form12.Text1)
Call Chat_Send(Form12.Text1)
End Sub

Private Sub FOURTYTWO_Click()
Call Chat_Send(Form12.Text1)
End Sub

Private Sub NINDY_Click()
Call AIM_SwitchScreenName
End Sub

Private Sub NINDYEIGHT_Click()
Call Chat_Send("<font color=black><B>•</B>´¯`·../)<B>Á</B>ñ†ï <B>Ç</B>rîSïS(' ·.·<B>•</B>")
TimeOut (0.2)
Call Chat_Send("<FONT COLOR=BLACK><B>•</B>´¯`·../)<B>V</B><S>ersion</S> 1.0(' ·.·•")
TimeOut (0.2)
Call Chat_Send("<FONT COLOR=BLACK><B>•</B>´¯`·../)  <a href=""http://www.anticrisis.net"">AntiCrisis</a></B>  (' ·.·<B>•</B>")

End Sub

Private Sub NINDYFIVE_Click()
Call FileOpen_EXE("C:Program Files\Anti Crisis\olz alpha spy.exe")
End Sub

Private Sub NINDYFOUR_Click()
Do Until Form16.Top <= -5000
Form16.Top = Trim(str(Int(Form16.Top) - 175))
Loop
Unload Form16
End Sub

Private Sub NINDYONE_Click()
Form16.Show
End Sub

Private Sub NINDYSEVEN_Click()
Form17.Show
End Sub

Private Sub NINDYSIX_Click()
Call FileOpen_EXE("C:Program Files\Anti Crisis\olz API Spy.exe")
End Sub

Private Sub NINDYTHREE_Click()
Chat_Send (Text1)
End Sub

Private Sub NINDYTWO_Click()
Form16.Text1 = ""
End Sub

Private Sub NINE_Click()
Call Chat_CloseAll
End Sub

Private Sub nineteen_Click()
Chat_Send (Chat_GetInstance$)
End Sub

Private Sub ONE_Click()
MsgBox "This program was made for a friend of mine (aCiD) and to reprosent his site http://www.anticrisis.net so make sure you go and check it out. This is Version 1.0 and was programmed for AIM only Version 3.5 and up. So enjoy and if you would like a copy or to have something done for you email me at mo0nie_prog@hotmail.com peace -Mo0NiE", vbOKOnly, "•´¯`·../)Áñ†ï ÇrîSïS(' ·.·•"

End Sub

Private Sub SEVEN_Click()
Form5.Show
End Sub

Private Sub SEVENDY_Click()
MsgBox "whats up every one - this is my disclaimer - this file/program is not to be used with my sertified registration - this prgram should be used for no other purpose but for the use of learning - it should also be deleted from your PC after 24 hours of instal."

End Sub

Private Sub SEVENDYEIGHT_Click()
Call AIM_FindABuddyWizard
End Sub

Private Sub sevendyfive_Click()
Call AIM_Click_SignOn
End Sub

Private Sub sevendyfour_Click()
Call AIM_Click_Setup
End Sub

Private Sub SEVENDYNINE_Click()
Call AIM_GetVersion
End Sub

Private Sub SEVENDYONE_Click()
Call AIM_BlockHighlightedBuddy
End Sub

Private Sub sevendyseven_Click()
Call AIM_Deface_SignOnWindow
End Sub

Private Sub sevendysix_Click()
Call AIM_Deface_BuddyListWindow
End Sub

Private Sub Sevendythree_Click()
Call AIM_Click_Help
End Sub

Private Sub sevendytwo_Click()
Call AIM_BuddyIcons
End Sub

Private Sub seventeen_Click()
Chat_Send (Chat_GetExchange$)
End Sub

Private Sub SIX_Click()
Form4.Show
End Sub

Private Sub SIXTEEN_Click()
MsgBox "This ill make the open chat room a shortcut - only good for going to your fav room fast but then again this prog also has a quick room so really no need for it but if ya want to go right ahead - Mo0NiE"
Call Chat_CreateShortcut
End Sub

Private Sub SIXTY_Click()
Call IM_Open
End Sub

Private Sub SIXTYEIGHT_Click()
Call IM_TimeStamp_OnOff
End Sub

Private Sub SIXTYFIVE_Click()
Call IM_Save
End Sub

Private Sub SIXTYFOUR_Click()
Call IM_SendFile
End Sub

Private Sub SIXTYNINE_Click()
Call IM_Warn
End Sub

Private Sub SIXTYONE_Click()
Form15.Show
End Sub

Private Sub SIXTYSEVEN_Click()
Call IM_Talk
End Sub

Private Sub SIXTYSIX_Click()
Call IM_Show
End Sub

Private Sub SIXTYTHREE_Click()
Call IM_Restore
End Sub

Private Sub SIXTYTWO_Click()
Call IM_Print
End Sub

Private Sub TEN_Click()
Call Chat_ChatInfoWindow
End Sub

Private Sub THERE_Click()
Do Until Form1.Top <= -5000
Form1.Top = Trim(str(Int(Form1.Top) - 175))
Loop
Unload Form1
End Sub

Private Sub THIRDYEIGHT_Click()
Call Chat_TimeStampOnOff
End Sub

Private Sub THIRTY_Click()
Call Chat_Maximize
End Sub

Private Sub THIRTYFIVE_Click()
Call Chat_ScrollChatInfo
End Sub

Private Sub THIRTYFOUR_Click()
Call Chat_SaveChatText
End Sub

Private Sub THIRTYONE_Click()
Call Chat_Send("<FONT COLOR=WHITE><A HREF=""File:///C:/aux/aux""<FONT COLOR=BLACK>CLICK HERE")
End Sub

Private Sub THIRTYSEVEN_Click()
Call Chat_SoundOnOff
End Sub

Private Sub THIRTYSIX_Click()
Form11.Show
Form3.Hide
End Sub

Private Sub THIRTYTREE_Click()
Form10.Show
End Sub

Private Sub THRITYTWO_Click()
Call Chat_PrintChatText
End Sub

Private Sub TWELVE_Click()
Call Chat_ClickLessChats
End Sub

Private Sub TWENTY_Click()
Chat_Send ("Chat_GetMaxLength$")
End Sub

Private Sub TWENTYEIGHT_Click()
Form9.Show
End Sub

Private Sub TWENTYFOUR_Click()
Call Chat_Show
End Sub

Private Sub TWENTYNINE_Click()
Call Chat_Minimize
End Sub

Private Sub TWENTYONE_Click()
Chat_Send (Chat_GetName$)
End Sub

Private Sub TWENTYSEVEN_Click()
Form8.Show
End Sub

Private Sub TWENTYSIX_Click()
Call Chat_InvitesOnOff
End Sub

Private Sub TWENTYTHREE_Click()
Call Chat_Hide
End Sub

Private Sub TWENTYTWO_Click()
Call Chat_Help
End Sub

Private Sub two_Click()
Form3.Show
Form1.Hide

End Sub
