VERSION 5.00
Begin VB.Form MenuForm 
   Caption         =   "MenuForm"
   ClientHeight    =   465
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4695
   HelpContextID   =   80
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menuPrefs 
      Caption         =   "&Prefs"
      HelpContextID   =   910
      Begin VB.Menu menuFile 
         Caption         =   "&File"
         Begin VB.Menu itemLog 
            Caption         =   "Save To &Log File"
         End
         Begin VB.Menu itemSave 
            Caption         =   "&Save Settings"
         End
         Begin VB.Menu itemSaveExit 
            Caption         =   "Save Settings At &Exit"
         End
         Begin VB.Menu itemTest 
            Caption         =   "Server &Test"
         End
      End
      Begin VB.Menu itemstep 
         Caption         =   "-"
         HelpContextID   =   960
      End
      Begin VB.Menu itemServer 
         Caption         =   "&Server "
         Begin VB.Menu itemview 
            Caption         =   "&View Total Sent Since Install "
         End
      End
      Begin VB.Menu menuMail 
         Caption         =   "&Mail Type"
         Begin VB.Menu itemNewMail 
            Caption         =   "&New Mail"
         End
         Begin VB.Menu itemOld 
            Caption         =   "Mail You've &Read"
         End
         Begin VB.Menu itemSent 
            Caption         =   "Mail You've &Sent"
         End
         Begin VB.Menu itemFlashMail 
            Caption         =   "Incoming &FlashMail"
         End
      End
      Begin VB.Menu subNotify 
         Caption         =   "&Notify"
         Begin VB.Menu ItemChat 
            Caption         =   "&Chat"
         End
         Begin VB.Menu itemIm 
            Caption         =   "&IM"
         End
      End
      Begin VB.Menu SubScroll 
         Caption         =   "Scro&lls"
         Begin VB.Menu ItemScroll 
            Caption         =   "&Scroll ""Server"" Commands"
         End
         Begin VB.Menu separaform26 
            Caption         =   "-"
         End
         Begin VB.Menu itembanned 
            Caption         =   "Scroll ""Banned"" Screen Names"
         End
      End
      Begin VB.Menu itemServing 
         Caption         =   "Ser&ving Styles"
         Begin VB.Menu ItemWrite 
            Caption         =   "&Write ""Sending""When Processing A Mail"
         End
         Begin VB.Menu separaform8 
            Caption         =   "-"
            HelpContextID   =   1050
         End
         Begin VB.Menu itemRemove 
            Caption         =   "&Remove ""Fwd:"" From Forwarded Mails"
         End
         Begin VB.Menu separaform30 
            Caption         =   "-"
         End
         Begin VB.Menu ItemKill 
            Caption         =   "&Kill Wait After Click Send"
         End
         Begin VB.Menu Separaform9 
            Caption         =   "-"
            HelpContextID   =   1060
         End
         Begin VB.Menu itemStatus 
            Caption         =   "&Scroll Status After List Is Sends"
         End
      End
      Begin VB.Menu separaform6 
         Caption         =   "-"
         HelpContextID   =   970
      End
      Begin VB.Menu menuIMs 
         Caption         =   "&IMz"
         Begin VB.Menu itemOff 
            Caption         =   "&IM's Off"
         End
         Begin VB.Menu itemon 
            Caption         =   "&IM's On"
         End
      End
      Begin VB.Menu menuKill 
         Caption         =   "&Killz"
         Begin VB.Menu itemBuddy 
            Caption         =   "Kill &Buddy List"
         End
         Begin VB.Menu itemwait 
            Caption         =   "Kill &Wait"
         End
      End
      Begin VB.Menu itemIdleBot 
         Caption         =   "Idle &Bot"
      End
      Begin VB.Menu Step 
         Caption         =   "-"
         HelpContextID   =   980
      End
      Begin VB.Menu itemComments 
         Caption         =   "Mail &Comment(s)"
      End
   End
   Begin VB.Menu menuOther 
      Caption         =   "&Others"
      HelpContextID   =   920
      Begin VB.Menu itemDuplicated 
         Caption         =   "Kill &Dupes Mails"
      End
      Begin VB.Menu separaform21 
         Caption         =   "-"
         HelpContextID   =   1000
      End
      Begin VB.Menu ItemChatServe 
         Caption         =   "&Chat && Serve"
      End
      Begin VB.Menu itemBust 
         Caption         =   "&Room Bust"
      End
      Begin VB.Menu Step21 
         Caption         =   "-"
         HelpContextID   =   1020
      End
      Begin VB.Menu itemHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu itemInfo 
         Caption         =   "I&nFo"
      End
      Begin VB.Menu step3 
         Caption         =   "-"
         HelpContextID   =   1030
      End
      Begin VB.Menu itemmini 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu itemexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menuIgnore 
      Caption         =   "&Ignore"
      HelpContextID   =   930
      Begin VB.Menu ItemName0 
         Caption         =   "ItemName0"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName1 
         Caption         =   "ItemName1"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName2 
         Caption         =   "ItemName2"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName3 
         Caption         =   "ItemName3"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName4 
         Caption         =   "ItemName4"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName5 
         Caption         =   "ItemName5"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName6 
         Caption         =   "ItemName6"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName7 
         Caption         =   "ItemName7"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName8 
         Caption         =   "ItemName8"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName9 
         Caption         =   "ItemName9"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName10 
         Caption         =   "ItemName10"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName11 
         Caption         =   "ItemName11"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName12 
         Caption         =   "ItemName12"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName13 
         Caption         =   "ItemName13"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName14 
         Caption         =   "ItemName14"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName15 
         Caption         =   "ItemName15"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName16 
         Caption         =   "ItemName16"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName17 
         Caption         =   "ItemName17"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName18 
         Caption         =   "ItemName18"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName19 
         Caption         =   "ItemName19"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName20 
         Caption         =   "ItemName20"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName21 
         Caption         =   "ItemName21"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName22 
         Caption         =   "ItemName22"
         Visible         =   0   'False
      End
      Begin VB.Menu ItemName23 
         Caption         =   "ItemName23"
         Visible         =   0   'False
      End
      Begin VB.Menu itemSN 
         Caption         =   "(Screen Names)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "&Help"
      HelpContextID   =   940
      Begin VB.Menu itemdisclaimer 
         Caption         =   "&Disclaimer"
      End
      Begin VB.Menu itemIntro 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu itemcommands 
         Caption         =   "&Server Commands"
      End
      Begin VB.Menu itemEnd 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub itemAdvertize_Click()
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
    MsgBox "America Online Is Not Open.", 16
    Exit Sub
End If
Wel = FindChildByTitle(AOLMDI(), "Welcome,")
Welc$ = String(255, 0)
WhichWel = GetWindowText(Wel, Welc$, 250)
If WhichWel < 8 Then
    MsgBox "You Need To Sign On First To Use The Server.", 16
    Exit Sub
End If
Room = AOLFindChatRoom
If Room = 0 Then
  MsgBox "You Got Be In a Chat Room!", 16
 Exit Sub
End If
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & App.Title): DoEvents
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "If You Dont Find This Program"): DoEvents
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Dynamically Superior To All Other Servers"): DoEvents
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "You'Re Being Way To Critical"): DoEvents
End Sub

Private Sub itemAFK_Click()
If ServerBot = True Then MsgBox "Turn Off The Server To Use The AFK Bot", 16: Exit Sub
AFK.Show
End Sub

Private Sub itemBanned_Click()
Dim AOL
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
    MsgBox "America Online Is Not Open.", 16
    Exit Sub
End If
Wel = FindChildByTitle(AOLMDI(), "Welcome,")
Welc$ = String(255, 0)
WhichWel = GetWindowText(Wel, Welc$, 250)
If WhichWel < 8 Then
    MsgBox "You Need To Sign On First To Use The Server.", 16
    Exit Sub
End If
 Room = AOLFindChatRoom
If Room = 0 Then
  MsgBox "You Got Be In a Chat Room!", 16
 Exit Sub
End If
If Server.List6.ListCount = 0 Then Exit Sub
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "- X-Treme Server '98 Banned Peoples! -")
For i = 0 To Server.List6.ListCount - 1
 ScreenNames$ = Server.List6.List(i)
 SendChat ("<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & " FACE=""Arial Narrow""><B>" & i & " )-" & ScreenNames$)
 Next i
End Sub

Private Sub itemBuddy_Click()
Dim AOL As Long, mdi As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
If AOL& = 0& Then
  MsgBox "America Online Is Not Loaded", 16
  Exit Sub
End If
  Wel = FindChildByTitle(mdi&, "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
   MsgBox "You Need To Sign On First To Use The Server", 16
   Exit Sub
 End If
 buddy& = FindChildByTitle(mdi&, "Buddy List Window")
 If buddy& = 0& Then Exit Sub
 If itemBuddy.Caption = "Kill &Buddy List" Then
   itemBuddy.Caption = "Show &Buddy List"
   Call ShowWindow(buddy&, SW_HIDE)
 Else
   itemBuddy.Caption = "Kill &Buddy List"
   Call ShowWindow(buddy&, SW_SHOWNORMAL)
 End If
 End Sub

Private Sub itemBug_Click()
Bug.Show
End Sub

Private Sub itemBust_Click()
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
  MsgBox "America Online Is Not Loaded", 16
  Exit Sub
End If
  Wel = FindChildByTitle(AOLMDI(), "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
  MsgBox "You Need To Sign On First To Use The Server", 16
  Exit Sub
 End If
Bust.Show
End Sub

Private Sub ItemChat_Click()
itemIm.Checked = False
ItemChat.Checked = True
End Sub

Private Sub ItemChatServe_Click()
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
  MsgBox "America Online Is Not Loaded", 16
  Exit Sub
End If
  Wel = FindChildByTitle(AOLMDI(), "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
  MsgBox "You Need To Sign On First To Use The Server", 16
  Exit Sub
 End If
 Room& = AOLFindChatRoom&
 If Room& = 0 Then MsgBox "You Have To Be In A Chat Room", 32: Exit Sub
 frmChat.Show
 StayOnTop frmChat
 CenterFormTop frmChat
End Sub

Private Sub itemcommands_Click()

MsgText = MsgText & "[Preference] Click On The Preferences Button To Display The Menu." & Chr(13) & Chr(10)
MsgText = MsgText & "[File]Allow You To Save The Settings,Create A Log File,Test The Server." & Chr(13) & Chr(10)
MsgText = MsgText & "[Mail Type]Allow You To Select What Mail You Want To Server[Fully Tested It Works.New Mail|Old Mails|Sent Mails|Incomming/Saved Flash Mails." & Chr(13) & Chr(10)
MsgText = MsgText & "[Notify]Allow You To Select How Do You Want The Server To Inform That The Mail Was Sent." & Chr(13) & Chr(10)
MsgText = MsgText & "[Scroll]Allow To Send The Server Commands (Note:This Is Done Every Minute Automaticly)" & Chr(13) & Chr(10)
MsgText = MsgText & "[Serving Style]There is 3 Commands.You Should know What They Do.[dont You?]" & Chr(13) & Chr(10)
MsgText = MsgText & "[Serve Through IM's]Allow You To Serve Only With Instant Message Just Like The Chat Room But The Commands Will Be Taken By The Im Window." & Chr(13) & Chr(10)
MsgText = MsgText & "[Im'z On And Off]Allow You To Turn On Your Instant Message Or Turn The Off."
MsgText = MsgText & "[Killz]Includes Kill Welcome(Kill The AnNoyn Welcome Window),Kill Buddy(Close Your Buddy List Window),Kill Wait(Kill America Online Timers & Stay On Line Message)." & Chr(13) & Chr(10)
MsgText = MsgText & "[Idle Bot]Same As Kill Wait But Automaticly."
MsgText = MsgText & "[Mail Comments]Allow You Insert Any Text On The Mails You Send.(Chat Message It Send The Text After The Mail Is Sent To The Chat Ex:Hey,[Dude] #2 Was Sent.Happy New Year)Get It?." & Chr(13) & Chr(10)
MsgText = MsgText & "[Advertize]I Added This,If You Want To Know What It's About Just Click On It.(Hey Dude Tell What You Think Please.Thank Tito)"
MsgText = MsgText & "[Bug Report & Suggetions]Please Use This Report Any Bugs To Make This Server Best On AOL,I Will Send You a Fix Pacth As Soon As I Fix The Problem.Of Course I Also Take FeedBacks & Suggestion Use This Option To Do It Faster & Easier For You.Thank You For Your Support." & Chr(13) & Chr(10)
MsgText = MsgText & "[Room Buster]Well You Should Know But,This Allow To Enter A Private Room If The Room Is Full It Will Trying To Get You In Every Second,To Stop The Proccess Click Stop." & Chr(13) & Chr(10)
MsgText = MsgText & "[Banned] Hum,This Allow You To Keep People Off The Server Just Add The Chat Room Screen Name By Clicking On The Room Button,Then Click On The Screen Name You Banned From The Server It Will Inform That Person That He/She Was Banned From The Server." & Chr(13) & Chr(10)
MsgBox MsgText, 64
End Sub

Private Sub itemComments_Click()
If itemComments.Checked Then
    itemComments.Checked = False
Else
    MailComment.Show
End If

End Sub

Private Sub itemdisclaimer_Click()
X = MsgBox("Freely Distruibuted Free Usage Granted For All.This Program Was Made For Educational Purposes Only.The Maker Of This Program Is Not Responsible For He/She Actions." & Chr(13) & Chr(10) & "This Program Can Not Directly Violate AOL's Terms Of Services On It's Own.So, You Assume Full Responsibility For Your Own Actions.", 64 + vbOKOnly)
End Sub

Private Sub itemDuplicated_Click()
    
    Dim MailBox As Long, AOL As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String, cSubject As String
    Dim SearchFor As Long, sSender As String, sSubject As String
    Dim CurCaption As String

AOL& = FindWindow("AOL Frame25", vbNullString)
If AOL& = 0& Then
  MsgBox "America Online Is Not Loaded", 16
  Exit Sub
End If
  Wel = FindChildByTitle(AOLMDI(), "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
  MsgBox "You Need To Sign On First To Use The Server", 16
  Exit Sub
 End If
    MailBox& = FindMailBox
    CurCaption$ = Server.Status.Caption
    If MailBox& = 0& Then MsgBox "Please Open Your Mail Box And Try Again.", 16: Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0& Then Exit Sub
    For SearchFor& = 0& To Count& - 2
        DoEvents
        sSender$ = MailSenderNew(SearchFor&)
        sSubject$ = MailSubjectNew(SearchFor&)
        If sSender$ = "" Then
            Server.Status.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To Count& - 1
            Server.Status.Caption = "( Now Checking # " & SearchFor& & " Out Of # " & SearchBox& & " )"
            cSender$ = MailSenderNew(SearchBox&)
            cSubject$ = MailSubjectNew(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    Server.Status.Caption = CurCaption$
End Sub

Private Sub itemEnabled_Click()
If itemEnabled.Checked = True Then
    itemEnabled.Checked = False
Else
    itemEnabled.Checked = True
    message.Show
End If
End Sub
Private Sub itemEnd_Click()
Unload Help
End Sub
Private Sub itemexit_Click()
SERVER_FILENAME = App.Path & "\Server.dat"
SERVER_FINDFILE = App.Path & "\ServerFind.dat"
If Server.Label14 = "0" Then Exit Sub
X = MsgBox(AOLUserSN() & " Are You Sure You Want To Exit?", 64 + vbYesNo)
If X = 6 Then
On Error Resume Next
Kill SERVER_FILENAME
Kill SERVER_FINDFILE
Call QuitHelp
If MenuForm.itemSaveExit.Checked Then
If MenuForm.itemNewMail.Checked Then a$ = "0"
If MenuForm.itemOld.Checked Then a$ = "1"
If MenuForm.itemSent.Checked Then a$ = "2"
If MenuForm.itemFlashMail.Checked Then a$ = "3"
Call WriteINI("Preferences", "Mail", a$, App.Path & "\Server.ini")
If MenuForm.itemIm.Checked Then a$ = "1"
If MenuForm.ItemChat.Checked Then a$ = "2"
Call WriteINI("Preferences", "Notify", a$, App.Path & "\Server.ini")
Call WriteINI("Preferences", "IM", a$, App.Path & "\Server.ini")
If MenuForm.itemIdleBot.Checked Then a$ = "1" Else a$ = "0"
Call WriteINI("Preferences", "Idle", a$, App.Path & "\Server.ini")
If MenuForm.ItemWrite.Checked Then a$ = "1"
Call WriteINI("Preferences", "Write", a$, App.Path & "\Server.ini")
If MenuForm.ItemKill.Checked Then a$ = "1"
Call WriteINI("Preferences", "Kill", a$, App.Path & "\Server.ini")
If MenuForm.itemStatus.Checked Then a$ = "1"
Call WriteINI("Preferences", "Status", a$, App.Path & "\Server.ini")
Call WriteINI("Preferences", "Comment", mComm$, App.Path & "\Server.ini")
End If
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
 End
End If
Wel = FindChildByTitle(AOLMDI(), "Welcome,")
Welc$ = String(255, 0)
WhichWel = GetWindowText(Wel, Welc$, 250)
If WhichWel < 8 Then
End
End If
Timeout (1)
 SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "       «–=•(·•· " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "X-Treme Server '99" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ·•·)•=–»·":  DoEvents:
 Timeout 0.01
 SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "        «–=•(·•·  " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & " Now Unloaded * " & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "·•·)•=–»":  DoEvents:
Timeout (1)
End
End If
End Sub

Private Sub itemFlashMail_Click()
itemFlashMail.Checked = True
itemNewMail.Checked = False
itemSent.Checked = False
itemOld.Checked = False
End Sub

Private Sub itemHelp_Click()

Call ShowHelpContents
Server.WindowState = 1
End Sub

Private Sub itemIdleBot_Click()
If itemIdleBot.Checked = True Then
    itemIdleBot.Checked = False
Else
    itemIdleBot.Checked = True
End If
End Sub

Private Sub itemIm_Click()
itemIm.Checked = True
ItemChat.Checked = False
End Sub

Private Sub itemIMServe_Click()
If itemIMServe.Checked = True Then
    itemIMServe.Checked = False
Else
    itemIMServe.Checked = True
End If
End Sub

Private Sub itemInfo_Click()
Info.Show
End Sub

Private Sub itemIntro_Click()
' Declare local variables.
    ';Display.Show
    Dim MsgText
    Dim PB
    ' Initialize the paragraph break variable.
    PB = Chr(10) & Chr(13) '& Chr(10) & Chr(13)
    ' Display the instructions.
    MsgText = "To Start The Server,Make Sure You Create A Mail List To Do This Click On Refresh Mails.I've Added Something Like An Error Box Just in Case You Forget To Create The Mail List" & Chr(13) & Chr(10) & ".Please Allow The Server To Count All Your Mails Wait Until The Status Bar Is On The Ready Command. "
    MsgText = MsgText & Chr(13) & Chr(10) & "Once You Have Created The Mail List You Must To Click On The Start Button To Start Serving,If The Serve Throught IM's Is Not Checked You Have To Be In A Chat Room[Very Important]." & Chr(13) & Chr(10) & "If You Will Like To Pause The Server Just Click On Pause[Hello],To Turn It Back On Click On The Start Button.[Get it?]"
    MsgBox MsgText, 64
End Sub

Private Sub ItemKill_Click()
If ItemKill.Checked = True Then
    ItemKill.Checked = False
Else
    ItemKill.Checked = True
End If
End Sub

Private Sub itemLog_Click()
If itemLog.Checked = True Then
    itemLog.Checked = False
Else
    itemLog.Checked = True
    MsgBox "All Actions Will Be Logged To C:\Server.Log.", 64
End If

End Sub

Private Sub itemMChat_Click()
Chat.Show
End Sub

Private Sub itemmini_Click()
Server.WindowState = 1
End Sub

Private Sub ItemName0_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName0.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName0.Caption)
End Sub

Private Sub ItemName1_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName1.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName1.Caption)
End Sub

Private Sub ItemName10_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName10.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName10.Caption)
End Sub

Private Sub ItemName11_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName11.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName11.Caption)
End Sub

Private Sub ItemName12_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName12.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName12.Caption)
End Sub

Private Sub ItemName13_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName13.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName13.Caption)
End Sub

Private Sub ItemName14_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName14.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName14.Caption)
End Sub

Private Sub ItemName15_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName15.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName15.Caption)
End Sub

Private Sub ItemName16_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName16.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName16.Caption)
End Sub

Private Sub ItemName17_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName17.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName17.Caption)
End Sub

Private Sub ItemName18_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName18.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName18.Caption)
End Sub

Private Sub ItemName19_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName19.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName19.Caption)
End Sub

Private Sub ItemName2_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName2.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName2.Caption)
End Sub

Private Sub ItemName20_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName20.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName20.Caption)
End Sub

Private Sub ItemName21_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName21.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName21.Caption)
End Sub

Private Sub ItemName22_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName22.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName22.Caption)
End Sub

Private Sub ItemName23_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName23.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName23.Caption)
End Sub

Private Sub ItemName3_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName3.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName3.Caption)
End Sub

Private Sub ItemName4_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName4.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName4.Caption)
End Sub

Private Sub ItemName5_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName5.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName5.Caption)
End Sub

Private Sub ItemName6_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName6.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName6.Caption)
End Sub

Private Sub ItemName7_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName7.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName7.Caption)
End Sub

Private Sub ItemName8_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName8.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName8.Caption)
End Sub
Private Sub ItemName9_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & ItemName9.Caption & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now Banned From The Server!": DoEvents
Server.List6.AddItem (ItemName9.Caption)
End Sub
Private Sub itemNewMail_Click()
itemFlashMail.Checked = False
itemNewMail.Checked = True
itemSent.Checked = False
itemOld.Checked = False
End Sub

Private Sub itemOff_Click()
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
  MsgBox "America Online Is Not Loaded", 16
  Exit Sub
End If
  Wel = FindChildByTitle(AOLMDI(), "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
  MsgBox "You Need To Sign On First To Use The Server", 16
  Exit Sub
 End If
 Call InstantMessage("$IM_Off", App.Title)
 
End Sub

Private Sub itemOld_Click()
itemFlashMail.Checked = False
itemNewMail.Checked = False
itemSent.Checked = False
itemOld.Checked = True
End Sub

Private Sub itemon_Click()
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
  MsgBox "America Online Is Not Loaded", 16
  Exit Sub
End If
  Wel = FindChildByTitle(AOLMDI(), "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
  MsgBox "You Need To Sign On First To Use The Server", 16
  Exit Sub
 End If
 Call InstantMessage("$IM_On", App.Title)
 End Sub

Private Sub itemonly_Click()
itemIm.Checked = False
ItemChat.Checked = False
itemonly.Checked = True
End Sub

Private Sub itemRequest_Click()
If ServerBot = True Then MsgBox "Turn Off The Server To Use The Request Bot", 16: Exit Sub
Request.Show
End Sub

Private Sub itemRemove_Click()
If itemRemove.Checked = True Then
    itemRemove.Checked = False
Else
    itemRemove.Checked = True
End If
End Sub

Private Sub itemsave_Click()
If MenuForm.itemNewMail.Checked Then a$ = "0"
If MenuForm.itemOld.Checked Then a$ = "1"
If MenuForm.itemSent.Checked Then a$ = "2"
If MenuForm.itemFlashMail.Checked Then a$ = "3"
Call WriteINI("Preferences", "Mail", a$, App.Path & "\Server.ini")
If MenuForm.itemIm.Checked Then a$ = "1"
If MenuForm.ItemChat.Checked Then a$ = "2"
Call WriteINI("Preferences", "Notify", a$, App.Path & "\Server.ini")
Call WriteINI("Preferences", "IM", a$, App.Path & "\Server.ini")
If MenuForm.itemIdleBot.Checked Then a$ = "1" Else a$ = "0"
Call WriteINI("Preferences", "Idle", a$, App.Path & "\Server.ini")
If MenuForm.ItemWrite.Checked Then a$ = "1"
Call WriteINI("Preferences", "Write", a$, App.Path & "\Server.ini")
If MenuForm.ItemKill.Checked Then a$ = "1"
Call WriteINI("Preferences", "Kill", a$, App.Path & "\Server.ini")
If MenuForm.itemStatus.Checked Then a$ = "1"
Call WriteINI("Preferences", "Status", a$, App.Path & "\Server.ini")
Call WriteINI("Preferences", "Comment", mComm$, App.Path & "\Server.ini")
If MenuForm.itemRemove.Checked Then a$ = "1"
Call WriteINI("Preferences", "Remove", a$, App.Path & "\Server.ini")
End Sub

Private Sub itemSaveExit_Click()
If itemSaveExit.Checked = True Then
    itemSaveExit.Checked = False
Else
    itemSaveExit.Checked = True
End If
End Sub

Private Sub itemScroll_Click()
Room = AOLFindChatRoom
If Room = 0 Then
  MsgBox "You Must Be In a Chat Room!", 16
Exit Sub
End If
If ServerBot = False Then MsgBox "Server is Not On!", 16: Exit Sub

SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">/" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send List - ( " & Trim(Str$(Server.List2.ListCount)) & " ) Mails":  DoEvents
Timeout 0.1
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">/" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send X - X Is Index":  DoEvents:
Timeout 0.1
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">/" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Find X - X Is Search Query":  DoEvents
End Sub

Private Sub itemSend_Click()


End Sub

Private Sub itemSent_Click()
itemFlashMail.Checked = False
itemNewMail.Checked = False
itemSent.Checked = True
itemOld.Checked = False
End Sub

Private Sub itemStatus_Click()
If itemStatus.Checked = True Then
    itemStatus.Checked = False
Else
    itemStatus.Checked = True
End If
End Sub

Private Sub itemSystem_Click()

End Sub

Private Sub itemTest_Click()

If IFileExists(App.Path & "\Server.ini") Then
 On Error Resume Next
 Open App.Path & "\Server.ini" For Binary As #1
 For i& = 1 To LOF(1)
 Get #1, i&, n%
 Done% = i&
 Call PercentBar(Server.Picture1, Done%, LOF(1))
 Next i&
 Close #1
 If Err Then MsgBox (Error(Err)): Exit Sub
 MsgBox "No Problems Were Found!.", 64: Server.Picture1.BackColor = QBColor(0): Exit Sub
 Else
 MsgBox "An Error Was Found On The Server Ini File" & Chr(13) & Chr(10) & "X-Treme Server Will Try To Fix This Error" & Chr(13) & Chr(10) & "Click OK To Continue Thank You.", 16
 If MenuForm.itemNewMail.Checked Then a$ = "0"
 If MenuForm.itemOld.Checked Then a$ = "1"
 If MenuForm.itemSent.Checked Then a$ = "2"
 If MenuForm.itemFlashMail.Checked Then a$ = "3"
 Call WriteINI("Preferences", "Mail", a$, App.Path & "\Server.ini")
 If MenuForm.itemIm.Checked Then a$ = "1"
 If MenuForm.ItemChat.Checked Then a$ = "2"
 Call WriteINI("Preferences", "Notify", a$, App.Path & "\Server.ini")
 Call WriteINI("Preferences", "IM", a$, App.Path & "\Server.ini")
 If MenuForm.itemIdleBot.Checked Then a$ = "1" Else a$ = "0"
 Call WriteINI("Preferences", "Idle", a$, App.Path & "\Server.ini")
 If MenuForm.ItemWrite.Checked Then a$ = "1"
 Call WriteINI("Preferences", "Write", a$, App.Path & "\Server.ini")
 If MenuForm.ItemKill.Checked Then a$ = "1"
 Call WriteINI("Preferences", "Kill", a$, App.Path & "\Server.ini")
 If MenuForm.itemStatus.Checked Then a$ = "1"
 Call WriteINI("Preferences", "Status", a$, App.Path & "\Server.ini")
 Call WriteINI("Preferences", "Comment", mComm$, App.Path & "\Server.ini")
 MsgBox "Problem Has Been Fixed!!", 64
 Exit Sub
 End If
End Sub

Private Sub itemtotal_Click()
a$ = ReadINI("Since", "Sent", App.Path & "\Server.ini")
MsgBox "X-Treme Server Served ( " & a$ & " ) Mails Since Installed", 64

End Sub

Private Sub itemUser_Click()
If itemUser.Checked = True Then
    itemUser.Checked = False
Else
    itemUser.Checked = True
End If
End Sub

Private Sub itemTrouble_Click()

End Sub

Private Sub itemUnWanted_Click()

End Sub


Private Sub itemview_Click()
a$ = ReadINI("Server", "Total", App.Path & "\server.ini")
MsgBox "X-Treme Server '99 Have Sent " & a$ & " Mails Since Installed.", 64
End Sub

Private Sub itemwait_Click()
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
    MsgBox "America Online Is Not Open.", 16
    Exit Sub
End If
Wel = FindChildByTitle(AOLMDI(), "Welcome,")
Welc$ = String(255, 0)
WhichWel = GetWindowText(Wel, Welc$, 250)
If WhichWel < 8 Then
    MsgBox "You Need To Sign On First To Use The Server.", 16
    Exit Sub
End If
Call KillWait

End Sub

Private Sub itemWelcome_Click()


End Sub

Private Sub ItemWrite_Click()
If ItemWrite.Checked = True Then
    ItemWrite.Checked = False
Else
    ItemWrite.Checked = True
End If
End Sub

Private Sub ItmFeatures_Click()
Features.Show
End Sub

