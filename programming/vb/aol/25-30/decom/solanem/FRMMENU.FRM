VERSION 2.00
Begin Form frmMenu 
   BackColor       =   &H8000000F&
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   2055
   ClientTop       =   2490
   ClientWidth     =   2070
   Height          =   1425
   Left            =   1995
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   2070
   Top             =   2145
   Visible         =   0   'False
   Width           =   2190
   Begin Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1065
      Top             =   480
   End
   Begin VBMsg VBMsg2 
      Height          =   420
      Left            =   405
      MessageCount    =   FRMMENU.FRX:0000
      MessageList     =   FRMMENU.FRX:0002
      MessageTypes    =   0  'Selected Messages
      PostDefault     =   0   'False
      Top             =   405
      Width           =   420
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1065
      Top             =   30
   End
   Begin VBMsg VBMsg1 
      Height          =   420
      Left            =   390
      MessageCount    =   FRMMENU.FRX:0200
      MessageList     =   FRMMENU.FRX:0202
      MessageTypes    =   0  'Selected Messages
      PostDefault     =   0   'False
      Top             =   30
      Width           =   420
   End
   Begin CheckBox chkIM 
      Caption         =   "Check1"
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin CheckBox chkTool 
      Caption         =   "Check1"
      Height          =   495
      Left            =   75
      TabIndex        =   1
      Top             =   390
      Value           =   1  'Checked
      Width           =   1215
   End
End

Sub Form_Load ()
'*** Code added by HelpWriter ***
    SetApphelp Me.hWnd
'***********************************
'first find the handle of AOL
aol% = FindWindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
If aol% = 0 Then MsgBox "America Online must be loaded first", 32: End
If app.PrevInstance = True Then MsgBox "Only one SoLuTioN can be loaded at a time", 32: End
        If IFileExists(app.Path & "\PopUps.INI") = -1 Then
        Open app.Path & "\PopUps.INI" For Input As #1
            Do While Not EOF(1)
             Line Input #1, pops$
                If InStr(UCase(pops$), UCase("PopUp0=")) Then PopUp0$ = Mid$(pops$, 8)
                If InStr(pops$, "PopUp1=") Then PopUp1$ = Mid$(pops$, 8)
                If InStr(pops$, "PopUp2=") Then PopUp2$ = Mid$(pops$, 8)
            Loop
          Close #1
        End If

'then the handle of AOL's menu
aolmenu% = getmenu(aol%)
'then we create 3 popup menus, one will be the main one added to AOL, the other two will be submenus of the main menu we add
mainmenu = CreatePopupMenu()
submenu1% = CreatePopupMenu()
submenu2% = CreatePopupMenu()
'this adds 3 menu items to the first submenu
X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 1, "&Room Buster")
X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 2, "&Mass Mailer")
X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 11, "&SiGN oN MaGiK!")
X% = AppendMenu(mainmenu, MF_String Or MF_ENABLED, 12, "M&ail Tools")
X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 3, "&Instant Messages Off")
X% = AppendMenu(mainmenu, MF_ENABLED Or MF_CHECKED, 10, "&AOL Toolbar")
'this adds 3 menu items to the second submenu
If PopUp0$ <> "" Then X% = AppendMenu(mainmenu, MF_Separator, 0&, 0&)
If PopUp0$ <> "" Then X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 4, PopUp0$)
If PopUp1$ <> "" Then X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 5, PopUp1$)
If PopUp2$ <> "" Then X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 6, PopUp2$)
'X% = AppendMenu(mainmenu, MF_String Or MF_Popup, submenu2%, "Sub Menu &2")
'this adds a seperator onto our mainmenu
X% = AppendMenu(mainmenu, MF_Separator, 0&, 0&)
'this adds there menu items to our mainmenu
'X% = AppendMenu(mainmenu, MF_Enabled Or MF_String, 7, "Menu Item &7")
'X% = AppendMenu(mainmenu, MF_Enabled Or MF_String, 8, "Menu Item &8")
X% = AppendMenu(mainmenu, MF_ENABLED Or MF_String, 9, "E&xit")
'this adds our main menu to aol's menu bar
X% = AppendMenu(aolmenu%, MF_String Or MF_Popup, mainmenu, "&SoLuTioN")
'this redraws aol's menu bar
Drawmenubar (aol%)
vbmsg1.SubClasshWnd = aol%
vbmsg2.SubClasshWnd = mdi%
Call ChatSay("SoLuTioN HoST", "SoLuTioN Mailing System has been loaded"): DoEvents
'this sets VBMsg to subclass AOL, allowing us to make our menu items do something
Rem MsgBox "The menu has been added"
End Sub

Sub Form_Unload (Cancel As Integer)
'*** Code added by HelpWriter ***
    QuitHelp
'***********************************
Call RMVMenu("&SoLuTioN")
End Sub

Sub Timer1_Timer ()
Select Case MNUClk
    Case 1 'buster
        frmBuster.Show
        timer1.Enabled = False
    
    Case 2 'mmer
        frmMMer.Show
        timer1.Enabled = False

    Case 11 'signon
        frmSignOn.Show
        timer1.Enabled = False

    Case 12 'mailtools
        frmMailTools.Show
        timer1.Enabled = False

    Case 3 'im`z
        If chkIM.Value = 1 Then
            Call TurnIM("Off"): DoEvents
            chkIM.Value = 0
            X = modifymenu((getmenu(FindWindow("AOL Frame25", 0&))), 3, 0, 3, "&Instant Messages On")
            timer1.Enabled = False
        Else
            Call TurnIM("On"): DoEvents
            chkIM.Value = 1
            X = modifymenu((getmenu(FindWindow("AOL Frame25", 0&))), 3, 0, 3, "&Instant Messages Off")
            timer1.Enabled = False
        End If
    
    Case 10 'im`z
        If chkTool.Value = 1 Then
            X = showwindow((findchildbyclass(FindWindow("AOL Frame25", 0&), "AOL Toolbar")), SW_HIDE): X = showwindow((findchildbyclass(FindWindow("AOL Frame25", 0&), "AOL Toolbar")), SW_HIDE): X = showwindow((FindWindow("AOL Frame25", 0&)), SW_NORMAL): X = showwindow((FindWindow("AOL Frame25", 0&)), SW_MAXIMIZE): DoEvents
            chkTool.Value = 0
            X = modifymenu((getmenu(FindWindow("AOL Frame25", 0&))), 10, MF_UNCHECKED, 10, "&AOL Toolbar")
            timer1.Enabled = False
        Else
            X = showwindow((findchildbyclass(FindWindow("AOL Frame25", 0&), "AOL Toolbar")), SW_SHOW): X = showwindow((FindWindow("AOL Frame25", 0&)), SW_NORMAL): X = showwindow((FindWindow("AOL Frame25", 0&)), SW_MAXIMIZE): DoEvents
            chkTool.Value = 1
            X = modifymenu((getmenu(FindWindow("AOL Frame25", 0&))), 10, MF_CHECKED, 10, "&AOL Toolbar")
            timer1.Enabled = False
        End If

    Case 4 'pop0
        PopBuster0 = True
        frmBuster.cmbRoom.Text = PopUp0$
        frmBuster.Show
        Do: DoEvents
            bst% = frmBuster.hWnd
        Loop Until bst% <> 0
        bst% = frmBuster.hWnd
        Do: DoEvents
        BustIn% = findchildbytitle(bst%, "Bust-In")
        Loop Until BustIn% <> 0
        Do: DoEvents
        rm$ = frmBuster.cmbRoom.Text
        Loop Until Len(rm$) > 0
        Call Click(BustIn%, 1): DoEvents
        timer1.Enabled = False

    Case 5 'pop1
        PopBuster0 = True
        frmBuster.cmbRoom.Text = PopUp1$
        frmBuster.Show
        Do: DoEvents
            bst% = frmBuster.hWnd
        Loop Until bst% <> 0
        bst% = frmBuster.hWnd
        Do: DoEvents
        BustIn% = findchildbytitle(bst%, "Bust-In")
        Loop Until BustIn% <> 0
        Do: DoEvents
        rm$ = frmBuster.cmbRoom.Text
        Loop Until Len(rm$) > 0
        Call Click(BustIn%, 1): DoEvents
        timer1.Enabled = False

    Case 6 'pop2
        PopBuster0 = True
        frmBuster.cmbRoom.Text = PopUp2$
        frmBuster.Show
        Do: DoEvents
            bst% = frmBuster.hWnd
        Loop Until bst% <> 0
        bst% = frmBuster.hWnd
        Do: DoEvents
        BustIn% = findchildbytitle(bst%, "Bust-In")
        Loop Until BustIn% <> 0
        Do: DoEvents
        rm$ = frmBuster.cmbRoom.Text
        Loop Until Len(rm$) > 0
        Call Click(BustIn%, 1): DoEvents
        timer1.Enabled = False

    Case 9 'exit
        Unload Me
        End
        timer1.Enabled = False
End Select
End Sub

Sub Timer2_Timer ()
  aol% = FindWindow("AOL Frame25", 0&)
  Call runmenubystring(aol%, "Edit &Address Book..."): DoEvents
    Do: DoEvents
        modal% = FindWindow("_AOL_MODAL", "Address Book")
    Loop Until modal% <> 0
  okbtn% = findchildbytitle(modal%, "OK")
  Do: DoEvents
    Call Click(okbtn%, 1): DoEvents
        modal% = FindWindow("_AOL_MODAL", "Address Book")
    Loop Until modal% = 0
    timer2.Enabled = False

End Sub

Sub VBMsg1_WindowDestroyed (hWindow As Integer)
End
End Sub

Sub VBMsg1_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
For MNUClk = 0 To 12
    If wParam = MNUClk Then
        timer1.Enabled = True
        Exit Sub
    End If
Next MNUClk
Rem Debug.Print wparam
End Sub

Sub VBMsg2_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
timer2.Enabled = True

End Sub

