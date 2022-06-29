VERSION 5.00
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SYSTRAY.OCX"
Begin VB.Form Form10 
   ClientHeight    =   2445
   ClientLeft      =   2595
   ClientTop       =   -4335
   ClientWidth     =   5115
   Icon            =   "KnK-menue.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5115
   Begin VB.OptionButton opt_FormIcon 
      Caption         =   "Use Form Icon"
      Height          =   375
      Left            =   80000
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin SysTray.SystemTray SystemTray1 
      Left            =   3000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   "KnK Founders Server Helper"
      IconFile        =   0
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu opttt 
         Caption         =   "Options"
      End
      Begin VB.Menu lineeeeeeeeeee 
         Caption         =   "-"
      End
      Begin VB.Menu ver 
         Caption         =   "Version"
      End
      Begin VB.Menu greetz 
         Caption         =   "Greetz"
      End
      Begin VB.Menu note 
         Caption         =   "Notes"
      End
      Begin VB.Menu PooK 
         Caption         =   "Dedicated too..."
      End
      Begin VB.Menu help2 
         Caption         =   "Help"
      End
      Begin VB.Menu mail 
         Caption         =   "E-MAIL"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu web 
         Caption         =   "Web Sites"
         Begin VB.Menu knk 
            Caption         =   "KnK VB Site"
         End
         Begin VB.Menu knkgn 
            Caption         =   "KnK Generation Next"
         End
         Begin VB.Menu knksh 
            Caption         =   "KnK Server Helper"
         End
         Begin VB.Menu pp 
            Caption         =   "PooKs Place"
         End
         Begin VB.Menu pw 
            Caption         =   "Proggie Warehouse"
         End
         Begin VB.Menu EnD 
            Caption         =   "EnD's WeB SiTE!!!!"
         End
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu min 
         Caption         =   "Minimize"
      End
      Begin VB.Menu exit2 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Toolz 
      Caption         =   "Toolz"
      Begin VB.Menu find 
         Caption         =   "Find Bin"
         Begin VB.Menu sm 
            Caption         =   "Setup/Modify"
         End
         Begin VB.Menu staryf 
            Caption         =   "Start"
         End
      End
      Begin VB.Menu ghdfhsd 
         Caption         =   "-"
      End
      Begin VB.Menu im 
         Caption         =   "IM's"
         Begin VB.Menu off 
            Caption         =   "IM's Off"
         End
         Begin VB.Menu on 
            Caption         =   "IM's On"
         End
      End
      Begin VB.Menu adver 
         Caption         =   "Advertisements"
         Begin VB.Menu adv1 
            Caption         =   "Advertise 1"
         End
         Begin VB.Menu adv2 
            Caption         =   "Advertise 2"
         End
         Begin VB.Menu ad3 
            Caption         =   "Advertise 3"
         End
      End
      Begin VB.Menu bust 
         Caption         =   "Room Buster"
      End
      Begin VB.Menu bots 
         Caption         =   "Bots"
         Begin VB.Menu afk 
            Caption         =   "AFK Bot"
         End
         Begin VB.Menu attention 
            Caption         =   "Attention Bot"
         End
         Begin VB.Menu idle 
            Caption         =   "Idle Bot"
         End
         Begin VB.Menu aib 
            Caption         =   "Anti Idle Bot"
         End
      End
      Begin VB.Menu tlz 
         Caption         =   "Toolz"
         Begin VB.Menu upcht 
            Caption         =   "UpnChat"
         End
         Begin VB.Menu unchat 
            Caption         =   "UnUpnChat"
         End
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu aoh 
         Caption         =   "Add Odd Handle"
      End
      Begin VB.Menu st 
         Caption         =   "Server Type"
         Begin VB.Menu atst 
            Caption         =   "About The Server Types"
         End
         Begin VB.Menu linj 
            Caption         =   "-"
         End
         Begin VB.Menu ais 
            Caption         =   "AO-NiN IM Style"
         End
         Begin VB.Menu ncs 
            Caption         =   "Normal Chatroom Server"
            Checked         =   -1  'True
         End
         Begin VB.Menu es 
            Caption         =   "èmbràcè sèrvèr"
         End
         Begin VB.Menu vms 
            Caption         =   "válkyrie máil server"
         End
      End
      Begin VB.Menu cdclar 
         Caption         =   "Clear/dont clear list after requsting"
         Begin VB.Menu clar 
            Caption         =   "Clear List after Requesting"
         End
         Begin VB.Menu dclar 
            Caption         =   "Dont Clear List after Requesting"
         End
      End
      Begin VB.Menu pbr 
         Caption         =   "Pauses Between Requests"
         Begin VB.Menu sec1 
            Caption         =   "1 sec"
         End
         Begin VB.Menu sec2 
            Caption         =   "2 sec"
         End
         Begin VB.Menu sec3 
            Caption         =   "3 sec"
         End
         Begin VB.Menu sec4 
            Caption         =   "4 sec"
         End
      End
   End
   Begin VB.Menu GetName 
      Caption         =   "Get Name"
   End
   Begin VB.Menu help 
      Caption         =   "help"
      Begin VB.Menu intro 
         Caption         =   "Intro"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
      Begin VB.Menu ascii 
         Caption         =   "Ascii"
      End
      Begin VB.Menu theme 
         Caption         =   "Theme"
      End
      Begin VB.Menu scroll 
         Caption         =   "Scroll"
      End
      Begin VB.Menu aol 
         Caption         =   "AOL"
      End
   End
   Begin VB.Menu ascii2 
      Caption         =   "ascii2"
      Begin VB.Menu bbb 
         Caption         =   "bbb"
      End
      Begin VB.Menu bgb 
         Caption         =   "bgb"
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Constants for the ACTION PROPERTY
Private Const sys_Add = 0       'Specifies that an icon is being add
Private Const sys_Modify = 1    'Specifies that an icon is being modified
Private Const sys_Delete = 2    'Specifies that an icon is being deleted

'Constants for the ERROR EVENT
Private Const errUnableToAddIcon = 1    'Icon can not be added to system tray
Private Const errUnableToModifyIcon = 2 'System tray icon can not be modified
Private Const errUnableToDeleteIcon = 3 'System tray icon can not be deleted
Private Const errUnableToLoadIcon = 4   'Icon could not be loaded (occurs while using icon property)

'Constants for MOUSE RELATETED EVENTS
Private Const vbLeftButton = 1     'Left button is pressed
Private Const vbRightButton = 2    'Right button is pressed
Private Const vbMiddleButton = 4   'Middle button is pressed

Private Const cdlOpen = 1
Private Const cdlSave = 2
Private Const cdlColor = 3
Private Const cdlPrint = 4

Private Const cdlOFNReadOnly = 1             'Checks Read-Only check box for Open and Save As dialog boxes.
Private Const cdlOFNOverwritePrompt = 2      'Causes the Save As dialog box to generate a message box if the selected file already exists.
Private Const cdlOFNHideReadOnly = 4         'Hides the Read-Only check box.
Private Const cdlOFNNoChangeDir = 8          'Sets the current directory to what it was when the dialog box was invoked.
Private Const cdlOFNHelpButton = 10          'Causes the dialog box to display the Help button.
Private Const cdlOFNNoValidate = 100         'Allows invalid characters in the returned filename.
Private Const cdlOFNAllowMultiselect = 200   'Allows the File Name list box to have multiple selections.
Private Const cdlOFNExtensionDifferent = 400 'The extension of the returned filename is different from the extension set by the DefaultExt property.
Private Const cdlOFNPathMustExist = 800      'User can enter only valid path names.
Private Const cdlOFNFileMustExist = 1000     'User can enter only names of existing files.
Private Const cdlOFNCreatePrompt = 2000      'Sets the dialog box to ask if the user wants to create a file that doesn't currently exist.
Private Const cdlOFNShareAware = 4000        'Sharing violation errors will be ignored.
Private Const cdlOFNNoReadOnlyReturn = 8000  'The returned file doesn't have the Read-Only attribute set and won't be in a write-protected directory.
Private Const cdlOFNExplorer = 8000          'Use the Explorer-like Open A File dialog box template.  (Windows 95 only.)
Private Const cdlOFNNoDereferenceLinks = 100000 'Do not dereference shortcuts (shell links).  By default, choosing a shortcut causes it to be dereferenced by the shell.  (Windows 95 only.)
Private Const cdlOFNLongNames = 200000       'Use Long filenames.  (Windows 95 only.)

Private Sub ad3_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`° Making it easy °´¯`×-»")
TimeOut (0.9)
AOLChatSend ("«-×´¯`° To get your shit °´¯`×-»")
End If

'AOL 4.o command
If aversion$ = "aol4" Then
If UserAOL = "AOL4" Then
Dim NumSel As Integer
NumSel = Random(2)
If NumSel = 1 Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`° Making it easy °´¯`×-»")
TimeOut (0.9)
SendChat BlackGreenBlack("«-×´¯`° To get your shit °´¯`×-»")
ElseIf NumSel = 2 Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`° Making it easy °´¯`×-»")
TimeOut (0.9)
SendChat BlackBlueBlack("«-×´¯`° To get your shit °´¯`×-»")
End If
End If
End If
End Sub

Private Sub adt_Click()


End Sub

Private Sub adv1_Click()
'AOL4.o command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol4" Then
Dim NumSel As Integer
NumSel = Random(2)
If NumSel = 1 Then
 SendChat BlackGreenBlack("«-×´¯`°   KnK Founders  °´¯`×-»")
TimeOut (0.6)
SendChat BlackGreenBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.6)
SendChat BlackGreenBlack("«-×´¯`°  Version: 2.o     °´¯`×-»")

ElseIf NumSel = 2 Then
SendChat BlackBlueBlack("«-×´¯`°   KnK Founders  °´¯`×-»")
TimeOut (0.6)
SendChat BlackBlueBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.6)
SendChat BlackBlueBlack("«-×´¯`°  Version: 2.o     °´¯`×-»")
End If
End If

'AOL95 command
If aversion$ = "aol95" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Version: 2.o     °´¯`×-»")
End If
End Sub

Private Sub adv2_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Get yours at   °´¯`×-»")
TimeOut (0.9)
AOLChatSend ("«-×´¯`° http://knk.tierranet.com/serv")
End If

'AOL4.o command
If aversion$ = "aol4" Then
Dim NumSel As Integer
NumSel = Random(2)
If NumSel = 1 Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Get yours at   °´¯`×-»")
TimeOut (0.9)
SendChat BlackGreenBlack("«-×´¯`° http://knk.tierranet.com/serv")
ElseIf NumSel = 2 Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Get yours at   °´¯`×-»")
TimeOut (0.9)
SendChat BlackBlueBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If
End If
End Sub

Private Sub afk_Click()
Form9.Show
End Sub

Private Sub aib_Click()

Call AntiIdle
End Sub

Private Sub ais_Click()
Form7.Label12.Caption = "im"
Form7.Option3.Enabled = True
Form13.Label12.Caption = "im"
Form13.Option3.Enabled = True

vms.Checked = False
ais.Checked = True
ncs.Checked = False
es.Checked = False
End Sub

Private Sub aoh_Click()
Form3.Show
End Sub

Private Sub aol_Click()
MsgBox "The Program automaticly detects your AOL on startup,  but if you switch AOLs with this prog still operating you must change the AOL in the options section.", vbInformation, "AOL"
End Sub

Private Sub ascii_Click()
MsgBox "This option is only for AOL 4.o.  This option alows you to choose your ascii color.  either BlackBlueBlack, or BlackGreenBlack.", vbQuestion, "Ascii"
End Sub

Private Sub atst_Click()
Form6.Show
End Sub

Private Sub attention_Click()
Form5.Show
End Sub

Private Sub bbb_Click()
bbb.Checked = True
bgb.Checked = False

End Sub

Private Sub bgb_Click()
bbb.Checked = False
bgb.Checked = True

End Sub

Private Sub bust_Click()
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
On Local Error Resume Next
Call ShellExecute(hwnd, "Open", "Fatebust.EXE", "", App.Path, 1)
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'Fatebust.EXE'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If

End If
If aversion$ = "aol4" Then
On Local Error Resume Next
Call ShellExecute(hwnd, "Open", "CASABLANCA.exe", "", App.Path, 1)
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'CASABLANCA.exe'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End If
End Sub

Private Sub clar_Click()
clar.Checked = True
dclar.Checked = False
On Local Error Resume Next
R% = WritePrivateProfileString("List", "clearornot", "Clear", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub cmd_browse_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub dclar_Click()
clar.Checked = False
dclar.Checked = True

On Local Error Resume Next
R% = WritePrivateProfileString("List", "clearornot", "dontclear", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub EnD_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
MsgBox "Check out the best Artist ever,  He is awsome!!!", vbInformation, "This site Rules"
AOLKeyword ("http://knk.tierranet.com/end-inc")
End If

'AOL4.o command
If aversion$ = "aol4" Then
MsgBox "Check out the best Artist ever,  He is awsome!!!", vbInformation, "This site Rules"
KeyWord ("http://knk.tierranet.com/end-inc")
End If
End Sub

Private Sub es_Click()
Form7.Label12.Caption = "embrace"
Form7.Option3.Enabled = False
Form13.Label12.Caption = "embrace"
Form13.Option3.Enabled = False
vms.Checked = False
ais.Checked = False
ncs.Checked = False
es.Checked = True

End Sub

Private Sub exit_Click()
MsgBox "You can chose if you wanna see the skull exit theme,  if you dont,  when you click exit,  the prog will just end", vbQuestion, "Exit"
End Sub



Private Sub exit2_Click()
Loads2$ = GetFromINI("Exit", "Loads2", App.Path + "\KnK.ini")
If Loads2$ = "no" Then
End
End If
If Loads2$ = "yes" Then
Unload Form7
Unload Form13
Form11.Show
End If
End Sub

Private Sub Form_Load()
SystemTray1.icon = Val(Form10.icon)
If UserAOL = "AOL95" Then
ascii.Visible = False
aib.Visible = False

End If

On Local Error Resume Next
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If

If Clearornot$ = "Clear" Then
clar.Checked = True
End If

If Clearornot = "dontclear" Then
dclar.Checked = True
End If

On Local Error Resume Next
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If

If TimeKnK$ = "1" Then
sec1.Checked = True
End If

If TimeKnK$ = "2" Then
sec2.Checked = True
End If

If TimeKnK$ = "3" Then
sec3.Checked = True
End If

If TimeKnK$ = "4" Then
sec4.Checked = True
End If
''''''''''''''''''''''''''''''''''''''''''''''''
'Ascii ini option
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
If Color$ = "bbb" Then
bbb.Checked = True
End If
If Color$ = "bgb" Then
bgb.Checked = True
End If

End Sub

Private Sub greetz_Click()
Form4.Show
End Sub



Private Sub help2_Click()
'Knkserv
On Local Error Resume Next
Call ShellExecute(hwnd, "Open", "Knkserv.hlp", "", App.Path, 1)
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'Knkserv.hlp'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If

End Sub

Private Sub idle_Click()
Form8.Show
End Sub

Private Sub intro_Click()
MsgBox "If you dont like or dont want to waite for the intro,  click no.  If you like it like me,  click yes.", vbQuestion, "Intro"
End Sub

Private Sub knk_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
AOLKeyword ("http://knk.tierranet.com/knk")
End If

'AOL4.o command
If aversion$ = "aol4" Then
KeyWord ("http://knk.tierranet.com/knk")
End If
End Sub

Private Sub knkgn_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
AOLKeyword ("http://knk.tierranet.com/knk4o")
End If

'AOL4.o command
If aversion$ = "aol4" Then
KeyWord ("http://knk.tierranet.com/knk4o")
End If
End Sub

Private Sub knksh_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
AOLKeyword ("http://knk.tierranet.com/serv")
End If

'AOL4.o command
If aversion$ = "aol4" Then
KeyWord ("http://knk.tierranet.com/serv")
End If
End Sub

Private Sub kv_Click()
Form15.Show
End Sub

Private Sub List1_Click()

End Sub

Private Sub mail_Click()

Form2.Show

End Sub

Private Sub min_Click()
Unload Form7
Unload Form13
SystemTray1.Action = sys_Add
End Sub

Private Sub ncs_Click()
Form7.Label12.Caption = "chat"
Form7.Option3.Enabled = True
Form13.Label12.Caption = "chat"
Form13.Option3.Enabled = True

vms.Checked = False
ais.Checked = False
ncs.Checked = True
es.Checked = False
End Sub

Private Sub note_Click()
MsgBox "This program does not have a ton of useless crap like other progs.  As an ex..  WTF do you need a SN decoder for?  You copy and past the name in the box to decode,  why not just past it ware u need it.  This prog also doesnt have a ton of scrollers,  no macro/macro killers or hardly any bots, and NO Punters or TOSers.  There not needed today.", vbInformation, "Notes"
End Sub

Private Sub off_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
Call AOLInstantMessage("$IM_OFF", "KnK 4 Life")
End If

'AOL4.o command
If aversion$ = "aol4" Then
Call IMKeyword("$IM_OFF", " KnK 4 Life")
End If
End Sub

Private Sub on_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
Call AOLInstantMessage("$IM_ON", "KnK 4 Life")
End If

'AOL 4.o command
If aversion$ = "aol4" Then
Call IMKeyword("$IM_ON", " KnK 4 Life")
End If
End Sub

Private Sub opt_Click()

End Sub

Private Sub opt_FormIcon_Click()
SystemTray1.icon = Val(Form10.icon)


End Sub

Private Sub Option1_Click()
SystemTray1.icon = Val(Form10.icon)


End Sub

Private Sub opttt_Click()
Form12.Show
End Sub

Private Sub PooK_Click()
MsgBox "This program is dedicated to PooK.  He has been such a great friend, Hes one of the best programers on AOL.  Thanx for helping me so much  and having such a positive attitude.", vbInformation, "Dedicated too........."
End Sub

Private Sub pp_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
AOLKeyword ("http://knk.tierranet.com/PooK")
MsgBox "This is Ma HoMyS Site, Come here often!!!", vbInformation, "PooKs Place"
End If

'AOL4.o command
If aversion$ = "aol4" Then
KeyWord ("http://knk.tierranet.com/PooK")
MsgBox "This is Ma HoMyS Site, Come here often!!!", vbInformation, "PooKs Place"

End If
End Sub

Private Sub pw_Click()
'AOL95 command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
MsgBox "Here you can get a ton of Progs!?", vbInformation, "ToN O Progs"
AOLKeyword ("http://ryan.tierranet.com/progy")
End If

'AOL4.o command
If aversion$ = "aol4" Then
KeyWord ("http://ryan.tierranet.com/progy")
End If
End Sub

Private Sub rr_Click()

End Sub

Private Sub scroll_Click()
MsgBox "Well some people don't like a program to scroll.  Well you can decide here.  Please, I ask you not to click no.  I need the exposure.", vbInformation, "Scroll"
End Sub

Private Sub sec1_Click()
MsgBox "This is not recomended because:" & vbCrLf & "<1> Ao-NiN and most others recomend at least 2 seconds between requests.", vbInformation, "Not recomended"
sec1.Checked = True
sec2.Checked = False
sec3.Checked = False
sec4.Checked = False

On Local Error Resume Next
R% = WritePrivateProfileString("Pauses", "TimeKnK", "1", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub sec2_Click()
sec1.Checked = False
sec2.Checked = True
sec3.Checked = False
sec4.Checked = False

On Local Error Resume Next
R% = WritePrivateProfileString("Pauses", "TimeKnK", "2", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub sec3_Click()
sec1.Checked = False
sec2.Checked = False
sec3.Checked = True
sec4.Checked = False

On Local Error Resume Next
R% = WritePrivateProfileString("Pauses", "TimeKnK", "3", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub sec4_Click()
sec1.Checked = False
sec2.Checked = False
sec3.Checked = False
sec4.Checked = True

On Local Error Resume Next
R% = WritePrivateProfileString("Pauses", "TimeKnK", "4", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub sm_Click()
Form14.Show
End Sub

Private Sub staryf_Click()

Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
KnKload$ = GetFromINI("KnKTheme", "KnKload", App.Path + "\KnK.ini")
If KnKload$ = "knk1" Then

If aversion$ = "aol4" Then
If adver$ = "yes" Then
'Black Blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°       Activated     °´¯`×-»")
TimeOut (2)
End If

'Black Green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°       Activated     °´¯`×-»")
TimeOut (2)
End If
End If
If adver$ = "no" Then
End If
Form7.Timer18.Enabled = True
End If

'AOL95 command
If aversion$ = "aol95" Then
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°       Activated     °´¯`×-»")
TimeOut (2)
End If
If adver$ = "no" Then
End If
Form7.Timer19.Enabled = True
End If
End If
If KnKload$ = "knk2" Then

If aversion$ = "aol4" Then
If adver$ = "yes" Then
'Black Blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°       Activated     °´¯`×-»")
TimeOut (2)
End If

'Black Green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°       Activated     °´¯`×-»")
TimeOut (2)
End If
End If
If adver$ = "no" Then
End If
Form13.Timer18.Enabled = True
End If

'AOL95 command
If aversion$ = "aol95" Then
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°       Activated     °´¯`×-»")
TimeOut (2)
End If
If adver$ = "no" Then
End If
Form13.Timer19.Enabled = True
End If
End If
End Sub






Private Sub SystemTray1_Error(ByVal ErrorNumber As Integer)
Beep

End Sub

Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)
SystemTray1.Action = sys_Delete
KnKload$ = GetFromINI("KnKTheme", "KnKload", App.Path + "\KnK.ini")
If KnKload$ = "knk1" Then
Form7.Show
End If
If KnKload$ = "knk2" Then
Form13.Show
End If
End Sub

Private Sub theme_Click()
MsgBox "This alows you to choose between two color themes.  This only aplys on the server helper its self.  If you decide to change the theme, you must either minimize or exit the prog so the changes take effect.", vbQuestion, "Theme"
End Sub

Private Sub unchat_Click()
Call UnUpchat
End Sub

Private Sub upcht_Click()

Call Upchat

End Sub

Private Sub ver_Click()
frmAbout.Show
End Sub

Private Sub vms_Click()
Form7.Label12.Caption = "valkyrie"
Form7.Option3.Enabled = True
Form13.Label12.Caption = "valkyrie"
Form13.Option3.Enabled = True
es.Checked = False
vms.Checked = True
ais.Checked = False
ncs.Checked = False
End Sub

Private Sub wn_Click()


End Sub

Private Sub wrb_Click()
Form16.Show
End Sub
