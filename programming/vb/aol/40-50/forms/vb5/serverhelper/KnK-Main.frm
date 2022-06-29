VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   960
   ClientTop       =   870
   ClientWidth     =   7575
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "KnK-Main.frx":0000
   ScaleHeight     =   5025
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   0
      Picture         =   "KnK-Main.frx":D964
      ScaleHeight     =   5025
      ScaleWidth      =   7545
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'This finds the AOL version
aol% = FindWindow("AOL Frame25", vbNullString)
If aol% Then
  AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
  AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
  If AOTool2% > 0 Then
    UserAOL = "AOL4"
R% = WritePrivateProfileString("AOL", "aversion", "aol4", App.Path + "\KnK.ini")
  Else
    UserAOL = "AOL95"
  R% = WritePrivateProfileString("AOL", "aversion", "aol95", App.Path + "\KnK.ini")

  End If
End If
If UserAOL = "" Then
MsgBox "KnK Founders Server helper could not determine your AOL version.  Either you are not using AOL95 or AOL4.o,  AOL is not open, or your AOL may be minimized.", vbExclamation, "Error"
End
End If
'---------------------------------------------------------------
On Local Error Resume Next
Loads$ = GetFromINI("Intro", "Loads", App.Path + "\KnK.ini")
If Err Then
R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\KnK.ini")
R% = WritePrivateProfileString("List", "clearornot", "Clear", App.Path + "\KnK.ini")
R% = WritePrivateProfileString("Pauses", "TimeKnK", "3", App.Path + "\KnK.ini")
R% = WritePrivateProfileString("Intro", "Loads", "yes", App.Path + "\KnK.ini")
R% = WritePrivateProfileString("Exit", "Loads2", "yes", App.Path + "\KnK.ini")
R% = WritePrivateProfileString("AOL", "aversion", "aol4", App.Path + "\KnK.ini")
R% = WritePrivateProfileString("Scroll", "adver", "yes", App.Path + "\KnK.ini")
R% = WritePrivateProfileString("KnKTheme", "KnKload", "knk1", App.Path + "\KnK.ini")
End If
'---------------------------------------------------------------
If Loads$ = "no" Then
KnKload$ = GetFromINI("KnKTheme", "KnKload", App.Path + "\KnK.ini")
If KnKload$ = "knk1" Then
Form7.Show
Unload Me
End If
If KnKload$ = "knk2" Then
Form13.Show
Unload Me
End If
End If
If Loads$ = "yes" Then


'Plays the opening wave
 
 Call Playwav(App.Path + "\knk.wav")
StayOnTop Me


aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")

'AOL95 startup code
If aversion$ = "aol95" Then
AOLChatSend ("«-×´¯`°   KnK Founders  °´¯`×-»")
TimeOut (0.6)
AOLChatSend ("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.6)
AOLChatSend ("«-×´¯`°  Version: 2.o     °´¯`×-»")
End If

'AOL4.o startup code
If aversion$ = "aol4" Then

Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")

If Color$ = "bgb" Then
 SendChat BlackGreenBlack("«-×´¯`°   KnK Founders  °´¯`×-»")
TimeOut (0.6)
SendChat BlackGreenBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.6)
SendChat BlackGreenBlack("«-×´¯`°  Version: 2.o     °´¯`×-»")
End If

If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°   KnK Founders  °´¯`×-»")
TimeOut (0.6)
SendChat BlackBlueBlack("«-×´¯`°  Server Helper   °´¯`×-»")
TimeOut (0.6)
SendChat BlackBlueBlack("«-×´¯`°  Version: 2.o     °´¯`×-»")
End If
End If

End If

End Sub



Private Sub Picture1_Click()
KnKload$ = GetFromINI("KnKTheme", "KnKload", App.Path + "\KnK.ini")
If KnKload$ = "knk1" Then
Form7.Show
Unload Me
End If
If KnKload$ = "knk2" Then
Form13.Show
Unload Me
End If
End Sub
