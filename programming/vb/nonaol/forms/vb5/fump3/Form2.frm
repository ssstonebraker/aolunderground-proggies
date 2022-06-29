VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Fucked Up Mp3 Example -   By Xen"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox File 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox RandFileList 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4020
      Width           =   2295
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   1080
      Top             =   600
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   4215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Multi As Boolean
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
Dim lngSpace As Long, xenCommand As String, xenArgument1 As String
   Dim xenArgument2 As String, lngComma As Long
   If Screen_Name$ = GetUser$ And InStr(What_Said, ".") = 1& Then
      lngSpace& = InStr(What_Said$, " ")
      If lngSpace& = 0& Then
         xenCommand$ = What_Said$
      Else
         xenCommand$ = Left(What_Said$, lngSpace& - 1&)
      End If
      Select Case xenCommand$
      Case ".pause"
      MediaPlayer1.Pause
      AoL4_ChatSendBlue "</b>• Fucked Up mp3 example - mp³ Player Paused"
      
       Case ".resume"
      MediaPlayer1.Play
      AoL4_ChatSendBlue "</b>• Fucked Up mp3 example - mp³ Player Resumed"

Case ".cancel"
If Multi = True Then
    List2.Clear
    Multi = False
    AoL4_ChatSendBlue "</b>• Fucked Up mp3 example  [multiple strings cancelled]"
End If


Case ".exit"
AoL4_ChatSendBlue "</b>,.·~°'º°”˜,.·~°˜`°~·.,˜`°º'°~·.,"
TimeOut 0.9
AoL4_ChatSendBlue "</b>    Fucked Up Mp3 example"
TimeOut 0.9
AoL4_ChatSendBlue "</b>                By: Xen"
TimeOut 0.9
AoL4_ChatSendBlue "</b>    UnLoaded By: " + LCase(GetUser) + ""
TimeOut 0.9
AoL4_ChatSendBlue "</b>    º '°~·.,'°~·.,,.·~°,.·~°'º"
End


Case ".vol"
If Len(What_Said$) > 4& Then
               xenArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
               If xenArgument1$ > 100 Then
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example - Invalid Property Assignment"
Else
                Pause 0.5
            
            MediaPlayer1.Volume = (xenArgument1$ * 25) - 2500
            AoL4_ChatSendBlue "</b>• Fucked up mp3 example - mp³ Volume at " + xenArgument1 + " %"
    End If
    End If
    
Case ".rand"
On Error Resume Next
            File1.ListIndex = RandomNumber1(File1.ListCount)
            xenArgument1$ = File1.Path & "\" & File1.filename
            On Error Resume Next
            MediaPlayer1.Open xenArgument1$
            MediaPlayer1.Play
            AoL4_ChatSendBlue "</b>• Fucked Up mp3 example - Playing [" & ReplaceString(ReplaceString(xenArgument1$, File1.Path, ""), "\", "") & "]"

Case ".play"
If Multi <> True Then
    Call PlayIt(Right(What_Said, Len(What_Said) - 6))
    Exit Sub
End If
If Multi = True Then
    On Error GoTo kdmerr
    intt% = Right(What_Said, Len(What_Said) - 6)
    MediaPlayer1.Open File1.Path & "\" & List2.List(intt%)
    AoL4_ChatSendBlue "</b<• Fucked Up mp3 example now playing [" & ReplaceString(List2.List(intt%), ".mp3", "") & "]"
    Multi = False
    List2.Clear
    Exit Sub
kdmerr:
    If What_Said <> "</b>• Fucked Up mp3 example [type ''.play'' and 0 - " & List2.ListCount - 1 & "or ''.cancel'']" Then AoL4_ChatSendBlue "</b>• Fucked Up mp3 example [invaled number]"
End If


Case ".stop"
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example Now Stopped"
MediaPlayer1.Stop


Case ".mute"
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example Now muted"
MediaPlayer1.Mute = True


Case ".xmute"
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example is now unmuted"
MediaPlayer1.Mute = False

Case ".dir"
xenArgument1$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = xenArgument1$
     Call WriteToINI("mp3dir", "dir", xenArgument1$, App.Path & "\mp3.ini")
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example dir is " + xenArgument1 + ""

Case ".mp3s?"
win$ = File1.ListCount
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example • total mp³'s [" & win & "]"

Case ".mini"
Me.Hide
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example · mp³ player minimized"

Case ".xmini"
Me.Show
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example · mp³ player unminimized"

Case ".length"
Dim lngSeconds As Long
Dim lngMinutes As Long
Dim strSeconds As String
Dim strMinutes As String
Dim lngPos As Long
    strSeconds$ = MediaPlayer1.Duration
    If InStr(strSeconds$, ".") <> 0& Then
        strSeconds$ = Right(strSeconds$, Len(strSeconds$) - InStr(strSeconds$, "*"))
    End If
    lngSeconds& = strSeconds$
    Do Until lngSeconds& <= 59
        DoEvents
        lngSeconds& = lngSeconds& - 60
        lngMinutes& = lngMinutes& + 1
    Loop
    strSeconds$ = Format$(lngSeconds&, "00")
    strMinutes$ = Format$(lngMinutes&, "00")
    SongLength$ = strMinutes$ & ":" & strSeconds$
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example · total time: " + SongLength + ""

Case ".refdir"
File1.Refresh
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example · mp3 list refreshed"

Case ".loop"
MediaPlayer1.PlayCount = 0
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example · looping now enabled"

Case ".xloop"
MediaPlayer1.PlayCount = 1
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example · looping now disabled"

Case ".pa"
Do
For i = 0 To File1.ListCount - 1
MediaPlayer1.Open File1.Path & "\" & File1.List(i)
AoL4_ChatSendBlue "</b>• Fucked Up mp3 example · playing - [ " + File1.List(i) + " ]"
TimeOut 4
TimeOut (MediaPlayer1.Duration)
Next i
Loop Until i = File1.ListCount - 1

Case ".ran"
'thanks to one of my homeboys for helpin
'me out with this random string, code..
'so thanks romey
Call PlayRandomGrp(Right(What_Said, Len(What_Said) - lngSpace&))

End Select
End If
End Sub

Public Sub PlayRandomFile()
            Randomize
            AudioFiles.ListIndex = Int(Val(AudioFiles.ListCount * Rnd) + 1)
If AudioFiles.filename = "" Then AudioFiles.filename = AudioFiles.List(0)
MediaPlayer1.Open AudioFiles.Path & "\" & AudioFiles.filename
File = AudioFiles.filename
sendchat2 "random -" & LCase(ReplaceString(File, ".mp3", "")) & ""
End Sub
'
'__________________________________________________
'string random sub
__________________________________________________
Public Sub PlayRandomGrp(RndFile As String)
If RndFile = ".ran" Then Exit Sub
For i = 0 To File1.ListCount - 1
If InStr(LCase(Trim(File1.List(i))), LCase(Trim(RndFile))) Then
RandFileList.AddItem File1.List(i)
End If
Next i

If RandFileList.ListCount = 1 Then
MediaPlayer1.Open AudioFiles.Path & "\" & RandFileList.List(0)
File = RandFileList.List(0)
AoL4_ChatSendBlue "random - " & LCase(ReplaceString(File, ".mp3", "")) & ""
RandFileList.Clear
Exit Sub
End If

If RandFileList.ListCount > 1 Then
Randomize
RandFileList.ListIndex = Int(Val(RandFileList.ListCount * Rnd) + 1)
If RandFileList = "" Then RandFileList = RandFileList.List(0)
MediaPlayer1.Open File1.Path & "\" & RandFileList
File = RandFileList
AoL4_ChatSendBlue "random - " & LCase(ReplaceString(File, ".mp3", "")) & ""
RandFileList.Clear
Exit Sub
End If

If RandFileList.ListCount = 0 Then
   AoL4_ChatSendBlue "not found " & LCase(RndFile) & "": Exit Sub
Exit Sub
End If

End Sub
Private Sub Form_Load()
'This example was made in a hurry.
'If it is all "Fucked Up" then
'You should know API to fix it
'The bas i used with this was
'dos32.bas. Edited by me
'I used it, because I haven't finished mine
'Dos32 is a pretty good bas to use for many
'Things.  But don't become too attached.
'Cause there are bigger and better
'bas' coming out everyday
'If you have Questions About This Example
'Please Mail me at, questions4xen@juno.com
On Error Resume Next
File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.ini")
Call FormOnTop(Form2)
Call Form_Center(Form2)
Chat1.ScanOn
List1.AddItem "                                 Mp³ Player Options       "
List1.AddItem "* - means something goes after command, like a song, etc.."
List1.AddItem ".pause - pauses the playing mp3"
List1.AddItem ".resume - resumes the paused mp3"
List1.AddItem ".exit - exits the program"
List1.AddItem ".vol * - changes the volume of the mp3"
List1.AddItem ".rand - plays random mp3"
List1.AddItem ".play * - plays a specific song"
List1.AddItem ".stop - stops the playing mp3"
List1.AddItem ".mute - mutes playing mp3"
List1.AddItem ".xmute - unmutes playing muted mp3"
List1.AddItem ".dir * - sets the mp3 directory"
List1.AddItem ".mp3s? - tells how many mp3's are playing"
List1.AddItem ".mini - minimises the mp3 player form"
List1.AddItem ".xmini - un minimises the mp3 player form"
List1.AddItem ".length - tells how long the song is"
List1.AddItem ".refdir - refreshes the current directory"
List1.AddItem ".loop - loops the song thats playing"
List1.AddItem ".xloop - stops the looping process"

AoL4_ChatSendBlue "</b>,.·~°'º°”˜,.·~°˜`°~·.,˜`°º'°~·.,"
TimeOut 0.9
AoL4_ChatSendBlue "</b>    Fucked Up Mp3 example"
TimeOut 0.9
AoL4_ChatSendBlue "</b>                By: Xen"
TimeOut 0.9
AoL4_ChatSendBlue "</b>    Loaded By: " + LCase(GetUser) + ""
TimeOut 0.9
AoL4_ChatSendBlue "</b>    º '°~·.,'°~·.,,.·~°,.·~°'º"
End Sub
Public Sub PlayIt(sString As String)
    For i = 0 To File1.ListCount - 1
        If InStr(LCase(TrimSpaces(File1.List(i))), LCase(TrimSpaces(sString$))) Then
            List2.AddItem (File1.List(i))
        End If
    Next i
    If List2.ListCount = 1 Then
         MediaPlayer1.Open File1.Path & "\" & List2.List(0)
        AoL4_ChatSendBlue "</b>• Fucked Up Mp3 Example -[" & ReplaceString(List2.List(0), ".mp3", "") & "]"
        List2.Clear
        Exit Sub
    End If
    If List2.ListCount = 0 Then
         AoL4_ChatSendBlue "</b>• Fucked Up Mp3 Example [no strings found]"
    End If
    If List2.ListCount > 1 Then
          usa$ = List2.ListCount
        AoL4_ChatSendBlue "</b>• Fucked Up Mp3 Example  [multiple strings found]"
        For i = 0 To List2.ListCount - 1
            AoL4_ChatSendBlue "•  [" & i & ") " & List2.List(i) & "]"
            Pause 1
        Next i
        AoL4_ChatSendBlue "</b>• Fucked Up Mp3 Example  [type ''.play'' and 0 - " & List2.ListCount - 1 & " or ''.cancel'']"
        Multi = True
    End If
End Sub
