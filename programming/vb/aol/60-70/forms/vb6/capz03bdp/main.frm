VERSION 5.00
Object = "{BC326F64-5766-11D5-9845-001E5AC10000}#3.0#0"; "CHATSCAN20.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form main 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock win 
      Left            =   2280
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1680
      Top             =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   150
   End
   Begin TBChatScan20.TBScan TBScan2 
      Left            =   240
      Top             =   1800
      _ExtentX        =   2196
      _ExtentY        =   953
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   3240
   End
   Begin VB.ListBox com 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1410
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin TBChatScan20.TBScan TBScan1 
      Left            =   960
      Top             =   1800
      _ExtentX        =   2196
      _ExtentY        =   953
   End
   Begin VB.Label Label2 
      Height          =   15
      Left            =   1320
      TabIndex        =   6
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      Height          =   135
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
On Error Resume Next
Me.Hide
Form1.Show
Form1.Hide
On Error Resume Next
End Sub


Private Sub Form_Load()
On Error Resume Next
FormNotOnTop sup
Call MsgBox("If you have any questions or need help with this example, E-mail me at bloodhoundbdp@myself.com", vbInformation)
com.AddItem ".xecho - ends echo of x"
com.AddItem ".exit - why do this fag"
com.AddItem ".cmds - views commz"
com.AddItem ".ver - tells aol version"
com.AddItem ".clr - clears chat"
com.AddItem ".adv - advertizment"
com.AddItem ".time - tells time"
com.AddItem ".imon - turns im'z on"
com.AddItem ".imoff - turns im'z off"
com.AddItem ".con - con\con bot"
com.AddItem ".afkon - turns afk on"
com.AddItem ".afkoff - turns afk off"
com.AddItem ".eat - eats chat"
com.AddItem ".echo s/n - echo's s/n"
com.AddItem ".scroll x - scrolls x"
com.AddItem ".write s/n - writes to s/n"
com.AddItem ".ip - tells your ip address"
com.AddItem ".im s/n - opens new im to s/n"
com.AddItem ".sup x - says sup to x"
com.AddItem ".cap x - sets aol caption to x"
com.AddItem ".kw x - goes to kw x"
com.AddItem ".pr x - goes to pr x"
TBScan1.Scan_On
On Error Resume Next
End Sub


Private Sub TBScan2_Scan(Screen_Name As String, What_Said As String)
On Error Resume Next
If Screen_Name = Text1.text Then
If LCase(What_Said) = "" Then
Else
Pause 0.4
If Screen_Name = Text1.text Then
ChatSend "<font color=" & """" & "#er9900" & """" & "><font face=" & """" & "abadi mt condensed light" & """" & "><B>" + Screen_Name + ":</B></html> <font color=" & """" & "#000000" & """" & "><font face=" & """" & "arial narrow" & """" & ">" & (What_Said)
Pause 0.6
End If
End If
End If
On Error Resume Next
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Label2.Caption = Val(Label1.Caption) + 1
Label1.Caption = (Label2.Caption)
ChatSend (fblAck) & (arialf) & (GetUser) & "  Hå§ ßèèñ å† ¡ÐLè Før " & (Label2.Caption) & " M¡ñ(§)"
On Error Resume Next
End Sub

Private Sub TBScan1_Scan(Screen_Name As String, What_Said As String)
'all codes did by my self not a one was from source i made them all
'so don't steal, this is made to learn from not to cheat me out of
'or at lest put me in the creds if u make any. I made this so easy
'a kid can make one out of it to chang the name of capz the Adv is
'what you need just chang it to wat ever put bdp in it as a programer
'too if ya would well this is the shit just work off this
On Error Resume Next
Dim eat As String
Dim eat2 As String
Dim fblAck As String
Dim fred As String
Dim TWhite As String
Dim con As String
Dim clr As String
Dim Clr2 As String
Dim Adv As String
Dim arialf As String
eat = "< a href=""><font color=" & """" & "#fffffe" & """" & "><i                                                                                                                                                                                                                                                                                                                                                                                                                                                       ><i                                                                                                                                                                                                                                                                                            >"
eat2 = "<i                                                                                                                                                                                                                                                                                                                                                                                                                                                       ><i                                                                                                                                                                                                                                                                                            >"
Adv = "< a href=" & """" & "bdp2k.250x.com" & """" & "></u><font color=" & """" & "#000000" & """" & "><font face=" & """" & "arial narrow" & """" & ">(Example)•˜Ca[]Dz v0.3 ßý Bdp˜•(Example)</a>"
fblAck = "<font color=" & """" & "#000000" & """" & ">"
fred = "<font color=" & """" & "#er9900" & """" & ">"
TWhite = "<font color=" & """" & "#fffffe" & """" & ">"
arialf = "<font face=" & """" & "arial narrow" & """" & ">"
con = "<font color=" & """" & "#fffffe" & """" & ">{< a href=""><font color=" & """" & "#fffffe" & """" & ">s con\con</a>"
clr = (TWhite) & "< a href="">.<I @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @"
Clr2 = "@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @@ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @ @></a>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'commands
If Screen_Name = GetUser Then
If LCase(What_Said) = ".exit" Then
dos32.FormExitUp Me
ChatSend (Adv)
Pause 0.2
End
End If
If LCase(What_Said) = ".cmds" Then
main.Show
End If
If LCase(What_Said) = ".ver" Then
ChatSend (fblAck) & (arialf) & (GetUser) & " I§ ù§ïñg " & (AOLversion)
End If
If LCase(What_Said) = ".afkon" Then
If Timer2.Enabled = True Then
ChatSend (fblAck) & (arialf) & "ÅFK I§ Already Øñ"
Else
ChatSend (fblAck) & (arialf) & "ÅFK I§ Øñ"
Timer2.Enabled = True
End If
End If
If LCase(What_Said) = ".afkoff" Then
If Timer2.Enabled = False Then
ChatSend (fblAck) & (arialf) & "ÅFK I§ Already Øff"
Else
ChatSend (fblAck) & (arialf) & "ÅFK I§ Øff"
Timer2.Enabled = False
Label1.Caption = "0"
End If
End If
If LCase(What_Said) = ".eat" Then
ChatSend (clr) & (Clr2)
ChatSend (eat) & (eat2)
ChatSend (eat) & (eat2)
Pause 1.5
ChatSend (eat) & (eat2)
ChatSend (eat) & (eat2)
Pause 1.5
ChatSend (Adv)
End If
If LCase(What_Said) = ".con" Then
ChatSend (con)
End If
If LCase(What_Said) = ".adv" Then
ChatSend (Adv)
End If
If LCase(What_Said) = ".clr" Then
ChatSend (clr) & (Clr2)
Pause 0.3
ChatSend (clr) & (Clr2)
Pause 0.3
ChatSend (clr) & (Clr2)
Pause 1.5
ChatSend (Adv)
End If
If LCase(What_Said) = ".imon" Then
IMsOn
ChatSend (fblAck) & (arialf) & "Im'z årè ñøw øñ"
End If
If LCase(What_Said) = ".time" Then
ChatSend (fblAck) & (arialf) & "I†'z " & (Time)
End If
If LCase(What_Said) = ".ip" Then
ip.Show
End If
If LCase(What_Said) = ".imoff" Then
IMsOff
ChatSend (fblAck) & (arialf) & "Im'z årè ñøw: øff"
End If
If LCase(What_Said) = ".xecho" Then
If Text1.text = "" Then
ChatSend (fblAck) & (arialf) & "Ñø†hïñg †ø §†øþ"
Else
ChatSend (fblAck) & (arialf) & "Èçhø Hå§ §†øþèÐ"
Text1.text = ""
TBScan2.Scan_Off
End If
End If
End If
'This is like you say .wb bdp and it takes out .wb
'only the name bdp will be the out come
Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If Screen_Name$ = GetUser$ And InStr(What_Said, ".") = 1& Then
      lngSpace& = InStr(What_Said$, " ")
      If lngSpace& = 0& Then
         strCommand$ = What_Said$
      Else
         strCommand$ = Left(What_Said$, lngSpace& - 1&)
      End If
      Select Case strCommand$

Case ".wb"
strArgument1$ = Mid(What_Said, 5)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
ChatSend (fblAck) & (arialf) & "WèLçømè ßåçK •" & (strArgument1$) & "• Høw Wåz ýøùR åçiÐ †Ríþ?¿?"
End If
End If

Case ".echo"
strArgument1$ = Mid(What_Said, 7)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
Text1.text = (strArgument1$)
ChatSend (fblAck) & (arialf) & "Ñøw Èçhøïñg •" & (strArgument1$) & "•"
TBScan2.Scan_On
End If
End If
Case ".write"
strArgument1$ = Mid(What_Said, 8)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "Ñø S/n †ø write"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "Ñø S/n †ø write"
Else
ChatSend (fblAck) & (arialf) & "Wrí†ïñg MåïL †ø " & (strArgument1$)
Pause 0.2
SendKeys "^m"
Pause 0.3
SendKeys (strArgument1$)
End If
End If

Case ".scroll"
strArgument1$ = Mid(What_Said, 9)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "†èx† ÑèèÐèÐ †ø §ç®øLL"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "†èx† ÑèèÐèÐ †ø §ç®øLL"
Else
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
Pause 0.6
ChatSend (fblAck) & (arialf) & (strArgument1$)
End If
End If

Case ".kw"
strArgument1$ = Mid(What_Said, 5)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "kw ÑèèÐèÐ †ø goto"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "kw ÑèèÐèÐ †ø goto"
Else
ChatSend (fblAck) & (arialf) & "Gøiñg †ø KèýWøRÐ: " & (strArgument1$)
Keyword (strArgument1$)
End If
End If

Case ".pr"
Dim ee As String
ee = Text_ChatTitle
strArgument1$ = Mid(What_Said, 5)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "Pr ÑèèÐèÐ †ø goto"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "Pr ÑèèÐèÐ †ø goto"
Else
ChatSend (fblAck) & (arialf) & "Gøiñg †ø Pr: " & (strArgument1$)
PrivateRoom (strArgument1$)
If ee = Text_ChatTitle Then
ChatSend (fblAck) & (arialf) & (strArgument1$) & " Iz Full"
Else
ChatSend (fblAck) & (arialf) & (GetUser) & " Entered " & (strArgument1$)
End If
End If
End If
Case ".cap"
strArgument1$ = Mid(What_Said, 6)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "caption ÑèèÐèÐ †ø chang"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "caption ÑèèÐèÐ †ø chang"
Else
Change_AOL_Caption (strArgument1$)
ChatSend (fblAck) & (arialf) & "Åøl'z Ñèw çåptiøñ Is " & (strArgument1$)
End If
End If

Case ".sup"
strArgument1$ = Mid(What_Said, 6)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
ChatSend (fblAck) & (arialf) & "ýø •" & (strArgument1$) & "• §ùþ Mý Ñìggå¿?¿?"
End If
End If

Case ".im"
strArgument1$ = Mid(What_Said, 5)
If strArgument1$ = "" Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
If strArgument1$ = " " Then
ChatSend (fblAck) & (arialf) & "ïf Yøü Døñ'† Kñøw †hå† å Ñåmè i§ ñèèÐèÐ, ýøù ñèèÐ Hèlþ!"
Else
Keyword "Aol://9293:" & (strArgument1$)
ChatSend (fblAck) & (arialf) & (GetUser) & " ì§ ìñ§†äñ† Mè§§ågïñg " & (strArgument1$)
End If
End If
End Select
End If
On Error Resume Next
End Sub
Private Sub Timer1_Timer()
FormTop Me
On Error Resume Next
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If Text_ChatTitle = "" Then
Else
Adv = "< a href=" & """" & "bdp2k.250x.com" & """" & "></u><font color=" & """" & "#000000" & """" & "><font face=" & """" & "arial narrow" & """" & ">(Example)•˜Ca[]Dz v0.3 ßý Bdp˜•(Example)</a>"
ChatSend (Adv)
Change_AOL_Caption "(Example)Capz v0.3 By Bdp(Example)"
Hide_Welcome
Timer3.Enabled = False
End If
On Error Resume Next
End Sub

