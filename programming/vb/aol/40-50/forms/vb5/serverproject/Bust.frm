VERSION 5.00
Begin VB.Form Bust 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1710
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   30
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000040&
      Caption         =   "Send Try(s) To Chat"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Value           =   1  'Checked
      WhatsThisHelpID =   30
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000040&
      Caption         =   "Minimize After Busting"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   720
      WhatsThisHelpID =   30
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000040&
      Caption         =   "Close After Busting"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      WhatsThisHelpID =   30
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X"
      Height          =   210
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      WhatsThisHelpID =   30
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop "
      Height          =   210
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      WhatsThisHelpID =   30
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start "
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      WhatsThisHelpID =   30
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000006&
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Text            =   "Room Here"
      Top             =   600
      WhatsThisHelpID =   30
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Bust.frx":0000
      Top             =   0
      WhatsThisHelpID =   30
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   975
      Left            =   1800
      Top             =   360
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   1680
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready..."
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      WhatsThisHelpID =   30
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   195
      Left            =   -120
      Picture         =   "Bust.frx":0021
      Top             =   360
      WhatsThisHelpID =   30
      Width           =   1500
   End
End
Attribute VB_Name = "Bust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
If Check1.Value = 1 Then Check2.Value = 0
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then Check1.Value = 0
End Sub

Private Sub Command1_Click()
Combo1.SetFocus
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
If AOL& = 0 Then
    MsgBox "America Online Is Not Open.", 16
    Exit Sub
End If
Wel = FindChildByTitle(AOLMDI(), "Welcome")
Welc$ = String(255, 0)
WhichWel = GetWindowText(Wel, Welc$, 250)
If WhichWel < 8 Then
    MsgBox "You Need To Sign On First To Use The Server.", 16
    Exit Sub
End If
a$ = Combo1.Text
If Len(a$) Then
 If LCase(a$) Like LCase("Room Here") Then Exit Sub
    Room$ = "aol://2719:2-2-" & a$
    AOL& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
     ChatRoom& = child&
     If LCase(GetText(ChatRoom&)) Like LCase(a$) Then Exit Sub
     Call SendMessage(ChatRoom&, WM_CLOSE, 0, 0&)
    Else
         Do
            DoEvents
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                ChatRoom& = child&
                If LCase(GetText(ChatRoom&)) Like LCase(a$) Then Exit Sub
                Call SendMessage(ChatRoom&, WM_CLOSE, 0, 0&)
            End If
        Loop Until child& = 0&
    End If
    DlogStop = False
    RoomBust = True
    T = 1
    Do
     If RoomBust = False Then Exit Sub
      DoEvents
        Call Keyword(Room$)
        Label1.Caption = "Busting..Please Wait"
        Do
          DoEvents
            ErrWin = FindWindow("#32770", "America Online")
            If ErrWin Then
               Call SendMessageByNum(ErrWin, WM_CLOSE, 0, 0)
                T = T + 1
            End If
            AOL& = FindWindow("AOL Frame25", vbNullString)
            mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
            child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
        Loop Until (Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0&) Or ErrWin <> 0&
   Loop Until (Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0&) Or DlogStop
   If DlogStop = False Then
        Label1.Caption = "Ready."
        If Check3.Value = 1 Then
            Timeout (1)
            SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–X-Treme Server ' 99 Room Buster")
            SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Busted in [" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & a$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] In " & Trim(Str$(T)) & " Try(s)":  DoEvents
           
        End If
        If Check1.Value = 1 Then Unload Me
        If Check2.Value = 1 Then Me.WindowState = 1
   End If
   RoomBust = False
Else
MsgBox "Please Enter a The Name Of The Private Room To Bust In."
Combo1.SetFocus
End If
End Sub

Private Sub Command2_Click()
RoomBust = False
Combo1.SetFocus
End Sub

Private Sub Command3_Click()
Combo1.SetFocus
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
StayOnTop Me
Combo1.AddItem "Server"
Combo1.AddItem "Server0"
Combo1.AddItem "Server1"
Combo1.AddItem "Server2"
Combo1.AddItem "Server3"
Combo1.AddItem "Server4"
Combo1.AddItem "Server5"
Combo1.AddItem "Server6"
Combo1.AddItem "Server7"
Combo1.AddItem "Server8"
Combo1.AddItem "Server9"
Combo1.AddItem "Server10"
Combo1.AddItem "Server11"
Combo1.AddItem "Server12"
Combo1.AddItem "Server13"
Combo1.AddItem "Server14"
Combo1.AddItem "Server15"
Combo1.AddItem "Server16"
Combo1.AddItem "Server17"
Combo1.AddItem "Server18"
Combo1.AddItem "Server19"
Combo1.AddItem "Server20"
Combo1.AddItem "AOL40"
Combo1.AddItem "Vb"
Combo1.AddItem "Vb1"
Combo1.AddItem "Vb2"
Combo1.AddItem "Vb3"
Combo1.AddItem "Vb4"
Combo1.AddItem "Vb5"
Combo1.AddItem "Vb6"
Combo1.AddItem "Warez"
Combo1.AddItem "Fate"
Combo1.AddItem "Fate1"
Combo1.AddItem "Fate2"
Combo1.AddItem "Fate3"
Combo1.AddItem "Fate4"
Combo1.AddItem "Fate5"
Combo1.AddItem "Fate6"
End Sub

