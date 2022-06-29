VERSION 5.00
Begin VB.Form idle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "idle example"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C0C0C0&
      Height          =   1290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "s&ave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3000
         TabIndex        =   13
         ToolTipText     =   "save reasons."
         Top             =   280
         Width           =   495
      End
      Begin VB.CommandButton command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   9
         ToolTipText     =   "add reason."
         Top             =   280
         Width           =   495
      End
      Begin VB.CheckBox stats 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&up stats"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox time 
         BackColor       =   &H00C0C0C0&
         Caption         =   "t&ime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox reason 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&reason"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "s&top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "stop idle."
         Top             =   645
         Width           =   975
      End
      Begin VB.CommandButton command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "start idle."
         Top             =   645
         Width           =   975
      End
      Begin VB.ComboBox combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         ItemData        =   "idle.frx":0000
         Left            =   720
         List            =   "idle.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "type in reason here."
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         TabIndex        =   1
         ToolTipText     =   "menu."
         Top             =   645
         Width           =   975
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "reasons - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   120
         TabIndex        =   17
         Top             =   280
         Width           =   600
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "scroll options -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   870
      End
   End
   Begin VB.Label lblday 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   600
      TabIndex        =   16
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label lblhr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   15
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label lblmin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   14
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label lblmins 
      AutoSize        =   -1  'True
      Caption         =   "min: n/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   600
      TabIndex        =   12
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label lblper 
      AutoSize        =   -1  'True
      Caption         =   "per: n/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   600
      TabIndex        =   11
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label lblfile 
      AutoSize        =   -1  'True
      Caption         =   "file: n/a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   600
      TabIndex        =   10
      Top             =   1560
      Width           =   465
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuims 
         Caption         =   "&ims off when idling"
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuupstats 
         Caption         =   "scroll upload stats"
      End
      Begin VB.Menu mnustats 
         Caption         =   "scroll idle stats history"
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuidle 
         Caption         =   "scroll how long u been idle"
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnureas 
         Caption         =   "clear reason history"
      End
      Begin VB.Menu mnustats2 
         Caption         =   "clear idle stats history"
      End
   End
End
Attribute VB_Name = "idle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If mnuims.Checked = True Then
Call IMsOff
Pause (1)
End If

If timer1.Enabled = True Then rusure = MsgBox("are you sure you want to start over your idling session?", vbYesNo + vbSystemModal)
If rusure = vbYes Then
lblday = "0": lblhr = "0": lblmin = "0"
End If
timer1.Enabled = True

ChatSend "idle example - on"
ChatSend "" + lblmin + " /min. " + lblhr + " /hr. " + lblday + " /day."

End Sub
Private Sub Command2_Click()
If mnuims.Checked = True Then
Call IMsOn
Pause (1)
End If

If timer1.Enabled = False Then Exit Sub
timer1.Enabled = False

ChatSend "idle example - off"
ChatSend "" + lblmin + " /min. " + lblhr + " /hr. " + lblday + " /day."

lblday = "0": lblhr = "0": lblmin = "0"
End Sub
Private Sub Command3_Click()
PopupMenu file, 0, command3.Left, command3.Top + command3.Height, mnuims
End Sub
Private Sub Command4_Click()
rusure = MsgBox("are you sure you want to add the reason?", vbYesNo + vbSystemModal)
If rusure = vbYes Then
combo1.AddItem (combo1.Text)
End If
End Sub
Private Sub Command5_Click()
rusure = MsgBox("are you sure you want to save reasons?", vbYesNo + vbSystemModal)
If rusure = vbYes Then
Call SaveComboBox(App.Path + "\reason.ini", combo1)
End If
End Sub
Private Sub Form_Load()
'example by mecca
'i dont take much credit because every sub was from dos32.bas,
'i just put them together, so i dont care what you do with this example
'so please just give me some credit -=)
'one more thing i didnt test this example out because i dont got aol
'but im pretty sure it works because i made an idle bot just like this example
'but this example is better then my idle because this took less coding
'so if theres any bugz just figure out how to make it work then
On Error Resume Next
If App.PrevInstance = True Then End
FormOnTop Me

a$ = GetFromINI("idle", "min", App.Path + "\idle.ini")
b$ = GetFromINI("idle", "hr", App.Path + "\idle.ini")
c$ = GetFromINI("idle", "day", App.Path + "\idle.ini")
d$ = GetFromINI("ims", "ims", App.Path + "\idle.ini")

If d$ = "" Then
d$ = "no"
WriteToINI "ims", "ims", d$, App.Path + "\idle.ini"
End If

If d$ = "yes" Then mnuims.Checked = True
If d$ = "no" Then mnuims.Checked = False

sfile = FileExists(App.Path + "\reason.ini")
If sfile = True Then
Call LoadComboBox(App.Path + "\reason.ini", combo1)
combo1.ListIndex = 0
End If

If a$ = "" Then a$ = "0"
If b$ = "" Then b$ = "0"
If c$ = "" Then c$ = "0"

ChatSend "<font face=arial>idle example - load"
ChatSend "<font face=arial>" + a$ + " /min. " + b$ + " /hr. " + c$ + " /day."
End Sub
Private Sub Form_Unload(Cancel As Integer)
a$ = GetFromINI("idle", "min", App.Path + "\idle.ini")
b$ = GetFromINI("idle", "hr", App.Path + "\idle.ini")
c$ = GetFromINI("idle", "day", App.Path + "\idle.ini")

If a$ = "" Then a$ = "0"
If b$ = "" Then b$ = "0"
If c$ = "" Then c$ = "0"

ChatSend "<font face=arial>idle example - unload"
ChatSend "<font face=arial>" + a$ + " /min. " + b$ + " /hr. " + c$ + " /day."

End
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
Private Sub mnuidle_Click()
sn$ = GetUser()
ChatSend "" + sn$ + " - been idled for"
ChatSend "" + lblmin + " /min. " + lblhr + " /hr. " + lblday + " /day."
End Sub
Private Sub mnuims_Click()
If mnuims.Checked = True Then
mnuims.Checked = False
a$ = GetFromINI("ims", "ims", App.Path + "\idle.ini")
a$ = "no"
WriteToINI "ims", "ims", a$, App.Path + "\idle.ini"
Exit Sub
End If
If mnuims.Checked = False Then
mnuims.Checked = True
a$ = GetFromINI("ims", "ims", App.Path + "\idle.ini")
a$ = "yes"
WriteToINI "ims", "ims", a$, App.Path + "\idle.ini"
Exit Sub
End If
End Sub
Private Sub mnureas_Click()
rusure = MsgBox("are you sure you want to clear reasons history?", vbYesNo + vbSystemModal)
If rusure = vbYes Then
combo1.Clear
Call SaveComboBox(App.Path + "\reason.ini", combo1)
End If
End Sub
Private Sub mnustats_Click()
a$ = GetFromINI("idle", "min", App.Path + "\idle.ini")
b$ = GetFromINI("idle", "hr", App.Path + "\idle.ini")
c$ = GetFromINI("idle", "day", App.Path + "\idle.ini")

If a$ = "" Then a$ = "0"
If b$ = "" Then b$ = "0"
If c$ = "" Then c$ = "0"

ChatSend "<font face=arial>idle example - stats"
ChatSend "<font face=arial>" + a$ + " /min. " + b$ + " /hr. " + c$ + " /day."
End Sub
Private Sub mnustats2_Click()
rusure = MsgBox("are you sure you want to clear your idle stats history?", vbYesNo + vbSystemModal)
If rusure = vbYes Then
a$ = GetFromINI("idle", "min", App.Path + "\idle.ini")
a$ = "0"
WriteToINI "idle", "min", a$, App.Path + "\idle.ini"

b$ = GetFromINI("idle", "hr", App.Path + "\idle.ini")
b$ = "0"
WriteToINI "idle", "hr", b$, App.Path + "\idle.ini"

c$ = GetFromINI("idle", "day", App.Path + "\idle.ini")
c$ = "0"
WriteToINI "idle", "day", c$, App.Path + "\idle.ini"
End If
Exit Sub
End Sub
Private Sub mnuupstats_Click()
Dim AOLModal As Long
Dim AOLStatic As Long
Dim AOLStatic2 As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)

lblfile = GetText(LCase(AOLStatic&))
lblfile = ReplaceString(LCase(lblfile), "Now Uploading ", "file:")

lblper = GetText(LCase(AOLModal&))
lblper = ReplaceString(LCase(lblper), "File Transfer - ", "pr: ")
lblper = ReplaceString(LCase(lblper), "%", "")
lblper = ReplaceString(LCase(lblper), "File Transfer", "pr: n/a")

lblmins = GetText(LCase(AOLStatic2&))
lblmins = ReplaceString(LCase(lblmins), "About", "mn:")
lblmins = ReplaceString(LCase(lblmins), "minutes", "")
lblmins = ReplaceString(LCase(lblmins), "Calculating Transfer Time", "mn: n/a")
lblmins = ReplaceString(LCase(lblmins), "Less than a ", "mn: n/a")
lblmins = ReplaceString(LCase(lblmins), "minute", "")
lblmins = ReplaceString(LCase(lblmins), "remaining.", "")

ChatSend "idle example - " + lblfile + ""
ChatSend "" + lblper + " " + lblmins + ""
End Sub
Private Sub timer1_Timer()
Dim AOLModal As Long
Dim AOLStatic As Long
Dim AOLStatic2 As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)

AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)

lblfile = GetText(LCase(AOLStatic&))
lblfile = ReplaceString(LCase(lblfile), "Now Uploading ", "file:")

lblper = GetText(LCase(AOLModal&))
lblper = ReplaceString(LCase(lblper), "File Transfer - ", "pr: ")
lblper = ReplaceString(LCase(lblper), "%", "")
lblper = ReplaceString(LCase(lblper), "File Transfer", "pr: n/a")

lblmins = GetText(LCase(AOLStatic2&))
lblmins = ReplaceString(LCase(lblmins), "About", "mn:")
lblmins = ReplaceString(LCase(lblmins), "minutes", "")
lblmins = ReplaceString(LCase(lblmins), "Calculating Transfer Time", "mn: n/a")
lblmins = ReplaceString(LCase(lblmins), "Less than a ", "mn: n/a")
lblmins = ReplaceString(LCase(lblmins), "minute", "")
lblmins = ReplaceString(LCase(lblmins), "remaining.", "")

If AOLModal& = False Then
lblfile = "file: n/a"
lblper = "pr: n/a"
lblmins = "mn: n/a"
End If

lblmin = lblmin + 1
If lblmin = "60" Then lblhr = lblhr + 1: lblmin = "0"
If lblhr = "24" Then lblday = lblday + 1: lblhr = "0"

If reason.Value = 1 And time.Value = 1 And stats.Value = 1 Then
ChatSend "" + lblmin + " /min. " + lblhr + " /hr. " + lblday + " /day."
ChatSend "" + lblfile + " " + lblper + " " + lblmins + ""
ChatSend "  reason: " + combo1.Text + ""
End If

If reason.Value = 1 And time.Value = 1 And stats.Value = 0 Then
ChatSend "idle example - idling"
ChatSend "" + lblmin + " /min. " + lblhr + " /hr. " + lblday + " /day."
ChatSend "     reason: " + combo1.Text + ""
End If

If reason.Value = 0 And time.Value = 1 And stats.Value = 1 Then
ChatSend "idle example - idling"
ChatSend "" + lblmin + " /min. " + lblhr + " /hr. " + lblday + " /day."
ChatSend "" + lblfile + " " + lblper + " " + lblmins + ""
End If

If reason.Value = 1 And time.Value = 0 And stats.Value = 1 Then
ChatSend "idle example - idling"
ChatSend "" + lblfile + " " + lblper + " " + lblmins + ""
ChatSend "  reason: " + combo1.Text + ""
End If

If reason.Value = 0 And time.Value = 1 And stats.Value = 0 Then
ChatSend "idle example - idling"
ChatSend "" + lblmin + " /min. " + lblhr + " /hr. " + lblday + " /day."
End If

If reason.Value = 1 And time.Value = 0 And stats.Value = 0 Then
ChatSend "idle example - idling"
ChatSend "  reason: " + combo1.Text + ""
End If

If reason.Value = 0 And time.Value = 0 And stats.Value = 1 Then
ChatSend "idle example - " + lblfile + ""
ChatSend "idling - " + lblper + " " + lblmins + ""
End If

a$ = GetFromINI("idle", "min", App.Path + "\idle.ini")
a$ = Val(a$) + 1
WriteToINI "idle", "min", a$, App.Path + "\idle.ini"

b$ = GetFromINI("idle", "hr", App.Path + "\idle.ini")
If a$ = "60" Then b$ = Val(b$) + 1: a$ = "0"
WriteToINI "idle", "min", a$, App.Path + "\idle.ini"
WriteToINI "idle", "hr", b$, App.Path + "\idle.ini"

c$ = GetFromINI("idle", "day", App.Path + "\idle.ini")
If b$ = "24" Then c$ = Val(c$) + 1: b$ = "0"
WriteToINI "idle", "hr", b$, App.Path + "\idle.ini"
WriteToINI "idle", "day", c$, App.Path + "\idle.ini"

End Sub
