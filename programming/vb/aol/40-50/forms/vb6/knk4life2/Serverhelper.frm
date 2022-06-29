VERSION 5.00
Object = "{DE8D4E3E-DD62-11D2-821F-444553540001}#1.0#0"; "CHATSCAN³.OCX"
Begin VB.Form frmServerhelp 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KnK 4 Life Server Helper                       By: KnK"
   ClientHeight    =   4260
   ClientLeft      =   2670
   ClientTop       =   945
   ClientWidth     =   8460
   Icon            =   "Serverhelper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnFind 
      Caption         =   "Find"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "#,#,#"
      Enabled         =   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   38
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame frmSetup 
      BackColor       =   &H00000000&
      Caption         =   "New"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   5000
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "Send "
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "Empty"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Empty"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Empty"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "Empty"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Send List"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Send"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Find X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Send Thanx"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Send Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "SN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton btnNarrow 
      Caption         =   "<---------"
      Height          =   255
      Left            =   3480
      TabIndex        =   36
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton btnAddroom 
      Caption         =   "Add Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox lstRoom 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame frameFindbin 
      BackColor       =   &H00000000&
      Caption         =   "Edit Find Bin"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   2655
      Begin VB.CommandButton btnUp 
         Caption         =   "^"
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnSavefindbin 
         Caption         =   "Save"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton btnRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.ComboBox fndBin 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox svrName 
      Height          =   315
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   1335
   End
   Begin chatscan³.Chat chatScan 
      Left            =   120
      Top             =   2520
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin VB.CommandButton btnThrough 
      Caption         =   "#-#"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton btnStatus 
      Caption         =   "Send Status"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton btnList 
      Caption         =   "Send List"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton btmReset 
      Caption         =   "Reset "
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton btnApplysave 
      Caption         =   "Apply/Save "
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame framRunning 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2175
      Left            =   5000
      TabIndex        =   4
      Top             =   0
      Width           =   3480
      Begin VB.PictureBox pctProgress 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   3075
         TabIndex        =   40
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CommandButton btnEmergencystop 
         Caption         =   "Emergency Stop"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lobPwaite 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblRunning 
         BackStyle       =   0  'Transparent
         Caption         =   "Currently Running"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   120
   End
   Begin VB.ListBox lstRequestbin 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   1035
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox lstNumbers 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton btnClearbin 
      Caption         =   "Clear Bin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton btnRequest 
      Caption         =   "Request"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblNotify 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1560
      TabIndex        =   26
      Top             =   50
      Width           =   1815
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu min 
         Caption         =   "Minimize"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Toolz"
      Begin VB.Menu cchat 
         Caption         =   "C-Chat"
      End
      Begin VB.Menu im 
         Caption         =   "IM's"
         Begin VB.Menu on 
            Caption         =   "&On"
         End
         Begin VB.Menu off 
            Caption         =   "&Off"
         End
      End
      Begin VB.Menu bust 
         Caption         =   "&Room Buster"
      End
      Begin VB.Menu fb 
         Caption         =   "Find Bin"
         Begin VB.Menu sm 
            Caption         =   "Setup/Modify"
         End
         Begin VB.Menu findbin2 
            Caption         =   "List Selection"
            Begin VB.Menu bin1 
               Caption         =   "Bin #1"
               Checked         =   -1  'True
               Shortcut        =   {F1}
            End
            Begin VB.Menu bin2 
               Caption         =   "Bin #2"
               Shortcut        =   {F2}
            End
         End
         Begin VB.Menu activate 
            Caption         =   "Activate"
         End
      End
   End
   Begin VB.Menu option 
      Caption         =   "&Options"
      Begin VB.Menu servermethod 
         Caption         =   "&Server M&ethods"
         Begin VB.Menu sandm 
            Caption         =   "Setup/Modify"
         End
         Begin VB.Menu line 
            Caption         =   "-"
         End
         Begin VB.Menu number1 
            Caption         =   "/SN Send x"
            Checked         =   -1  'True
         End
         Begin VB.Menu number4 
            Caption         =   "Empty"
         End
      End
   End
   Begin VB.Menu namess 
      Caption         =   "&Get N&ames"
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu linkhelp 
         Caption         =   "Bill's Help"
      End
      Begin VB.Menu linkbillimage 
         Caption         =   "Bill's images"
      End
      Begin VB.Menu linkfreejava 
         Caption         =   "100% FREE Java"
      End
      Begin VB.Menu linkknk2000ex 
         Caption         =   "KnK 2000 Exc"
      End
      Begin VB.Menu linkknk2000 
         Caption         =   "KnK 2000"
      End
      Begin VB.Menu pline 
         Caption         =   "-"
      End
      Begin VB.Menu max 
         Caption         =   "Maxamize"
      End
   End
End
Attribute VB_Name = "frmServerhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bOn As Boolean
Private vbTray As NOTIFYICONDATA
Private Declare Function GetTickCount Lib "User" () As Long
Private pBar As New CProgressBar

Private Sub TrayIt()
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = Me.hwnd
    vbTray.uId = vbNull
    vbTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    vbTray.ucallbackMessage = WM_MOUSEMOVE
    vbTray.hIcon = Me.Icon
    vbTray.szTip = Me.Caption & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
    Me.Hide
End Sub

Private Sub UnTrayIt()
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = Me.hwnd
    vbTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
End Sub
Private Sub activate_Click()
If svrName = "" Then
MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
Exit Sub
End If

If fndBin.ListCount = 0 Then
MsgBox "error: Nothing to request.", vbCritical, "error"
Exit Sub
End If



frmServerhelp.Height = 2430
Call Setupform

'<-----------
'ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>KnK 4 Life v2.o" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
'Pause (0.8)
'ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>Server Helper" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
'Pause (0.8)

Timer3.Enabled = True
End Sub

Private Sub auto_Click()


Frame1.Top = 4800
Frame1.Left = 1920
Frame2.Top = 4800
Frame2.Left = 4320
Frame6.Top = 0
Frame6.Left = 3480
Form1.Width = 4875

End Sub


Private Sub bin1_Click()
bin1.Checked = True
bin2.Checked = False
fndBin.Clear
Call LoadComboBox(App.Path + "\find.lst", fndBin)
End Sub

Private Sub bin2_Click()
bin1.Checked = False
bin2.Checked = True
fndBin.Clear
Call LoadComboBox(App.Path + "\find2.lst", fndBin)
End Sub

Private Sub btmReset_Click()
Select Case MsgBox("Are you sure you wanna reset method #4?", vbYesNo + vbQuestion + vbDefaultButton2, "Confermation")
Case vbYes
    Text8.Text = ""
    Text4.Text = "Empty"
    Text5.Text = "Empty"
    Text6.Text = "Empty"
    Text7.Text = "Empty"
    Text9.Text = "Empty"
Case vbNo
    Exit Sub
End Select
End Sub

Private Sub btnAdd_Click()
fndBin.AddItem fndBin.Text
End Sub

Private Sub btnAddroom_Click()
lstRoom.Clear
Call AddRoomToListbox(lstRoom, False)
End Sub

Private Sub btnApplysave_Click()

If Text4 = "Empty" Then
    number4.Caption = "Empty"
End If
If Text8 <> "Empty" Then
    number4.Caption = Text8.Text + "SN" + " " + Text4.Text
        R% = WritePrivateProfileString("Number4", "ascii", Text8.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendlist", Text4.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendstatus", Text5.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendthanx", Text6.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "find", Text7.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "send", Text9.Text, App.Path + "\KnK4Life.ini")

End If
End Sub

Private Sub btnClearbin_Click()
lstRequestbin.Clear
End Sub

Private Sub btnEmergencystop_Click()
    Timer2.Enabled = False
    Timer3.Enabled = False
    Call Endsetups
End Sub

Private Sub btnFind_Click()
  If svrName.Text = "" Then
        MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
        Exit Sub
    End If
    
    Select Case True
        Case number1.Checked
            ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">/" + svrName + " Find " + fndBin & "</font>")
        Case number4.Checked
            If Text7.Text = "Empty" Then
                MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
                Exit Sub
            End If
            ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">" & Text8 + svrName + " " + Text7 + " " + fndBin & "</font>")
    End Select

End Sub

Private Sub btnList_Click()
    If svrName.Text = "" Then
        MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
        Exit Sub
    End If
    Select Case True
        Case number1.Checked
            ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">/" + svrName + " Send List</font>")
        Case number4.Checked
            If Text4.Text = "Empty" Then
                MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
                Exit Sub
            End If
            ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">" & Text8 + svrName + " " + Text4 & "</font>")
    End Select
End Sub

Private Sub btnNarrow_Click()
frmServerhelp.Width = 3480
End Sub

Private Sub btnNext_Click()
' For i = 0 To lstNumbers.Selected - 1
'    X = i & ","
'    Pause (2)
    'SmoothProgress1.Value = SmoothProgress1.Value + 1
'    Next i
    
'    ChatSend ("/" + svrName + " Send " + lstNumbers.List(i))
x = lstNumbers

ChatSend ("/" & svrName & " Send " & x & "")
End Sub

Private Sub btnRemove_Click()
fndBin.RemoveItem fndBin.ListIndex
End Sub

Private Sub btnRequest_Click()

If svrName.Text = "" Then
    MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
    Exit Sub
End If

If lstRequestbin.ListCount = 0 Then
    MsgBox "error: Nothing to request.", vbCritical, "error"
    Exit Sub
End If

Call Setupform
'ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>KnK 4 Life v2.o" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
'Pause (0.8)
'ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>Server Helper" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
'Pause (0.8)

Timer2.Enabled = True
End Sub

Private Sub btnSavefindbin_Click()
    If bin1.Checked = True Then
        Call SaveComboBox(App.Path + "\find.lst", fndBin)
    End If
    If bin2.Checked = True Then
        Call SaveComboBox(App.Path + "\find2.lst", fndBin)
    End If
End Sub

Private Sub btnStatus_Click()
    If svrName.Text = "" Then
        MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
        Exit Sub
    End If

    Select Case True
        Case number1.Checked
            ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">/" + svrName + " Send Status</font>")
        Case number4.Checked
            If Text5.Text = "Empty" Then
                MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
                Exit Sub
            End If
            ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">" & Text8 + svrName + " " + Text5 & "</font>")
            
    End Select
End Sub

Private Sub btnThrough_Click()
 For i = 0 To lstNumbers.ListCount - 1
    ChatSend ("/" + svrName + " Send " + lstNumbers.List(i))
    Pause (2)
    'SmoothProgress1.Value = SmoothProgress1.Value + 1
    Next i
    
    
End Sub

Private Sub btnUp_Click()
frmServerhelp.Height = 2430
End Sub

Private Sub bust_Click()
'On Local Error Resume Next
'Call ShellExecute(hwnd, "Open", "/C·RB.EXE", "", App.Path, 1)
'If Err Then
'MsgBox "Error: File Not Found!", vbExclamation, "Error"
'End If
frmRoombust.Show
End Sub

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If What_Said Like "/*" Then
    x$ = What_Said
    T$ = Left$(x$, 2)
    S$ = Right$(x$, 10)
    y$ = Mid$(x$, 2, 10)

If List5.ListCount = 0 Then List5.AddItem y$
    For i = 0 To List5.ListCount - 1
    num = List5.List(i)
If num = y$ Then Exit Sub
    Next i
    List5.AddItem y$
    ' List5.AddItem Y$
End If
End Sub

Private Sub cl_Click()
dcl.Checked = False
cl.Checked = True
SaveSetting "KnK4Life", "KnK_Clearlist", "yesorno", "yes"

End Sub

Private Sub dcl_Click()
dcl.Checked = True
cl.Checked = False
SaveSetting "KnK4Life", "KnK_Clearlist", "yesorno", "no"

End Sub


Private Sub cchat_Click()
frmChat.Show
End Sub

Private Sub exit_Click()
SetWindowLong Me.hwnd, GWL_WNDPROC, lProcOld
End
End Sub

Private Sub exit2_Click()
End
End Sub

Private Sub Form_Load()

FormOnTop Me
chatScan.ScanOn
Set pBar.Canvas = pctProgress
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static lngMsg As Long
    Dim blnFlag As Boolean, lngResult As Long
    lngMsg = x / Screen.TwipsPerPixelX
    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
            Case WM_LBUTTONDBLCLICK
                Me.WindowState = 0
                Me.Show
            Case WM_RBUTTONUP
                lngResult = SetForegroundWindow(Me.hwnd)
                Me.PopupMenu mnuPopUp
        End Select
        blnFlag = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnTrayIt
End Sub

Private Sub chatScan_ChatMsg(Screen_Name As String, What_Said As String)
'X$ = What_Said
User$ = GetUser
User$ = ReplaceString(User$, " ", "")
'X$ = ReplaceString(X$, " ", "")

'If LCase(X$) Like "*" & LCase(User$) & "*" & "list" & "*" Then
 If LCase(ReplaceString(What_Said, " ", "")) Like "*" & LCase(User$) & "*" & "list" & "*" Then
        lblNotify.Caption = "" & Screen_Name & " Sent you a list!"

        Playwav ("C:\Windows\MEDIA\Ding.wav")
    Pause (3)
    lblNotify.Caption = ""
End If


End Sub


Private Sub linkbillimage_Click()
ChatSend "<font face=Arial color=#000000>Link - </font>< a href=http://bia.resource-zone.com/></u>Bill's Image Archive" & "</a>"

End Sub

Private Sub linkfreejava_Click()
ChatSend "<font face=Arial color=#000000>Link - </font>< a href=http://knk.8op.com/java/></u>100% FREE Java" & "</a>"

End Sub

Private Sub linkhelp_Click()
ChatSend "<font face=Arial color=#000000>Link - </font>< a href=http://knk.8op.com/help/></u>KnK Help Files" & "</a>"

End Sub

Private Sub linkknk2000_Click()
ChatSend "<font face=Arial color=#000000>Link - </font>< a href=http://www.knk2000.com/knk/></u>KnK 2000 VB Site" & "</a>"

End Sub

Private Sub linkknk2000ex_Click()
ChatSend "<font face=Arial color=#000000>Link - </font>< a href=http://www.knk2000.com/exchange/></u>KnK 2000 VB Banner Exchange" & "</a>"

End Sub

Private Sub lstNumbers_DblClick()
If lstRequestbin.ListCount = 0 Then lstRequestbin.AddItem lstNumbers
For i = 0 To lstRequestbin.ListCount - 1
num = lstRequestbin.List(i)
If num = lstNumbers Then Exit Sub
Next i
lstRequestbin.AddItem lstNumbers
End Sub



Private Sub lstRequestbin_DblClick()
lstRequestbin.RemoveItem lstRequestbin.ListIndex
End Sub

Private Sub fndBin_DblClick()
fndBin.RemoveItem fndBin.ListIndex
End Sub

Private Sub lstRoom_DblClick()
svrName.Text = lstRoom
svrName.AddItem lstRoom
End Sub

Private Sub max_Click()
 Call UnTrayIt
    Me.WindowState = 0
    Me.Show
End Sub

Private Sub min_Click()
 Call TrayIt
End Sub


Private Sub namess_Click()
frmServerhelp.Width = 4875
End Sub

Private Sub number1_Click()
number1.Checked = True
number4.Checked = False
End Sub


Private Sub number4_Click()
If number4.Caption = "Empty" Then
MsgBox "error: Current method hasnt been setup.", vbExclamation, "error"

Exit Sub
End If
number1.Checked = False
number4.Checked = True


End Sub
Private Sub off_Click()
 Call InstantMessage("$IM_OFF", "=)")
End Sub

Private Sub on_Click()
 Call InstantMessage("$IM_ON", "=)")
End Sub


Private Sub sandm_Click()
Command5.Enabled = False
Frame4.Top = 1800
Frame4.Left = 0
Form1.Height = 4875
End Sub

Private Sub sm_Click()
framRunning.Top = 0
framRunning.Left = 5000
frmSetup.Top = 2040
frmSetup.Left = 5000
frmServerhelp.Height = 3030
End Sub


Private Sub Timer2_Timer()
'<!---- Request Bin -->
    pBar.min = 0
    Z = lstRequestbin.ListCount
    pBar.max = Z
    

Select Case True

Case number1.Checked
    For i = 0 To lstRequestbin.ListCount - 1
    If Timer2.Enabled = False Then
        Exit Sub
    End If
    ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">/" & svrName & " Send " & lstRequestbin.List(i) & "</font>")
    pBar.Value = pBar.Value + 1
    Pause (2)
    Next i
    Timer2.Enabled = False

Case number4.Checked
    For i = 0 To lstRequestbin.ListCount - 1
    If Timer2.Enabled = False Then
        Exit Sub
    End If
    ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">" & Text8 & svrName & " " & Text9 & " " & lstRequestbin.List(i) & "</font>")
     pBar.Value = pBar.Value + 1
   Pause (2)
    Next i
    Timer2.Enabled = False

End Select
Call Endsetups
End Sub

Private Sub Timer3_Timer()
    pBar.min = 0
    If pBar.Value <> 0 Then
    pBar.Value = 0
    End If
    Z = fndBin.ListCount
    
    SmoothProgress1.max = Z
    

Select Case True
Case number1.Checked
    For i = 0 To fndBin.ListCount - 1
    If Timer3.Enabled = False Then
        Exit Sub
    End If
    ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">/" & svrName & " Find " & fndBin.List(i) & "</font>")
    pBar.Value = pBar.Value + 1
    Pause (2)
    Next i
        Timer3.Enabled = False

Case number4.Checked
    For i = 0 To lstRequestbin.ListCount - 1
    If Timer3.Enabled = False Then
        Exit Sub
    End If
    ChatSend ("<font face=" & Chr(34) & "Arial" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & ">" & Text8 & svrName & " " & Text7 & " " & fndBin.List(i) & "</font>")
    pBar.Value = pBar.Value + 1
    Pause (2)
    Next i
    Timer3.Enabled = False


End Select

Call Endsetups
End Sub

