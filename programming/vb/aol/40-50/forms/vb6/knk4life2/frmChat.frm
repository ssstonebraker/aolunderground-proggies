VERSION 5.00
Object = "{DE8D4E3E-DD62-11D2-821F-444553540001}#1.0#0"; "CHATSCAN³.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KnK 4 Life MChat"
   ClientHeight    =   2580
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   8925
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7440
      Top             =   960
   End
   Begin VB.ListBox lstRoom 
      Height          =   1425
      Left            =   7200
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   2280
      Width           =   615
   End
   Begin VB.ComboBox cmbChat 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   6450
   End
   Begin chatscan³.Chat Chat1 
      Left            =   1800
      Top             =   0
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":030A
   End
   Begin VB.Label Label2 
      Caption         =   "People"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu save 
         Caption         =   "Save Chat"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu min 
         Caption         =   "Minimize"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu max 
         Caption         =   "Maxamize"
      End
   End
   Begin VB.Menu Sites 
      Caption         =   "&Sites"
      Begin VB.Menu site1 
         Caption         =   "KnK VB Site"
      End
      Begin VB.Menu site2 
         Caption         =   "KnK VB 3 site"
      End
      Begin VB.Menu site3 
         Caption         =   "KnK 2000 Banner Exchange"
      End
      Begin VB.Menu site4 
         Caption         =   "KnK Help Files"
      End
      Begin VB.Menu site5 
         Caption         =   "100% FREE Java"
      End
      Begin VB.Menu site6 
         Caption         =   "Bill's Image Archive"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bOn As Boolean
Private vbTray As NOTIFYICONDATA
Private Declare Function GetTickCount Lib "User" () As Long

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

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If Screen_Name = GetUser Then
Call DoChatStuff2(Screen_Name, What_Said, False)
End If
If Screen_Name <> GetUser Then
Call DoChatStuff(Screen_Name, What_Said, False)
End If

End Sub

Private Sub cmbChat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbChat.Text = "" Then
    Exit Sub
End If
If cmbChat.Text <> "" Then
cmbChat.AddItem cmbChat
    ChatSend cmbChat
    cmbChat.Text = ""
End If
End If

End Sub


Private Sub Command1_Click()
If cmbChat.Text = "" Then
    Exit Sub
End If
If cmbChat.Text <> "" Then
cmbChat.AddItem cmbChat
    ChatSend cmbChat
    cmbChat.Text = ""
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
MsgBox lstRoom.ListCount
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
'frmChat.Width = 7217
'frmChat.Height = 2250
Call DoChatStuff("OnlineHost", "Welcome to KnK 4 Life MChat", False)
Call AddRoomToListbox(lstRoom, False)
Label1.Caption = lstRoom.ListCount + 1
Chat1.ScanOn
FormOnTop Me

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

Private Sub lstRoom_Click()
cmbChat.Text = lstRoom
End Sub

Private Sub lstRoom_DblClick()
Call ChatIgnoreByName(lstRoom)
cmbChat.Text = lstRoom & " [ x ] "
End Sub

Private Sub max_Click()
 Call UnTrayIt
    Me.WindowState = 0
    Me.Show
End Sub

Private Sub min_Click()
 Call TrayIt
End Sub

Private Sub save_Click()
Open "C:\Windows\Desktop\chat.rtf" For Output As #1
   Print #1, txtChat.Text
   Close #1
End Sub

Private Sub test1_Click()
cmbChat.Text = "Testing lalalalalala"
'Command1.SetFocus
Command1_Click

End Sub

Private Sub site1_Click()
cmbChat.Text = "<font face=Arial color=#000000>Link - </font>< a href=http://www.knk2000.com/knk/></u>KnK 2000 VB Site" & "</a>"
Command1_Click
End Sub

Private Sub site2_Click()
cmbChat.Text = "<font face=Arial color=#000000>Link - </font>< a href=http://www.knk2000.com/knk3/></u>KnK 3 Site" & "</a>"
Command1_Click
End Sub

Private Sub site3_Click()
cmbChat.Text = "<font face=Arial color=#000000>Link - </font>< a href=http://www.knk2000.com/exchange/></u>KnK 2000 VB Banner Exchange" & "</a>"
Command1_Click
End Sub

Private Sub site4_Click()
cmbChat.Text = "<font face=Arial color=#000000>Link - </font>< a href=http://knk.8op.com/help/></u>KnK Help Files" & "</a>"
Command1_Click
End Sub

Private Sub site5_Click()
cmbChat.Text = "<font face=Arial color=#000000>Link - </font>< a href=http://knk.8op.com/java/></u>100% FREE Java" & "</a>"
Command1_Click
End Sub

Private Sub site6_Click()
cmbChat.Text = "<font face=Arial color=#000000>Link - </font>< a href=http://bia.resource-zone.com/></u>Bill's Image Archive" & "</a>"
Command1_Click
End Sub

Private Sub Timer1_Timer()
Dim AOLFrame25 As Long, MDIClient As Long, AOLChild As Long, AOLStatic As Long
AOLFrame25& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame25&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Dim TheText As String, TL As Long
TL& = SendMessageLong(AOLStatic&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String$(TL& + 1, " ")
Call SendMessageByString(AOLStatic&, WM_GETTEXT, TL + 1, TheText$)
TheText$ = Left(TheText$, TL&)
x = lstRoom.ListCount + 1
If TheText$ = x Then
Exit Sub
End If
If TheText$ <> x Then
lstRoom.Clear
Call AddRoomToListbox(lstRoom, False)
Label1.Caption = lstRoom.ListCount + 1
End If
End Sub
