VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSignOn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic AIM"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmSignOn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbScreenName 
      Height          =   315
      ItemData        =   "frmSignOn.frx":1272
      Left            =   1200
      List            =   "frmSignOn.frx":1274
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSignOn 
      Caption         =   "Sign On"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Timer tmrStatus 
      Interval        =   100
      Left            =   720
      Top             =   240
   End
   Begin MSWinsockLib.Winsock wskAIM 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      Height          =   2190
      Left            =   120
      Picture         =   "frmSignOn.frx":1276
      ScaleHeight     =   2130
      ScaleWidth      =   2730
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2790
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   600
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   12
         ImageHeight     =   12
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSignOn.frx":142B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSignOn.frx":1435D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSignOn.frx":143EF
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   1320
      Picture         =   "frmSignOn.frx":1446D
      Top             =   3840
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   1320
      Picture         =   "frmSignOn.frx":14508
      Top             =   3600
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Modifications by Thomas Grimshaw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2340
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Screen Name"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   690
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   615
   End
End
Attribute VB_Name = "frmSignOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nodX As Node

Sub groupadd(groupname)
  Dim lngCounter As Long, strKey As String, nodGroup As Node
  If ExistsInTree(frmBuddyList.tvwSetup, "New Group", True) = False Then
    If frmBuddyList.tvwSetup.SelectedItem Is Nothing Then
      Set nodGroup = frmBuddyList.tvwSetup.Nodes.Add(, , , "New Group", 1, 1)
    Else
      If frmBuddyList.tvwSetup.SelectedItem.parent Is Nothing Then
        Set nodGroup = frmBuddyList.tvwSetup.Nodes.Add(frmBuddyList.tvwSetup.SelectedItem.Index, tvwNext, , "New Group", 1, 1)
      Else
        Set nodGroup = frmBuddyList.tvwSetup.Nodes.Add(frmBuddyList.tvwSetup.SelectedItem.parent.Index, tvwNext, , "New Group", 1, 1)
      End If
    End If
    nodGroup.Selected = True
    frmBuddyList.tvwSetup.SetFocus
    frmBuddyList.tvwSetup.SelectedItem.Text = UCase$(Left$(Trim(groupname), 1)) + LCase$(Right$(Trim(groupname), Len(Trim(groupname)) - 1))
    
  End If
  
  
  If ExistsInTree(frmBuddyList.tvwBuddies, "New Group", True) = False Then
    If frmBuddyList.tvwBuddies.SelectedItem Is Nothing Then
      Set nodGroup = frmBuddyList.tvwBuddies.Nodes.Add(, , , "New Group", 1, 1)
    Else
      If frmBuddyList.tvwBuddies.SelectedItem.parent Is Nothing Then
        Set nodGroup = frmBuddyList.tvwBuddies.Nodes.Add(frmBuddyList.tvwBuddies.SelectedItem.Index, tvwNext, , "New Group", 1, 1)
      Else
        Set nodGroup = frmBuddyList.tvwBuddies.Nodes.Add(frmBuddyList.tvwBuddies.SelectedItem.parent.Index, tvwNext, , "New Group", 1, 1)
      End If
    End If
    nodGroup.Selected = True
    frmBuddyList.tvwBuddies.SetFocus
    frmBuddyList.tvwBuddies.SelectedItem.Text = UCase$(Left$(Trim(groupname), 1)) + LCase$(Right$(Trim(groupname), Len(Trim(groupname)) - 1))
    
  End If
End Sub
Sub buddyadd(buddy, groupname)
'SendProc "toc_add_permit " + buddy
Call SendProc(2, "toc_add_buddy " & buddy & Chr(0))
  Dim nodBuddy As Node
  
  If ExistsInTree(frmBuddyList.tvwSetup, "New Buddy", True) = False Then
    If frmBuddyList.tvwSetup.Nodes.Count < 1 Then
      MsgBox "You need a group to add buddies to.", vbOKOnly + vbCritical, "Error"
      Exit Sub
    End If
    If frmBuddyList.tvwSetup.SelectedItem Is Nothing Then
      Set nodBuddy = frmBuddyList.tvwSetup.Nodes.Add(1, tvwChild, , "New Buddy", 3, 3)
    Else
      If frmBuddyList.tvwSetup.SelectedItem.parent Is Nothing Then
        Set nodBuddy = frmBuddyList.tvwSetup.Nodes.Add(frmBuddyList.tvwSetup.SelectedItem.Index, tvwChild, , "New Buddy", 3, 3)
      Else
        Set nodBuddy = frmBuddyList.tvwSetup.Nodes.Add(frmBuddyList.tvwSetup.SelectedItem.Index, tvwPrevious, , "New Buddy", 3, 3)
      End If
    End If
    nodBuddy.Selected = True
    frmBuddyList.tvwSetup.SetFocus
    frmBuddyList.tvwSetup.SelectedItem.Text = Trim(buddy)
  End If


End Sub
Private Sub cmbScreenName_Click()
  txtPassword.Text = ""
End Sub

Private Sub cmbScreenName_GotFocus()
  cmbScreenName.SelStart = 0
  cmbScreenName.SelLength = Len(cmbScreenName.Text)
End Sub

Private Sub cmbScreenName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSignOn_Click
  End If
End Sub

Private Sub cmdAbout_Click()
  frmAbout.Show
End Sub

Private Sub cmdSignOn_Click()
  Dim lngDo As Long, blnFound As Boolean
  If cmbScreenName.Text = "" Then
    MsgBox "You must enter a screen name.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  If txtPassword.Text = "" Then
    MsgBox "You must enter a password.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  m_strScreenName$ = LCase(Replace(cmbScreenName.Text, " ", ""))
  m_strPassword$ = EncryptPW(txtPassword.Text)
  Randomize
  m_lngLocalSeq& = Int((65535 * Rnd) + 1)
  wskAIM.Close
  wskAIM.RemoteHost = "toc.oscar.aol.com"
  wskAIM.RemotePort = 5190
  wskAIM.Connect
  For lngDo& = 0 To cmbScreenName.ListCount - 1
    If cmbScreenName.List(lngDo&) = m_strScreenName$ Then
      blnFound = True
      Exit For
    End If
  Next
  If blnFound = False Then cmbScreenName.AddItem cmbScreenName.Text
End Sub

Private Sub Command1_Click()
frmIM.Show
End Sub

Private Sub Form_Load()
  'we load our screen names from the ini and the last used screen name.
  Dim arrNames() As String, strCurrent As String, lngDo As Long
  arrNames$ = Split(GetINIString("settings", "names", App.Path & "\aim.ini"), " ")
  strCurrent$ = GetINIString("settings", "current", App.Path & "\aim.ini")
  For lngDo& = LBound(arrNames$) To UBound(arrNames$)
    cmbScreenName.AddItem arrNames$(lngDo&)
  Next
  cmbScreenName.Text = strCurrent$
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'we save our screen names as well as our last used screen name.
  Dim lngDo As Long, strNames As String
  For lngDo& = 0 To cmbScreenName.ListCount - 1
    strNames$ = strNames$ & cmbScreenName.List(lngDo&) & " "
  Next
  Call WriteINIString("settings", "names", Trim(strNames$), App.Path & "\aim.ini")
  Call WriteINIString("settings", "current", LCase(Replace(cmbScreenName.Text, " ", "")), App.Path & "\aim.ini")
  End
End Sub

Private Sub tmrStatus_Timer()
  Dim strState As String
  Select Case wskAIM.State
    Case 0
      strState$ = "0 - Closed"
    Case 1
      strState$ = "1 - Open"
    Case 2
      strState$ = "2 - Listening"
    Case 3
      strState$ = "3 - Connection pending"
    Case 4
      strState$ = "4 - Resolving host"
    Case 5
      strState$ = "5 - Host resolved"
    Case 6
      strState$ = "6 - Connecting"
    Case 7
      strState$ = "7 - Connected"
    Case 8
      strState$ = "8 - Peer closing"
    Case 9
      strState$ = "9 - Error"
  End Select
  If lblStatus.Caption <> "Status: " & strState$ Then
    lblStatus.Caption = "Status: " & strState$
  End If
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSignOn_Click
  End If
End Sub

Private Sub wskAIM_Connect()
  'the FLAPON is our first message sent to the aim toc server after a connection is made.
  'from here we will a flap response containing the flap version.
  If wskAIM.State = sckConnected Then
    wskAIM.SendData "FLAPON" & vbCrLf & vbCrLf
    Call DoDebug("OUT: FLAPON")
  End If
End Sub

Private Sub wskAIM_DataArrival(ByVal bytesTotal As Long)
  'this procedure is where all the data is handled. it is important for us to pay attention
  'to the flap headers since more than one command may be sent per packet. the payload in
  'the flap header is very important here. it allows us to know how much data is in that
  'command.
  Dim strData As String, lngMark As Long, lngDataLen As Long
  Dim lngFrameType As Long, lngSeqA As Long, lngSeqB As Long
  Dim lngPayLo As Long, lngPayHi As Long, lngPayload As Long
  Dim strCommand As String
  wskAIM.GetData strData$, vbString
  Debug.Print strData$
  
  lngDataLen& = Len(strData$)
  lngMark& = 1
  Do While lngMark& < lngDataLen&
    lngFrameType& = Asc(Mid(strData$, lngMark& + 1))
    lngSeqA& = Asc(Mid(strData$, lngMark& + 2))
    lngSeqB& = Asc(Mid(strData$, lngMark& + 3))
    lngPayLo& = Asc(Mid(strData$, lngMark& + 4))
    lngPayHi& = Asc(Mid(strData$, lngMark& + 5))
    lngPayload& = MakeLong(lngPayHi&, lngPayLo&)
    strCommand$ = Mid(strData$, lngMark& + 6, lngPayload&)
    Call DoDebug("IN:  42 " & lngFrameType& & " " & lngSeqA& & " " & lngSeqB& & " " & lngPayHi& & " " & lngPayLo& & " " & strCommand$)
    'you'll notice that i am not outputing the incoming data to the debug window as it is
    'received. this is because of the null characters. they normally cancel out anything
    'after them when added to a text box. for this reason, i am outputing the asc values
    'of the flap header rather than the actually characters they are received as. also
    ' with the incoming data, i am replacing the null characters with "/0" in order to
    'maintain a readable format.
    Call HandleProc(lngFrameType&, strCommand$)
    'here we have seperated a command from the incoming data. we will call the handlproc
    'procedure for each command found in the incoming data.
    lngMark& = lngMark& + lngPayload& + 6
  Loop
End Sub

Private Sub wskAIM_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call DoDebug("ERROR: " & Number & " " & Description)
End Sub

Private Sub HandleProc(lngFrameType As Long, strData As String)
  Dim strSNFiller As String, arrCommand() As String, arrArgs() As String
  Dim x As Integer
  Dim arse$
  
  Dim lngFormIndex As Long, arrNames() As String, blnShowJoin As Boolean
  Dim lngNameLoop As Long, lngListLoop As Long, lngTreeLoop As Long
  Dim strName As String, strParent As String, nodBuddy As Node
  Select Case lngFrameType&
    Case 1 'a frame type of "1" indicates this message is part of the signon sequence
      If strData$ = Chr(0) & Chr(0) & Chr(0) & Chr(1) Then
        strSNFiller$ = String(Len(CStr(Len(m_strScreenName$))), "0")
        Call SendProc(1, Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(1) & Chr(CLng(strSNFiller$)) & Chr(CLng(CStr(Len(m_strScreenName$)))) & m_strScreenName$)
        'here we send our flap version, tvl tag, normalized screen name length, and
        'our normalized screen name
        Call SendProc(2, "toc_signon login.oscar.aol.com " + Format$(Int(Rnd(9 * Rnd + 1))) + "234  " & m_strScreenName$ & " " & m_strPassword$ & " english " & Chr(34) & "AOLInstantMessenger" & Chr(34) & Chr(0))
        'now we send our signon message to start the signon process.
      End If
    Case 2 'a frame type of "2" indicates normal data
      arrCommand$ = Split(strData$, ":", 2)
      'If Left$(arrCommand$(0), 2) = "b " Then
      
      'End If
    
      Select Case UCase(arrCommand$(0))
        Case "CONFIG"
        'MsgBox UCase$(arrCommand$(1))
        
        For x = 1 To Len(arrCommand$(1))
      If Mid$(arrCommand$(1), x, 1) = Chr$(10) Then
      If Left$(arse$, 1) = "g" Then
      groupadd Right$(arse$, Len(arse$) - 1)
      lastgroup = Right$(arse$, Len(arse$) - 1)
      End If
      If Left$(arse$, 1) = "b" Then
      buddyadd Right$(arse$, Len(arse$) - 1), lastgroup
      End If
            
      arse$ = ""
      Else
      arse$ = arse$ + Mid$(arrCommand$(1), x, 1)
      End If
     
      Next x
       If Left$(arse$, 1) = "g" Then
      groupadd Right$(arse$, Len(arse$) - 1)
      lastgroup = Right$(arse$, Len(arse$) - 1)
      End If
      If Left$(arse$, 1) = "b" Then
      buddyadd Right$(arse$, Len(arse$) - 1), lastgroup
      End If
      arse$ = ""
      'Call SendProc(2, "toc_init_done" & Chr(0))
        'Debug.Print "Params: " + UBound(arrCommand$)
        Call SendProc(2, "toc_init_done" & Chr(0))
        Case "CHAT_IN"
          'incoming chat room text
          'argument 1: chat room id
          'argument 2: sender's screen name
          'argument 3: whisper t/f (not handled here)
          'argument 4: message
          arrArgs$() = Split(arrCommand$(1), ":", 4)
          lngFormIndex& = FormByTag(arrArgs$(0))
          If lngFormIndex& > -1 Then
            If arrArgs$(1) = m_strFormattedSN$ Then
            'here we update in red if the message is our's and blue if not.
              Call RTFUpdate(Forms(lngFormIndex&).rtfDisplay, "\par\plain\fs16\cf2\b " & arrArgs$(1) & ": \plain\fs16\cf0 " & FixRTF(KillHTML(arrArgs$(3))))
            Else
              Call RTFUpdate(Forms(lngFormIndex&).rtfDisplay, "\par\plain\fs16\cf1\b " & arrArgs$(1) & ": \plain\fs16\cf0 " & FixRTF(KillHTML(arrArgs$(3))))
            End If
          End If
        Case "CHAT_INVITE"
          'incoming invitation to a chat room
          'argument 1: chat room name
          'argument 2: chat room id
          'argument 3: invite sender
          'argument 4: invitation message
          arrArgs$() = Split(arrCommand$(1), ":", 4)
          Dim frmNewInvitation As New frmInvitation
          With frmNewInvitation
            .Caption = arrArgs$(0)
            .Tag = "j" & arrArgs$(1)
            .lblInfo.Caption = arrArgs$(2) & " has invited you to join " & Chr(34) & arrArgs$(0) & Chr(34) & "." & vbCrLf & vbCrLf & arrArgs$(3)
            .Show
          End With
        Case "CHAT_JOIN"
          'indicates that our attempt to join a chat room was successful
          'argument 1: chat room id
          'argument 2: chat room name
          arrArgs$() = Split(arrCommand$(1), ":", 2)
          lngFormIndex& = FormByTag(arrArgs$(0))
          If lngFormIndex& < 0 Then
            Dim frmNewChat As New frmChatRoom
            With frmNewChat
              .Caption = arrArgs$(1)
              .Tag = arrArgs$(0)
              .Show
            End With
            Call RTFUpdate(frmNewChat.rtfDisplay, "\par\plain\fs16\cf3\b *** You have joined " & Chr(34) & arrArgs$(1) & Chr(34))
          Else
            Call RTFUpdate(Forms(lngFormIndex&).rtfDisplay, "\par\plain\fs16\cf3\b *** You have joined " & Chr(34) & arrArgs$(1) & Chr(34))
          End If
          If strInviteRoom$ <> "" Then
            Call SendProc(2, "toc_chat_invite " & arrArgs$(0) & " " & Chr(34) & strInviteMessage$ & Chr(34) & " " & strInviteBuddies$ & Chr(0))
            strInviteRoom$ = ""
          End If
        Case "CHAT_UPDATE_BUDDY"
          'indicates that a user has either joined or parted a chat room
          'argument 1: chat room id
          'argument 2: joined t/f
          'argument 3: list of users joining or parting the room
          arrArgs$() = Split(arrCommand$(1), ":", 3)
          arrNames$() = Split(arrArgs$(2), ":")
          lngFormIndex& = FormByTag(arrArgs$(0))
          If lngFormIndex& > -1 Then
            If arrArgs$(1) = "T" Then
              If LCase(Replace(arrNames$(0), " ", "")) <> m_strScreenName$ Then
                blnShowJoin = True
              End If
              For lngNameLoop& = LBound(arrNames$) To UBound(arrNames$)
                Forms(lngFormIndex&).lstNames.AddItem arrNames$(lngNameLoop&)
                Forms(lngFormIndex&).lblPeople.Caption = Forms(lngFormIndex&).lstNames.ListCount & " people here"
                If blnShowJoin = True Then
                  Call RTFUpdate(Forms(lngFormIndex&).rtfDisplay, "\par\plain\fs16\cf3\b *** " & arrNames$(lngNameLoop&) & " has joined " & Forms(lngFormIndex&).Caption)
                End If
              Next
            Else
              For lngNameLoop& = LBound(arrNames$) To UBound(arrNames$)
                For lngListLoop& = 0 To Forms(lngFormIndex&).lstNames.ListCount - 1
                  If Forms(lngFormIndex&).lstNames.List(lngListLoop&) = arrNames$(lngNameLoop&) Then
                    Forms(lngFormIndex&).lstNames.RemoveItem lngListLoop&
                    Forms(lngFormIndex&).lblPeople.Caption = Forms(lngFormIndex&).lstNames.ListCount & " people here"
                    Call RTFUpdate(Forms(lngFormIndex&).rtfDisplay, "\par\plain\fs16\cf4\b *** " & arrNames$(lngNameLoop&) & " has left " & Forms(lngFormIndex&).Caption)
                    Exit For
                  End If
                Next
              Next
            End If
          End If
        Case "ERROR"
          'indicates there was an error
          'argument 1: error id number
          'argument 2: varies depending on the error
          arrArgs$() = Split(arrCommand$(1), ":", 2)
          Dim frmNewError As New frmError
          Select Case arrArgs$(0)
            Case "901"
              frmNewError.lblErrorType.Caption = "General Error: 901"
              frmNewError.lblInfo.Caption = arrArgs$(1) & " is not currently available."
            Case "902"
              frmNewError.lblErrorType.Caption = "General Error: 902"
              frmNewError.lblInfo.Caption = "Warning of " & arrArgs$(1) & " is not currently available."
            Case "903"
              frmNewError.lblErrorType.Caption = "General Error: 903"
              frmNewError.lblInfo.Caption = "A message has been dropped, you are exceeding the server speed limit."
            Case "950"
              frmNewError.lblErrorType.Caption = "Chat Error: 950"
              frmNewError.lblInfo.Caption = "Chat in " & arrArgs$(1) & " is unavailable."
            Case "960"
              frmNewError.lblErrorType.Caption = "IM/Info Error: 960"
              frmNewError.lblInfo.Caption = "You are sending messages too fast to " & arrArgs$(1) & "."
            Case "961"
              frmNewError.lblErrorType.Caption = "IM/Info Error: 961"
              frmNewError.lblInfo.Caption = "You missed a a message from " & arrArgs$(1) & " because it was too big."
            Case "962"
              frmNewError.lblErrorType.Caption = "IM/Info Error: 962"
              frmNewError.lblInfo.Caption = "You missed a a message from " & arrArgs$(1) & " because it was sent too fast."
            Case "970"
              frmNewError.lblErrorType.Caption = "Directory Error: 970"
              frmNewError.lblInfo.Caption = "Failure."
            Case "971"
              frmNewError.lblErrorType.Caption = "Directory Error: 971"
              frmNewError.lblInfo.Caption = "Too many matches."
            Case "972"
              frmNewError.lblErrorType.Caption = "Directory Error: 972"
              frmNewError.lblInfo.Caption = "Need more qualifiers."
            Case "973"
              frmNewError.lblErrorType.Caption = "Directory Error: 973"
              frmNewError.lblInfo.Caption = "Dir service temporarily unavailable."
            Case "974"
              frmNewError.lblErrorType.Caption = "Directory Error: 974"
              frmNewError.lblInfo.Caption = "Email lookup restricted."
            Case "975"
              frmNewError.lblErrorType.Caption = "Directory Error: 975"
              frmNewError.lblInfo.Caption = "Keyword Ignored."
            Case "976"
              frmNewError.lblErrorType.Caption = "Directory Error: 976"
              frmNewError.lblInfo.Caption = "No Keywords."
            Case "977"
              frmNewError.lblErrorType.Caption = "Directory Error: 977"
              frmNewError.lblInfo.Caption = "Language not supported."
            Case "978"
              frmNewError.lblErrorType.Caption = "Directory Error: 978"
              frmNewError.lblInfo.Caption = "Country not supported."
            Case "979"
              frmNewError.lblErrorType.Caption = "Directory Error: 979"
              frmNewError.lblInfo.Caption = "Failure unknown " & arrArgs$(1) & "."
            Case "980"
              frmNewError.lblErrorType.Caption = "Authorization Error: 980"
              frmNewError.lblInfo.Caption = "Incorrect nickname or password."
            Case "981"
              frmNewError.lblErrorType.Caption = "Authorization Error: 981"
              frmNewError.lblInfo.Caption = "The service is temporarily unavailable."
            Case "982"
              frmNewError.lblErrorType.Caption = "Authorization Error: 982"
              frmNewError.lblInfo.Caption = "Your warning level is currently too high to sign on."
            Case "983"
              frmNewError.lblErrorType.Caption = "Authorization Error: 983"
              frmNewError.lblInfo.Caption = "You have been connecting and disconnecting too frequently." & vbCrLf & "Wait 10 minutes and try again." & _
                                            vbCrLf & "If you continue to try, you will need to wait even longer."
            Case "989"
              frmNewError.lblErrorType.Caption = "Authorization Error: 989"
              frmNewError.lblInfo.Caption = "An unknown signon error has occurred " & arrArgs$(1) & "."
          End Select
          frmNewError.Show
        Case "IM_IN"
          'indicates an incoming instant message
          'argument 1: sender's screen name
          'argument 2: auto resonse t/f (not handled here)
          'argument 3: message
          arrArgs$() = Split(arrCommand$(1), ":", 3)
          
          If MeAway = True Then
          'frmAway.List1.AddItem arrArgs$(0)
          ListAdd arrArgs$(0)
          End If
          
          lngFormIndex& = FormByTag(LCase(Replace(arrArgs$(0), " ", "")))
          If lngFormIndex& > -1 Then
            Call RTFUpdate(Forms(lngFormIndex&).rtfDisplay, "\par\plain\fs16\cf1\b " & arrArgs$(0) & ": \plain\fs16\cf0 " & FixRTF(KillHTML(arrArgs$(2))))
            Call PlayWav(strSoundIMIn$)
          Else
            Dim frmNewIM As New frmIM
            With frmNewIM
              .rtfDisplay.SelStart = Len(.rtfDisplay.Text)
              .Caption = arrArgs$(0) & " - Instant Message"
              .Tag = LCase(Replace(arrArgs$(0), " ", ""))
              If MeAway = True Then
                .Hide
              Else
                .Show
              End If
            End With
            Call RTFUpdate(frmNewIM.rtfDisplay, "\par\plain\fs16\cf1\b " & arrArgs$(0) & ": \plain\fs16\cf0 " & FixRTF(KillHTML(arrArgs$(2))))
            Call PlayWav(strSoundFirstIM$)
          End If
        Case "NICK"
          'this sends us the format our screen name should be used for display
          'argument 1: formatted screen name
          m_strFormattedSN$ = arrCommand$(1)
        Case "SIGN_ON"
          'the sign on message is sent letting us know is is ok to send our configuriations.
          frmBuddyList.Show
          'Call SendProc(2, "toc_set_config " & BuddyConfig$ & Chr(0))
          'send our buddy list
          'Select Case m_strMode$
            'Case "3"
              'Call SendProc(2, "toc_add_permit " & m_strPDList$ & Chr(0))
            
            'Case "4"
              'Call SendProc(2, "toc_add_deny " & m_strPDList$ & Chr(0))
            'Case "5"
              'Call SendProc(2, "toc_add_permit " & GetBuddies$ & Chr(0))
          'End Select
          'send our permit/deny lists
'          Call SendProc(2, "toc_init_done" & Chr(0))
          'end our configurations. it is important to send our configurations before we
          'send toc_init_done so we do not flash on the buddy lists of users we have
          'blocked
          'Call SendProc(2, "toc_add_buddy " & GetBuddies$ & Chr(0))
          'send a list of our buddies to the server so we can receive the UPDATE_BUDDY
          'messages.
          frmSignOn.Hide
          Call DoDebug("RECIEVED: SIGN_ON")
        Case "UPDATE_BUDDY"
          'indicates an update in the online status of one of our buddies
          'argument 1: buddies screen name
          'argument 2: online status t/f
          'argument 3: evil amount (not handled here)
          'argument 4: sign on time (not handled here)
          'argument 5: idle time (not handled here)
          'argument 6: user class (not handled here)
          arrArgs$() = Split(arrCommand$(1), ":", 6)
          strName$ = LCase(Replace(arrArgs$(0), " ", ""))
          
            If arrArgs$(5) = " UU" Then
                For Each nodX In frmBuddyList.tvwBuddies.Nodes
                  If nodX.Text = arrArgs$(0) Then
                    nodX.Image = "Away"
                  End If
                Next
            ElseIf arrArgs$(5) = " U" Then
                For Each nodX In frmBuddyList.tvwBuddies.Nodes
                  If nodX.Text = arrArgs$(0) Then
                    nodX.Image = "Online"
                  End If
                Next
            End If
          
          If arrArgs$(1) = "T" Then
            If ExistsInTree(frmBuddyList.tvwBuddies, arrArgs$(0)) = False Then
              For lngTreeLoop& = 1 To frmBuddyList.tvwSetup.Nodes.Count
                If LCase(Replace(frmBuddyList.tvwSetup.Nodes.Item(lngTreeLoop&).Text, " ", "")) = strName$ Then
                  If Not frmBuddyList.tvwSetup.Nodes.Item(lngTreeLoop&).parent Is Nothing Then
                    strParent$ = frmBuddyList.tvwSetup.Nodes.Item(lngTreeLoop&).parent.Text
                  End If
                  Exit For
                End If
              Next
            End If
            For lngTreeLoop& = 1 To frmBuddyList.tvwBuddies.Nodes.Count
              If frmBuddyList.tvwBuddies.Nodes.Item(lngTreeLoop&).parent Is Nothing Then
                If LCase(Replace(frmBuddyList.tvwBuddies.Nodes.Item(lngTreeLoop&).Text, " ", "")) = LCase(Replace(strParent$, " ", "")) Then
                  Set nodBuddy = frmBuddyList.tvwBuddies.Nodes.Add(lngTreeLoop&, tvwChild, , arrArgs$(0), 3, 3)
                  nodBuddy.EnsureVisible
                  Call PlayWav(strSoundSignOn$)
                End If
              End If
            Next
          Else
            Call ExistsInTree(frmBuddyList.tvwBuddies, strName$, False, True)
            Call PlayWav(strSoundSignOff$)
          End If
      End Select
    Case Else
      Call DoDebug("Invalid Frame: " & lngFrameType&)
  End Select
End Sub

