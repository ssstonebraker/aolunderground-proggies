VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBuddyList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buddy List"
   ClientHeight    =   5655
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2805
   Icon            =   "frmBuddyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgList 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imgBuddy 
      Left            =   1440
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":1272
            Key             =   "Collapsed"
            Object.Tag             =   "Collapsed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":13CC
            Key             =   "Expanded"
            Object.Tag             =   "Expanded"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":1526
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":27A8
            Key             =   "Away"
            Object.Tag             =   "Away"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":283B
            Key             =   "Idle"
            Object.Tag             =   "Idle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddyList.frx":28BF
            Key             =   "Online"
            Object.Tag             =   "Online"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabBuddy 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Online"
      TabPicture(0)   =   "frmBuddyList.frx":2B20
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvwBuddies2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdIM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdInvite"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdChat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tvwBuddies"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "List Setup"
      TabPicture(1)   =   "frmBuddyList.frx":2B3C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Timer1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdDelete"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAddGroup"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAddBuddy"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "tvwSetup"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin MSComctlLib.TreeView tvwBuddies 
         Height          =   3975
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7011
         _Version        =   393217
         Indentation     =   5
         Style           =   1
         ImageList       =   "imgBuddy"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   1215
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmBuddyList.frx":2B58
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -74760
         Top             =   2160
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "delete"
         Height          =   375
         Left            =   -73440
         TabIndex        =   4
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton cmdAddGroup 
         Caption         =   "+group"
         Height          =   375
         Left            =   -74160
         TabIndex        =   5
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton cmdChat 
         Height          =   495
         Left            =   1920
         Picture         =   "frmBuddyList.frx":2B5E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Join A Chat Room"
         Top             =   4800
         Width           =   735
      End
      Begin VB.CommandButton cmdInvite 
         Height          =   495
         Left            =   960
         Picture         =   "frmBuddyList.frx":3DD0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Send Chat Invitation"
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdAddBuddy 
         Caption         =   "+buddy"
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton cmdIM 
         Height          =   495
         Left            =   120
         Picture         =   "frmBuddyList.frx":3E80
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Send Instant Message"
         Top             =   4800
         Width           =   735
      End
      Begin MSComctlLib.TreeView tvwSetup 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7011
         _Version        =   393217
         Indentation     =   88
         LabelEdit       =   1
         Style           =   1
         ImageList       =   "imgBuddy"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwBuddies2 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   7011
         _Version        =   393217
         Indentation     =   88
         Style           =   7
         ImageList       =   "imgBuddy"
         Appearance      =   1
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuFileAway 
         Caption         =   "&Away"
      End
      Begin VB.Menu mnuFileSignOff 
         Caption         =   "&Sign Off"
      End
   End
   Begin VB.Menu mnuPeople 
      Caption         =   "&People"
      Begin VB.Menu mnuPeopleJoinChat 
         Caption         =   "&Join Chat"
      End
      Begin VB.Menu mnuPeopleGetInfo 
         Caption         =   "Get &Info"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Begin VB.Menu mnuDebugShow 
         Caption         =   "S&how Debug Window"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmBuddyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nodX As Node

Private Sub cmdAddBuddy_Click()
  Dim nodBuddy As Node
  If ExistsInTree(tvwSetup, "New Buddy", True) = False Then
    If tvwSetup.Nodes.Count < 1 Then
      MsgBox "You need a group to add buddies to.", vbOKOnly + vbCritical, "Error"
      Exit Sub
    End If
    If tvwSetup.SelectedItem Is Nothing Then
      Set nodBuddy = tvwSetup.Nodes.Add(1, tvwChild, , "New Buddy", 0, 0)
    Else
      If tvwSetup.SelectedItem.parent Is Nothing Then
        Set nodBuddy = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.Index, tvwChild, , "New Buddy", 0, 0)
      Else
        Set nodBuddy = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.Index, tvwPrevious, , "New Buddy", 0, 0)
      End If
    End If
    nodBuddy.Selected = True
    tvwSetup.SetFocus
    tvwSetup.StartLabelEdit
  End If
End Sub

Private Sub cmdAddGroup_Click()
  Dim lngCounter As Long, strKey As String, nodGroup As Node
  If ExistsInTree(tvwSetup, "New Group", True) = False Then
    If tvwSetup.SelectedItem Is Nothing Then
      Set nodGroup = tvwSetup.Nodes.Add(, , , "New Group", 1, 1)
    Else
      If tvwSetup.SelectedItem.parent Is Nothing Then
        Set nodGroup = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.Index, tvwNext, , "New Group", 1, 1)
      Else
        Set nodGroup = tvwSetup.Nodes.Add(tvwSetup.SelectedItem.parent.Index, tvwNext, , "New Group", 1, 1)
      End If
    End If
    nodGroup.Selected = True
    tvwSetup.SetFocus
    tvwSetup.StartLabelEdit
  End If
End Sub

Private Sub cmdDelete_Click()
  Dim lngDo As Long
  If tvwSetup.SelectedItem Is Nothing Then Exit Sub
  If Not tvwSetup.SelectedItem.parent Is Nothing Then
    Call SendProc(2, "toc_remove_buddy " & Chr(34) & Replace(tvwSetup.SelectedItem.Text, " ", "") & Chr(34) & Chr(0))
  Else
    If tvwSetup.SelectedItem.Children > 0 Then
      For lngDo& = 1 To tvwSetup.SelectedItem.Children
        Call SendProc(2, "toc_remove_buddy " & Chr(34) & Replace(tvwSetup.SelectedItem.Child.FirstSibling, " ", "") & Chr(34) & Chr(0))
        DoEvents
        tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Child.FirstSibling.Index)
      Next
    End If
  End If
  tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
  
  Call SendProc(2, "toc_set_config " & BuddyConfig$ & Chr(0))
  
End Sub

Private Sub cmdInvite_Click()
  frmInvite.Show
End Sub

Private Sub cmdIM_Click()
  Dim lngFormIndex As Long
  If Not tvwBuddies.SelectedItem Is Nothing Then
    If Not tvwBuddies.SelectedItem.parent Is Nothing Then
      lngFormIndex& = FormByTag(LCase(Replace(tvwBuddies.SelectedItem.Text, " ", "")))
      If lngFormIndex& > -1 Then
        Forms(lngFormIndex&).SetFocus
      Else
        Dim frmNewIM As New frmIM
        With frmNewIM
          .Caption = tvwBuddies.SelectedItem.Text
          .Tag = LCase(Replace(tvwBuddies.SelectedItem.Text, " ", ""))
          .Show
        End With
      End If
    End If
  End If
End Sub

Private Sub cmdChat_Click()
  frmChatJoin.Show
End Sub

Private Sub Command1_Click()
Call SendProc(2, "toc_set_idle 2849340" & Chr(0))
'Call SendProc(2, "toc_set_idle 0" & Chr(0))
End Sub

Private Sub Form_Load()
  Dim lngDo As Long, nod() As Node
  m_strMode$ = GetINIString(m_strScreenName$, "mode", App.Path & "\aim.ini", "1")
  'Call LoadBuddies(m_strScreenName$)
  For lngDo& = 1 To tvwSetup.Nodes.Count
    If tvwSetup.Nodes.Item(lngDo&).parent Is Nothing Then
      ReDim Preserve nod(tvwBuddies.Nodes.Count + 1)
      Set nod(tvwBuddies.Nodes.Count) = tvwBuddies.Nodes.Add(, , , tvwSetup.Nodes.Item(lngDo&).Text, 1)
    End If
  Next
  m_strPDList$ = GetINIString(m_strScreenName$, "pdlist", App.Path & "\aim.ini", "")
  strSoundSignOn$ = GetINIString(m_strScreenName$, "signon sound", App.Path & "\aim.ini", "")
  strSoundSignOff$ = GetINIString(m_strScreenName$, "signoff sound", App.Path & "\aim.ini", "")
  strSoundFirstIM$ = GetINIString(m_strScreenName$, "firstim sound", App.Path & "\aim.ini", "")
  strSoundIMIn$ = GetINIString(m_strScreenName$, "imin sound", App.Path & "\aim.ini", "")
  strSoundIMOut$ = GetINIString(m_strScreenName$, "imout sound", App.Path & "\aim.ini", "")

'Dim intTransparent As Integer
'intTransparent parameter is the level of transparency, must be in between 0 and 255.
'intTransparent = 210
'SetLayeredWindow Me.hWnd, True
'SetLayeredWindowAttributes Me.hWnd, 0, intTransparent, LWA_ALPHA
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SaveBuddies(m_strScreenName$)
  Call WriteINIString(m_strScreenName$, "mode", m_strMode$, App.Path & "\aim.ini")
  frmSignOn.wskAIM.Close
  frmSignOn.Show
  Unload frmAway
End Sub

Private Sub mnuDebugShow_Click()
  frmDebug.Visible = True
End Sub

Private Sub mnuFileAway_Click()
frmAway.Show
End Sub

Private Sub mnuFileOptions_Click()
  frmOptions.Show
End Sub

Private Sub mnuFileSignOff_Click()
  Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuPeopleGetInfo_Click()
  Dim lngFormIndex As Long
  If Not tvwBuddies.SelectedItem Is Nothing Then
    If Not tvwBuddies.SelectedItem.parent Is Nothing Then
      lngFormIndex& = FormByTag(LCase(Replace(tvwBuddies.SelectedItem.Text, " ", "")))
      If lngFormIndex& > -1 Then
        Forms(lngFormIndex&).SetFocus
      Else
        frmInfo.Show
        frmInfo.WhoInfo.Text = tvwBuddies.SelectedItem.Text
        frmInfo.Caption = "Buddy Info: " & frmInfo.WhoInfo.Text
        Call SendProc(2, "toc_get_info " & Chr(34) & frmInfo.WhoInfo.Text & Chr(34) & Chr(0))
      End If
    End If
  Else
    frmInfo.Show
  End If
End Sub

Private Sub mnuPeopleJoinChat_Click()
  frmChatJoin.Show
End Sub

Private Sub Timer1_Timer()
Call SendProc(2, "toc_set_config " & BuddyConfig$ & Chr(0))
Timer1.Enabled = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
frmAway.Show
frmAway.Check1.Value = 1
If frmAway.Command1.Caption = "&Away" Then
frmAway.Text1.Enabled = False
frmAway.Command1.Caption = "&I'm Back"
Call SendProc(2, "toc_set_away " & Chr(34) & "<HTML>" & Text1.Text & "</HTML>" & Chr(34) & Chr(0))
MeAway = True
Else
frmAway.Text1.Enabled = True
Call SendProc(2, "toc_set_away " & Chr(34) & Chr(34) & Chr(0))
'Command1.Caption = "&Away"
Call ShowForm
MeAway = False
Unload frmAway
End If
End If
End Sub

Private Sub tvwBuddies_Collapse(ByVal Node As MSComctlLib.Node)
'Node.Image = "Collapsed"
End Sub

Private Sub tvwBuddies_DblClick()
  Dim lngFormIndex As Long
  If Not tvwBuddies.SelectedItem Is Nothing Then
    If Not tvwBuddies.SelectedItem.parent Is Nothing Then
      lngFormIndex& = FormByTag(LCase(Replace(tvwBuddies.SelectedItem.Text, " ", "")))
      If lngFormIndex& > -1 Then
        Forms(lngFormIndex&).SetFocus
      Else
        Dim frmNewIM As New frmIM
        With frmNewIM
          .Caption = tvwBuddies.SelectedItem.Text
          .Tag = LCase(Replace(tvwBuddies.SelectedItem.Text, " ", ""))
          .Show
        End With
      End If
    End If
  End If
End Sub

Private Sub tvwBuddies_Expand(ByVal Node As MSComctlLib.Node)
'Node.Image = "Expanded"
End Sub

Private Sub tvwSetup_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim parent As Boolean
Dim Y As Integer
Dim X As Integer
Dim tru$
tru$ = ""
  Dim nodGroup As Node
  If Trim(NewString) = "" Then
    MsgBox "Item can not be nothing.", vbCritical + vbOKOnly, "Error"
    tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
  ElseIf IsValidItem(NewString$) = False Then
    MsgBox "Item can contain only letters, numbers, and spaces.", vbCritical + vbOKOnly, "Error"
    tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
  ElseIf ExistsInTree(tvwSetup, NewString$) = True Then
    MsgBox Chr(34) & NewString$ & Chr(34) & "Already exists.", vbCritical + vbOKOnly, "Error"
    tvwSetup.Nodes.Remove (tvwSetup.SelectedItem.Index)
  Else
    If Not tvwSetup.SelectedItem.parent Is Nothing Then
      If ExistsInTree(tvwBuddies, tvwSetup.SelectedItem.Text, False, True) = True Then
        Call SendProc(2, "toc_remove_buddy " & Chr(34) & Replace(tvwSetup.SelectedItem.Text, " ", "") & Chr(34) & Chr(0))
      End If
      Call SendProc(2, "toc_add_buddy " & Replace(NewString, " ", "") & Chr(0))
   
      
      
 
      
    Else
      If ExistsInTree(tvwBuddies, tvwSetup.SelectedItem.Text, False, False, NewString$) = False Then
        Set nodGroup = tvwBuddies.Nodes.Add(, , , NewString$, 1, 1)
      End If
    End If
  End If
'  g - Buddy Group (All Buddies until the next g or the end of config
             'are in this group.)
    'b - A Buddy
    'p - Person on permit list
    'd - Person on deny list
    'm - Permit/Deny Mode.  Possible values are
    '1 - Permit All
    '2 - Deny All
    '3 - Permit Some
    '4 - Deny Some

Timer1.Enabled = True

End Sub

Private Function IsValidItem(strItem As String) As Boolean
  Dim lngDo As Long, blnIsValid As Boolean, strChar As String
  blnIsValid = True
  For lngDo& = 1 To Len(strItem$)
    strChar$ = Mid(strItem$, lngDo&, 1)
    If Asc(strChar$) < 65 Or Asc(strChar$) > 90 Then
      If Asc(strChar$) < 97 Or Asc(strChar$) > 122 Then
        If IsNumeric(strChar$) = False Then
          If strChar$ <> " " Then
            blnIsValid = False
            Exit For
          End If
        End If
      End If
    End If
  Next
  IsValidItem = blnIsValid
End Function

Private Sub tvwSetup_DblClick()
  If tvwSetup.SelectedItem Is Nothing Then Exit Sub
  tvwSetup.StartLabelEdit
End Sub
