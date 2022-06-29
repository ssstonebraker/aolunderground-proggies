VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSound 
      Left            =   120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "wav (*.wav)|*.wav"
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   29
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   3120
      Width           =   975
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Control"
      TabPicture(0)   =   "frmOptions.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Sounds"
      TabPicture(1)   =   "frmOptions.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "User List"
         Height          =   2295
         Left            =   2640
         TabIndex        =   24
         Top             =   480
         Width           =   2055
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   855
         End
         Begin VB.ListBox lstUsers 
            Height          =   1425
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Control"
         Height          =   2295
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   2415
         Begin VB.OptionButton optControl 
            Caption         =   "Block listed users."
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   1800
            Width           =   1935
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Block all users."
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1935
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Allow listed users only."
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Allow buddies only."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optControl 
            Caption         =   "Allow all users."
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buddy Sounds"
         Height          =   975
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   4575
         Begin VB.CommandButton cmdSignOff 
            Caption         =   "..."
            Height          =   285
            Left            =   4080
            TabIndex        =   14
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdSignOn 
            Caption         =   "..."
            Height          =   285
            Left            =   4080
            TabIndex        =   13
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtSoundSignoff 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtSoundSignon 
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Sign Off"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Sign On"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "IM Sounds"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   2
         Top             =   1440
         Width           =   4575
         Begin VB.CommandButton cmdIMOut 
            Caption         =   "..."
            Height          =   285
            Left            =   4080
            TabIndex        =   17
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdIMIn 
            Caption         =   "..."
            Height          =   285
            Left            =   4080
            TabIndex        =   16
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdFirstIM 
            Caption         =   "..."
            Height          =   285
            Left            =   4080
            TabIndex        =   15
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtSoundIMout 
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtSoundIMin 
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtSoundFirstIM 
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label5 
            Caption         =   "Message Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Message In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "First Message"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  Dim strRes As String, lngDo As Long, blnMatch As Boolean
  strRes$ = InputBox("Enter a screen name", "AIM")
  strRes$ = LCase(Replace(strRes$, " ", ""))
  If strRes$ <> "" Then
    For lngDo& = 0 To lstUsers.ListCount
      If lstUsers.List(lngDo&) = strRes$ Then
        MsgBox Chr(34) & strRes$ & Chr(34) & " already exists in the list."
        blnMatch = True
        Exit For
      End If
    Next
    If blnMatch = False Then
      lstUsers.AddItem strRes$
      If optControl(2).Value = True Then
        Call SendProc(2, "toc_add_permit " & strRes$ & " " & Chr(0))
      ElseIf optControl(3).Value = True Then
        Call SendProc(2, "toc_add_deny " & strRes$ & Chr(0))
      End If
    End If
  End If
End Sub

Private Sub cmdApply_Click()
  Dim lngDo As Long
  m_strPDList$ = ""
  If lstUsers.ListCount > 0 Then
    For lngDo& = 0 To lstUsers.ListCount
      m_strPDList$ = m_strPDList$ & " " & lstUsers.List(lngDo&)
    Next
  End If
  m_strPDList$ = Trim(m_strPDList$)
  strSoundSignOn$ = txtSoundSignon.Text
  strSoundSignOff$ = txtSoundSignoff.Text
  strSoundFirstIM$ = txtSoundFirstIM.Text
  strSoundIMIn$ = txtSoundIMin.Text
  strSoundIMOut$ = txtSoundIMout.Text
  If optControl(0).Value = True Then
    m_strMode$ = 1
  ElseIf optControl(1).Value = True Then
    m_strMode$ = 5
  ElseIf optControl(2).Value = True Then
    m_strMode$ = 3
  ElseIf optControl(3).Value = True Then
    m_strMode$ = 2
  Else
    m_strMode$ = 4
  End If
  Call WriteINIString(m_strScreenName$, "mode", m_strMode$, App.Path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "signon sound", strSoundSignOn$, App.Path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "signoff sound", strSoundSignOff$, App.Path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "firstim sound", strSoundFirstIM$, App.Path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "imin sound", strSoundIMIn$, App.Path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "imout sound", strSoundIMOut$, App.Path & "\aim.ini")
  Call WriteINIString(m_strScreenName$, "pdlist", m_strPDList$, App.Path & "\aim.ini")
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFirstIM_Click()
  dlgSound.ShowOpen
  txtSoundFirstIM.Text = dlgSound.FileName
End Sub

Private Sub cmdIMIn_Click()
  dlgSound.ShowOpen
  txtSoundIMin.Text = dlgSound.FileName
End Sub

Private Sub cmdIMOut_Click()
  dlgSound.ShowOpen
  txtSoundIMout.Text = dlgSound.FileName
End Sub

Private Sub cmdOK_Click()
  cmdApply_Click
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  If lstUsers.ListIndex > -1 Then
    If optControl(2).Value = True Then
      Call SendProc(2, "toc_add_deny " & lstUsers.Text & " " & Chr(0))
    ElseIf optControl(3).Value = True Then
      Call SendProc(2, "toc_add_permit " & lstUsers.Text & Chr(0))
    End If
    lstUsers.RemoveItem lstUsers.ListIndex
  End If
End Sub

Private Sub cmdSignOff_Click()
  dlgSound.ShowOpen
  txtSoundSignoff.Text = dlgSound.FileName
End Sub

Private Sub cmdSignOn_Click()
  dlgSound.ShowOpen
  txtSoundSignon.Text = dlgSound.FileName
End Sub

Private Sub Form_Load()
  Dim lngDo As Long, arrUsers() As String
  If m_strScreenName$ <> "" Then
    Select Case m_strMode$
      Case "1"
        optControl(0).Value = True
      Case "2"
        optControl(3).Value = True
      Case "3"
        optControl(2).Value = True
      Case "4"
        optControl(4).Value = True
      Case "5"
        optControl(1).Value = True
      Case Else
        optControl(0).Value = True
    End Select
    txtSoundSignon.Text = strSoundSignOn$
    txtSoundSignoff.Text = strSoundSignOff$
    txtSoundFirstIM.Text = strSoundFirstIM$
    txtSoundIMin.Text = strSoundIMIn$
    txtSoundIMout.Text = strSoundIMOut$
    If m_strPDList$ <> "" And m_strMode$ <> "5" Then
      lstUsers.Clear
      arrUsers$() = Split(m_strPDList$, " ")
      For lngDo& = LBound(arrUsers$) To UBound(arrUsers$)
        lstUsers.AddItem arrUsers$(lngDo&)
      Next
    End If
  End If
End Sub
