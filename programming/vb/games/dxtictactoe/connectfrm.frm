VERSION 5.00
Begin VB.Form connectfrm 
   Caption         =   "Choose Connection Type"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdok 
      Caption         =   "Connect"
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ListBox connectiontype 
      Height          =   2010
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
   Begin VB.TextBox PlayersName 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Players Name."
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2760
      Width           =   3855
   End
End
Attribute VB_Name = "connectfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  ' Force initialization of DirectPlay if user comes back in
  Unload Me
  'frmMainMenu.Show
End Sub

Private Sub cmdOK_Click()
  Dim cindex As Long
  Dim ConnectionMade As Boolean
  Dim dxAddress As DirectPlayAddress
  ' Initialize the connection. Any service provider dialogs are not called till the
  ' connection is used, e.g. to enumerate sessions.
  
  cindex = connectiontype.ListIndex + 1
  On Error GoTo INITIALIZEFAILED
  usermode = "host"
  
  If usermode = "host" Then
   Hide
   Lobby.Show
   Exit Sub
  Else
   Set dxAddress = EnumConnections.GetAddress(cindex)
  Call dxplay.InitializeConnection(dxAddress)
  End If
  
  On Error GoTo 0
  
  ' Enumerate the sessions to be shown in the SessionForm listbox. If this fails with
  ' DPERR_USERCANCEL, the user cancelled out of a service provider dialog. This is
  ' not a fatal error, because for the modem connection it simply indicates the
  ' player wishes to host a session and not to make a dial-up connection. The "answer"
  ' dialog will come up when the user attempts to create a session.
  
  ConnectionMade = Lobby.UpdateSessionList
  If ConnectionMade Then
    Hide
    Lobby.Show
  Else
    InitDPlay
  End If
  Exit Sub
  ' Error handlers
INITIALIZEFAILED:
  If Err.Number <> DPERR_ALREADYINITIALIZED Then
    MsgBox ("Failed to initialize connection.")
    Exit Sub
  End If
  
End Sub

Private Sub Form_Load()
  ' Enumerate connections
  If Not InitConnectionList Then
    Call MsgBox("Failed to Enumerate Connections.")
    CloseDownDPlay
    End
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'frmMainMenu.Show
End Sub

Private Sub lstConnections_DblClick()
    cmdOK_Click
End Sub

' Highlight player name when selected

Private Sub txtYourName_GotFocus()
  'txtYourName.SelStart = 0
  'txtYourName.SelLength = txtYourName.MaxLength
End Sub

' Enumerate connections

Public Function InitConnectionList() As Boolean

  Dim NumConnections As Long
  Dim strName As String
  Dim x As Long
  
  Call InitDPlay     ' Program aborts on failure
 
  connectiontype.Clear
  
  On Error GoTo FAILED
  NumConnections = EnumConnections.GetCount
  For x = 1 To NumConnections
    strName = EnumConnections.GetName(x)
    Call connectiontype.AddItem(strName)
  Next x
  
  ' Initialize selection
  connectiontype.ListIndex = 0
  InitConnectionList = True
  Exit Function
  
  ' Error handlers
FAILED:
  InitConnectionList = False
  Exit Function
  
End Function

