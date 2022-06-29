VERSION 5.00
Begin VB.Form frmViewIMAPMessages 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Internet Mail  Client"
   ClientHeight    =   5985
   ClientLeft      =   1395
   ClientTop       =   2220
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleWidth      =   9510
   Begin VB.Frame grpMessages 
      Caption         =   "Messages"
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   8895
      Begin VB.CommandButton btnCopy 
         Caption         =   "Co&py"
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ListBox lstMessageList 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   8415
      End
      Begin VB.CommandButton btnShow 
         Caption         =   "&Show"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton btnExtract 
         Caption         =   "&Extract"
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
   End
   Begin VB.Frame grpFolders 
      Caption         =   "Folders"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CheckBox chkShowAllFolders 
         Caption         =   "Show all folders"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnDeleteFolder 
         Caption         =   "De&lete"
         Height          =   375
         Left            =   7080
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton btnCreateFolder 
         Caption         =   "C&reate"
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton btnUnsubscribeFolder 
         Caption         =   "&Unsubscribe"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton btnSubscribeFolder 
         Caption         =   "Su&bscribe"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ListBox lstFolderList 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   8415
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuCompose 
      Caption         =   "&Create"
      Begin VB.Menu mnuCreateFolder 
         Caption         =   "&Folder"
      End
      Begin VB.Menu mnuCreateMessage 
         Caption         =   "&Message"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmViewIMAPMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
' Copyright 1999 - Digital Press, John Rhoton
'
' This program has been written to illustrate the Internet Mail protocols.
' It is provided free of charge and unconditionally.  However, it is not
' intended for production use, and therefore without warranty or any
' implication of support.
'
' You can find an explanation of the concepts behind this code in
' the book:  Programmer's Guide to Internet Mail by John Rhoton,
' Digital Press 1999.  ISBN: 1-55558-212-5.
'
' For ordering information please see http://www.amazon.com or
' you can order directly with http://www.bh.com/digitalpress.
'
'========================================================================

Option Explicit

Public IMAPServer As String
Public IMAPUser As String
Public IMAPPassword As String

Const QuoteQuote = """"""
Const QuoteStarQuote = """*"""
Const QuoteSlashQuote = """/"""
Const Quote = """"
Const Blank = " "

Dim lngIMAPSocket As Long
Const IMAPPort = 143
Dim CRLF As String
Dim intStatus As Integer
Dim intCommandIndex As Integer

Private Sub btnCopy_Click()
  Dim strSocketResponse As String
  Dim intRetrieveIndex As Integer
  
  intRetrieveIndex = lstMessageList.ListIndex + 1
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + _
                                        "COPY " + Trim(Str(intRetrieveIndex)) + _
                                        " " + Quote + _
                                        InputBox("Folder Name:") + Quote)

  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Debug.Print strSocketResponse
  
End Sub

Private Sub btnCreateFolder_Click()
  Dim strSocketResponse As String
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + "CREATE " + Quote + _
                                        InputBox("Folder Name:") + Quote)
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  ListFolders
End Sub

Private Sub btnDeleteFolder_Click()
  Dim strSocketResponse As String
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + "DELETE " + Quote + _
                                        lstFolderList.Text + Quote)
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  ListFolders
End Sub

Private Sub btnSubscribeFolder_Click()
  Dim strSocketResponse As String
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + "SUBSCRIBE " + Quote + _
                                        lstFolderList.Text + Quote)
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  ListFolders
End Sub

Private Sub btnUnsubscribeFolder_Click()
  Dim strSocketResponse As String
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + "UNSUBSCRIBE " + Quote + _
                                        lstFolderList.Text + Quote)
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  ListFolders
End Sub


Private Sub Form_Load()
  
  CRLF = Chr(13) + Chr(10)
  intStatus = StartWinSock()
  Me.Left = Me.Width

  Dim lngServerAddress As Long
  Dim strServerName As String
  
  Dim strSocketResponse As String
  
  lngIMAPSocket = 0

  LoadParameters

  intStatus = GetIPAddress(lngServerAddress, IMAPServer)
  intStatus = CreateSocket(lngIMAPSocket, 0)
  
  frmVBWinSockDialog.Left = Me.Left
  frmVBWinSockDialog.Top = Me.Height
  
  intStatus = ConnectSocket(lngIMAPSocket, lngServerAddress, IMAPPort)
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + _
                                        "LOGIN " + _
                                        Quote + IMAPUser + Quote + Blank + _
                                        Quote + IMAPPassword + Quote)
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  
  ListFolders
  
'  btnShow.Enabled = True
'  btnExtract.Enabled = True
'  btnDelete.Enabled = True
'  btnShow.DEFAULT = True
    
End Sub

Private Sub LoadParameters()

    ' These parameters/preferences can be stored in a file or in the
    ' registry

  IMAPServer = InputBox("IMAP Server")
  IMAPUser = InputBox("IMAP User")
  IMAPPassword = InputBox("IMAP Password")

  frmLDAPClient.LDAPServer = InputBox("LDAP Server", , IMAPServer)
  
  frmSendSMTPMessage.SMTPServer = InputBox("SMTP Server", , IMAPServer)
  frmSendSMTPMessage.SMTPFrom = InputBox("SMTP From", , IMAPUser)

End Sub

Private Function NextIndex() As String
  intCommandIndex = intCommandIndex + 1
  NextIndex = CurrentIndex
End Function
Private Function CurrentIndex() As String
  If intCommandIndex < 1000 Then intCommandIndex = 1000
  CurrentIndex = Right(Trim(Str(intCommandIndex)), 4) + Blank
End Function

Private Function GetToken(InputString As String, TokenIndex As Integer) As String

  Dim intCharacterIndex As Integer
  Dim intTokenIndex As Integer
  
  
  For intCharacterIndex = 1 To Len(InputString)
  
  
  Next
  
  GetToken = Mid(InputString, 15)
End Function

Private Sub ListFolders()
  Dim strSocketResponse As String

  lstFolderList.Clear
  If chkShowAllFolders.Value = 0 Then
    intStatus = SendSocket(lngIMAPSocket, NextIndex + "LSUB " + QuoteQuote + _
                                          Blank + QuoteStarQuote)
    lstFolderList.AddItem "INBOX"
  Else
    intStatus = SendSocket(lngIMAPSocket, NextIndex + "LIST " + QuoteQuote + _
                                          Blank + QuoteStarQuote)
  End If
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  
 
  Do While Left(strSocketResponse, Len(CurrentIndex)) <> CurrentIndex
    lstFolderList.AddItem Mid(strSocketResponse, InStr(5, strSocketResponse, QuoteSlashQuote) + 4)
    intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Loop

End Sub
Private Sub chkShowAllFolders_Click()
  ListFolders
End Sub


Private Sub btnExtract_Click()

  Dim strSocketResponse As String
  Dim intWorkIndex As Integer
  Dim intRetrieveIndex As Integer
  Dim strFileSpecification As String
  
  intRetrieveIndex = lstMessageList.ListIndex + 1
 
  strFileSpecification = InputBox("Please enter file: ")
  Open strFileSpecification For Output As #1
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + _
                                        "FETCH " + Trim(Str(intRetrieveIndex)) + " " + _
                                        "(RFC822)")

  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  
  Debug.Print strSocketResponse
  
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Do
    Print #1, strSocketResponse
    Debug.Print ">" + strSocketResponse + "<"
    intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Loop While Left(Trim(strSocketResponse), 5) <> CurrentIndex
  
  Close #1

End Sub
Private Sub btnShow_Click()

  Dim strSocketResponse As String
  Dim intWorkIndex As Integer
  Dim intRetrieveIndex As Integer
  
  intRetrieveIndex = lstMessageList.ListIndex + 1
  
  Load frmReadMIMEMessages
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + _
                                        "FETCH " + Trim(Str(intRetrieveIndex)) + " " + _
                                        "(RFC822)")

  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  
  Debug.Print strSocketResponse
  
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Do
    frmReadMIMEMessages.DisplayMessage.Lines.Add strSocketResponse
    Debug.Print ">" + strSocketResponse + "<"
    intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Loop While Left(Trim(strSocketResponse), 5) <> CurrentIndex
  
  frmReadMIMEMessages.Show
    
End Sub
Private Sub btnDelete_Click()
  Dim strSocketResponse As String
  Dim intRetrieveIndex As Integer
  
  intRetrieveIndex = lstMessageList.ListIndex + 1
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + _
                                        "STORE " + Trim(Str(intRetrieveIndex)) + " " + _
                                        "+FLAGS (\Deleted)")

  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Debug.Print strSocketResponse
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + "EXPUNGE")
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Debug.Print strSocketResponse
  

End Sub

Private Sub btnQuit_Click()
    
  Dim strSocketResponse As String
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + "LOGOUT ")
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  intStatus = ReleaseSocket(lngIMAPSocket)

  End
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(intStatus)
End Sub

Private Sub lstFolderList_DblClick()
  ' select and display

  Dim strSocketResponse As String
  Dim strMessageDate As String
  Dim strMessageSubject As String
  Dim strMessageFrom As String
'  Dim strMessageDisplay As String
  
  lstMessageList.Clear
  intStatus = SendSocket(lngIMAPSocket, NextIndex + "SELECT " + lstFolderList.Text)
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Do While Left(strSocketResponse, Len(CurrentIndex)) <> CurrentIndex
    intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  Loop
  
  intStatus = SendSocket(lngIMAPSocket, NextIndex + _
                        "FETCH 1:* (BODY[HEADER.FIELDS (DATE SUBJECT FROM)])")
  intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
  
  Do While Left(strSocketResponse, Len(CurrentIndex)) <> CurrentIndex
    intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
    strMessageDate = Mid(strSocketResponse, InStr(5, strSocketResponse, ":") + 2)
    
    intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
    strMessageSubject = Mid(strSocketResponse, InStr(5, strSocketResponse, ":") + 2)
    
    intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
    strMessageFrom = Mid(strSocketResponse, InStr(5, strSocketResponse, ":") + 2)
    
    lstMessageList.AddItem Left(strMessageDate + Space(20), 20) + " " + _
                           Left(strMessageFrom + Space(35), 35) + " " + _
                           Left(strMessageSubject, 20)
    Do While Left(strSocketResponse, 1) <> "*" And _
             Left(strSocketResponse, Len(CurrentIndex)) <> CurrentIndex
           intStatus = ReceiveSocket(lngIMAPSocket, strSocketResponse)
    Loop
  Loop

End Sub

Private Sub mnuNewMessage_Click()
  Load frmSendSMTPMessage
  frmSendSMTPMessage.Show
End Sub

Private Sub mnuCreateMessage_Click()
  Load frmSendSMTPMessage
  frmSendSMTPMessage.Show
End Sub
