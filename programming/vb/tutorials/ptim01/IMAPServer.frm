VERSION 5.00
Begin VB.Form frmIMAPServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "IMAP Server"
   ClientHeight    =   5895
   ClientLeft      =   7815
   ClientTop       =   720
   ClientWidth     =   7470
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
   ScaleHeight     =   5895
   ScaleWidth      =   7470
   Begin VB.ListBox lstFileFlags 
      Height          =   2010
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ListBox lstFolderFlags 
      Height          =   2010
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtSelectedFolder 
      Height          =   315
      Index           =   0
      Left            =   5280
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.ListBox lstFolders 
      Height          =   2010
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Width           =   5415
   End
   Begin VB.ComboBox cboConnection 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtRequest 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtSocket 
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtState 
      Height          =   315
      Index           =   0
      Left            =   5880
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.ListBox lstFiles 
      Height          =   2010
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   3720
      Width           =   5415
   End
   Begin VB.Label lblSelectedFolder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Folder:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Connection:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblSocket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Socket:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "frmIMAPServer"
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
Option Base 1

Private Enum ParseStateType
  WhiteSpace = 0
  SimpleString = 1
  QuotedString = 2
  Parentheses = 3
  Braces = 4
End Enum

Dim ListenSocket As Long
Dim TransmitSocket As Long
Dim Status As Integer

Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Const IMAP4Port = 143
Dim CurrentConnection As Integer
Const PostOfficeDirectory = "C:\Post Office\"
Dim CRLF As String

Private Sub Form_Load()
  Dim ListenSocketAddress As InputSocketDescriptor

  Status = StartWinSock()
  CRLF = Chr(13) + Chr(10)
  cboConnection.AddItem ""
    
  txtRequest(0).Text = ""
  txtSocket(0).Text = ""
  txtRequest(0).Text = ""
  txtState(0).Text = ""
  txtUser(0).Text = ""
  lstFiles(0).AddItem ""
    
  Status = CreateSocket(ListenSocket, IMAP4Port)
  
  If listen(ListenSocket, 5) Then
      MsgBox "Could not listen on Port " + Str$(IMAP4Port)
      End
  End If
  
  If WSAAsyncSelect(ListenSocket, txtRequest(0).hwnd, WM_MBUTTONUP, FD_ACCEPT) Then
      MsgBox "Unable to set Asynch mode"
  End If

End Sub

  ' Accept incoming connection
Private Sub txtRequest_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim TransmitSocketAddress As OutputSocketDescriptor
  Dim TransmitSocketAddressSize As Integer
  Dim CurrentConnection As Integer
  
  TransmitSocketAddressSize = LenB(TransmitSocketAddress)
  TransmitSocketAddressSize = 20
  
  TransmitSocket = accept(ListenSocket, TransmitSocketAddress, TransmitSocketAddressSize)

  CurrentConnection = cboConnection.ListCount
    
  cboConnection.AddItem Trim(Str(CurrentConnection))
  Load txtSocket(CurrentConnection)
  Load txtRequest(CurrentConnection)
  Load txtState(CurrentConnection)
  Load txtUser(CurrentConnection)
  Load txtSelectedFolder(CurrentConnection)
  Load lstFolders(CurrentConnection)
  Load lstFolderFlags(CurrentConnection)
  Load lstFiles(CurrentConnection)
  Load lstFileFlags(CurrentConnection)
  
  txtSocket(CurrentConnection) = Trim(Str(TransmitSocket))
  txtState(CurrentConnection).Text = "Authorisation"
  
  DisplayConnection CurrentConnection
  
  Status = SendSocket(TransmitSocket, "* OK IMAP server here")
    
  If WSAAsyncSelect(TransmitSocket, txtRequest(CurrentConnection).hwnd, WM_MBUTTONDBLCLK, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If
    
End Sub

  ' Process incoming request
Private Sub txtRequest_DblClick(Index As Integer)
  Dim SendLine As String
  Dim ClientRequest As String
  Dim CommandIdentifier As String
  Dim CommandKeyword As String
  Dim StatusResponse As String
  
  Status = ReceiveSocket(TransmitSocket, ClientRequest)
  DisplayConnection Index
  txtRequest(Index).Text = ClientRequest
  
  Debug.Print GetToken(ClientRequest, 1)
  Debug.Print GetToken(ClientRequest, 2)
  Debug.Print GetToken(ClientRequest, 3)
  Debug.Print GetToken(ClientRequest, 4)
  Debug.Print GetToken(ClientRequest, 5)
  
  CommandIdentifier = GetToken(ClientRequest, 1)
  
  CommandKeyword = UCase(GetToken(ClientRequest, 2))
  If CommandKeyword = "UID" Then
    CommandKeyword = "UID " + UCase(GetToken(ClientRequest, 3))
  End If
  
  Debug.Print CommandKeyword & " started"
  
  Select Case CommandKeyword
    Case "CAPABILITY"
      Status = SendSocket(TransmitSocket, "* CAPABILITY IMAP4 IMAP4rev1")
    Case "LOGIN"
      LoadUser Index, GetToken(ClientRequest, 3)
    Case "LIST", "LSUB"
      ListFolders Index, ClientRequest
    Case "SELECT"
      LoadFolder Index, GetToken(ClientRequest, 3)
    Case "UID FETCH", "FETCH"
      FetchMessageRange Index, GetToken(ClientRequest, 4), GetToken(ClientRequest, 5)
    Case "CREATE"
      Debug.Print PostOfficeDirectory + "\" + txtUser(Index).Text + "\" + GetToken(ClientRequest, 3)
      MkDir PostOfficeDirectory + "\" + txtUser(Index).Text + "\" + GetToken(ClientRequest, 3)
    Case "DELETE"
      RmDir PostOfficeDirectory + "\" + txtUser(Index).Text + "\" + GetToken(ClientRequest, 3)
    Case "STORE"
'      lstFileFlags(Index).List(GetToken(ClientRequest, 3)) = GetToken(ClientRequest, 3)
    Case "COPY"
        FileCopy GetToken(ClientRequest, 3), GetToken(ClientRequest, 3)

    Case "SUBSCRIBE"
    Case "UNSUBSCRIBE"

  End Select
  
  Debug.Print CommandKeyword & " completed"

  StatusResponse = CommandIdentifier + " OK " + CommandKeyword + " completed"
  Status = SendSocket(TransmitSocket, StatusResponse)
  
  Debug.Print CommandKeyword & " confirmed"

End Sub

Private Sub LoadUser(Index As Integer, UserName As String)
  Dim CurrentFile As String
  Dim CurrentDirectory As String
  Dim varTemporaryItem
  Dim colTemporaryList As New Collection
  
  txtUser(Index).Text = UserName

  CurrentDirectory = PostOfficeDirectory + UserName + "\"
  CurrentFile = Dir$(CurrentDirectory, vbDirectory)
  
  Do Until CurrentFile = ""
    If CurrentFile <> "." And CurrentFile <> ".." And _
       (GetAttr(CurrentDirectory & CurrentFile) And vbDirectory) = vbDirectory Then
      Debug.Print CurrentFile
      colTemporaryList.Add CurrentFile
    End If
    CurrentFile = Dir$
  Loop

  For Each varTemporaryItem In colTemporaryList
      lstFolders(CurrentConnection).AddItem varTemporaryItem
      LoadSubfolders Index, UserName, CStr(varTemporaryItem)
  Next

End Sub

Private Sub LoadSubfolders(Index As Integer, UserName As String, FolderRoot As String)
  Dim CurrentFile As String
  Dim CurrentDirectory As String
  Dim varTemporaryItem
  Dim colTemporaryList As New Collection
  
  CurrentDirectory = PostOfficeDirectory + UserName + "\" + FolderRoot + "\"
  CurrentFile = Dir$(CurrentDirectory, vbDirectory)
  
  Do Until CurrentFile = ""
    If CurrentFile <> "." And CurrentFile <> ".." And _
       (GetAttr(CurrentDirectory & CurrentFile) And vbDirectory) = vbDirectory Then
      Debug.Print CurrentFile
      colTemporaryList.Add CurrentFile
    End If
    CurrentFile = Dir$
  Loop

  For Each varTemporaryItem In colTemporaryList
      lstFolders(CurrentConnection).AddItem FolderRoot + "\" + varTemporaryItem
      LoadSubfolders Index, UserName, FolderRoot + "\" + varTemporaryItem
  Next

End Sub

Private Sub ListFolders(Index As Integer, ClientRequest As String)
  Dim strResponseLine As String
  Dim strRequestedFolder As String
  Dim strCurrentFolder As String
  Dim intFolderIndex As Integer
  
  strRequestedFolder = GetToken(ClientRequest, 4)
  
  For intFolderIndex = 0 To lstFolders(Index).ListCount - 1
      strCurrentFolder = lstFolders(Index).List(intFolderIndex)
      Debug.Print strCurrentFolder
      strResponseLine = "* " + GetToken(ClientRequest, 2) + " () ""\"" """ + strCurrentFolder + """"
      Debug.Print strResponseLine
      Status = SendSocket(TransmitSocket, DoubleBackslashes(strResponseLine))
  Next

End Sub

Private Sub LoadFolder(CurrentConnection As Integer, FolderName As String)
  Dim CurrentUser As String
  Dim CurrentDirectory As String
  Dim CurrentFile As String
  Dim CurrentFileSize As Integer

  txtSelectedFolder(CurrentConnection).Text = FolderName
  
  CurrentUser = txtUser(CurrentConnection).Text
  CurrentDirectory = PostOfficeDirectory + CurrentUser + "\" + FolderName + "\"
  CurrentDirectory = PostOfficeDirectory + CurrentUser + "\"
  
  CurrentFile = Dir$(CurrentDirectory + "*")
  
  Do Until CurrentFile = ""
  
    lstFiles(CurrentConnection).AddItem CurrentFile
    lstFileFlags(CurrentConnection).AddItem ""
    CurrentFile = Dir$
    
  Loop
  
End Sub


Private Sub FetchMessageRange(CurrentConnection As Integer, MessageRange As String, MessageComponents As String)
  Dim intBeginIndex As Integer
  Dim intEndIndex As Integer
  Dim intCurrentIndex As Integer
  
  If InStr(1, MessageRange, ":") = 0 Then
    FetchMessage CurrentConnection, Int(val(MessageRange)), MessageComponents
  Else
    intBeginIndex = Int(val(Left(MessageRange, InStr(1, MessageRange, ":"))))
    If intBeginIndex = 0 Then intBeginIndex = 1
    intEndIndex = Int(val(Mid(MessageRange, InStr(1, MessageRange, ":") + 1)))
    If intEndIndex = 0 Then intEndIndex = lstFiles(CurrentConnection).ListCount
    For intCurrentIndex = intBeginIndex To intEndIndex
      FetchMessage CurrentConnection, intCurrentIndex, MessageComponents
    Next
  End If
End Sub

Private Sub FetchMessage(CurrentConnection As Integer, MessageIndex As Integer, MessageComponents As String)
  Dim strSendLine As String
  Dim strMessageFile As String
  Dim intMessageFileNumber As Integer
  Dim strCurrentUser As String
  Dim strCurrentDirectory As String
  Dim strComponents As String
  Dim strCurrentComponent As String
  Dim intComponentIndex As Integer
  Dim strCurrentHeader As String
  Dim strHeaders As String
  
  intComponentIndex = 1
  strComponents = Trim(UCase(MessageComponents))
  If Left(strComponents, 1) = "(" Then
    strComponents = Trim(Mid(strComponents, 2, Len(strComponents) - 2))
  End If
  strCurrentComponent = GetToken(strComponents, intComponentIndex)

  strCurrentUser = txtUser(CurrentConnection).Text
  
  strCurrentDirectory = PostOfficeDirectory + strCurrentUser + "\"
  strMessageFile = strCurrentDirectory + lstFiles(CurrentConnection).List(MessageIndex - 1)

  strSendLine = "* " + Trim(Str(MessageIndex)) + " FETCH ("

  Do While strCurrentComponent <> ""
  
    Debug.Print strCurrentComponent
    Select Case strCurrentComponent
      Case "RFC822", "BODY[]", "BODY.PEEK[]"
        strSendLine = strSendLine + strCurrentComponent + _
                      " {" + Trim(Str(FileLen(strMessageFile)) + "}")
        Status = SendSocket(TransmitSocket, strSendLine)
        intMessageFileNumber = FreeFile
        Open strMessageFile For Input As #intMessageFileNumber
        Do While Not EOF(intMessageFileNumber)
          Line Input #intMessageFileNumber, strSendLine
          Status = SendSocket(TransmitSocket, strSendLine)
        Loop
        Close #intMessageFileNumber
        strSendLine = " "
      
      Case "RFC822.HEADER"
        strHeaders = ""
        intMessageFileNumber = FreeFile
        Open strMessageFile For Input As #intMessageFileNumber
        Do
          Line Input #intMessageFileNumber, strCurrentHeader
          strHeaders = strHeaders + CRLF + strCurrentHeader
        Loop While Not EOF(intMessageFileNumber) And strCurrentHeader <> ""
        Close #intMessageFileNumber
        strHeaders = strHeaders + CRLF
        strSendLine = strSendLine + strCurrentComponent + _
                      " {" + Trim(Str(Len(strHeaders)) + "}")
        Status = SendSocket(TransmitSocket, strSendLine)
        Status = SendSocket(TransmitSocket, Mid(strHeaders, 3))
        strSendLine = " "
      
      Case "RFC822.SIZE"
        strSendLine = strSendLine + strCurrentComponent + " " + _
                      Trim(Str(FileLen(strMessageFile))) + " "
      
      Case "INTERNALDATE"
        strSendLine = strSendLine + strCurrentComponent + " """ + _
                       Format(FileDateTime(strMessageFile), "dd-mmm-yy hh:mm:ss") + _
                       " +0000"" "
      
      Case "UID"
        strSendLine = strSendLine + strCurrentComponent + " " + _
                      Trim(Str(MessageIndex)) + " "
      
      Case "FLAGS"
        strSendLine = strSendLine + strCurrentComponent + " (\Seen) "
    
    End Select
    
    intComponentIndex = intComponentIndex + 1
    strCurrentComponent = GetToken(strComponents, intComponentIndex)
  Loop
  strSendLine = RTrim(strSendLine) + ")"
  Status = SendSocket(TransmitSocket, strSendLine)

End Sub

Private Function GetToken(SearchString As String, TokenIndex As Integer) As String
  Dim ScanState As ParseStateType
  Dim CurrentCharacter As String
  Dim ScanIndex As Integer
  Dim TokenFound As Boolean
  Dim ScanToken As Integer
  Dim TokenBegin As Integer
  Dim NestDepth As Integer
  Dim WorkToken As String
    
  WorkToken = ""
  ScanIndex = 1
  ScanToken = 0
  TokenFound = False
  ScanState = WhiteSpace
  
  Do While ScanIndex <= Len(SearchString)
    CurrentCharacter = Mid(SearchString, ScanIndex, 1)
    Select Case ScanState
      Case WhiteSpace
        If Not CurrentCharacter = " " Then
          ScanToken = ScanToken + 1
          WorkToken = CurrentCharacter
          Select Case CurrentCharacter
            Case "("
              ScanState = Parentheses
              NestDepth = 1
            Case """"
              ScanState = QuotedString
              WorkToken = ""
            Case "{"
              ScanState = Braces
              NestDepth = 1
            Case Else
              ScanState = SimpleString
          End Select
        End If
        ScanIndex = ScanIndex + 1
      Case SimpleString
        Select Case CurrentCharacter
          Case " "
            ScanState = WhiteSpace
            If ScanToken = TokenIndex Then
              GetToken = WorkToken
              Exit Function
            End If
          Case Else
            WorkToken = WorkToken + CurrentCharacter
            ScanIndex = ScanIndex + 1
        End Select
      Case QuotedString
        If CurrentCharacter = """" Then
          ScanState = WhiteSpace
          If ScanIndex < Len(SearchString) Then
            If Mid(SearchString, ScanIndex + 1, 1) = """" Then
              ScanIndex = ScanIndex + 1
              ScanState = QuotedString
            End If
          End If
          If ScanState = WhiteSpace And ScanToken = TokenIndex Then
            GetToken = WorkToken
            Exit Function
          End If
        End If
        WorkToken = WorkToken + CurrentCharacter
        ScanIndex = ScanIndex + 1
      Case Parentheses
        Select Case CurrentCharacter
          Case "("
            NestDepth = NestDepth + 1
          Case ")"
            NestDepth = NestDepth - 1
        End Select
        WorkToken = WorkToken + CurrentCharacter
        If NestDepth = 0 Then
          ScanState = WhiteSpace
          If ScanToken = TokenIndex Then
            GetToken = WorkToken
            Exit Function
          End If
        End If
        ScanIndex = ScanIndex + 1
      Case Braces
        Select Case CurrentCharacter
          Case "{"
            NestDepth = NestDepth + 1
          Case "}"
            NestDepth = NestDepth - 1
        End Select
        WorkToken = WorkToken + CurrentCharacter
        If NestDepth = 0 Then
          ScanState = WhiteSpace
          If ScanToken = TokenIndex Then
            GetToken = WorkToken
            Exit Function
          End If
        End If
        ScanIndex = ScanIndex + 1
    End Select
  Loop

  If ScanToken = TokenIndex Then
    GetToken = WorkToken
  Else
    GetToken = ""
  End If

End Function

Private Function DoubleBackslashes(InputString As String) As String
  Dim intIndex As Integer

  For intIndex = 1 To Len(InputString)
    If Mid(InputString, intIndex, 1) = "\" Then
      DoubleBackslashes = DoubleBackslashes + "\\"
    Else
      DoubleBackslashes = DoubleBackslashes + Mid(InputString, intIndex, 1)
    End If
  Next

End Function

Private Sub cboConnection_Click()
  DisplayConnection cboConnection.ListIndex
End Sub
 
Private Sub DisplayConnection(NewConnection As Integer)
  If NewConnection < 0 Or NewConnection >= cboConnection.ListCount Then
    NewConnection = 0
  End If
  
  If NewConnection = CurrentConnection Then Exit Sub
  
  txtRequest(CurrentConnection).Visible = False
  txtRequest(NewConnection).Visible = True
  txtSocket(CurrentConnection).Visible = False
  txtSocket(NewConnection).Visible = True
  txtState(CurrentConnection).Visible = False
  txtState(NewConnection).Visible = True
  txtUser(CurrentConnection).Visible = False
  txtUser(NewConnection).Visible = True
  txtSelectedFolder(CurrentConnection).Visible = False
  txtSelectedFolder(NewConnection).Visible = True
  lstFolders(CurrentConnection).Visible = False
  lstFolders(NewConnection).Visible = True
  lstFolderFlags(CurrentConnection).Visible = False
  lstFolderFlags(NewConnection).Visible = True
  lstFiles(CurrentConnection).Visible = False
  lstFiles(NewConnection).Visible = True
  lstFileFlags(CurrentConnection).Visible = False
  lstFileFlags(NewConnection).Visible = True
      
  cboConnection.ListIndex = NewConnection
  
  CurrentConnection = NewConnection
  
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  Status = ReleaseSocket(ListenSocket)
  Status = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(Status)
End Sub

Private Sub Form_Deactivate()
  Status = ReleaseSocket(ListenSocket)
  Status = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(Status)
End Sub
