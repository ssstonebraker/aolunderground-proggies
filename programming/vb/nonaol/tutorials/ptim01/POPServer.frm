VERSION 5.00
Begin VB.Form frmPOPServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "POP Server"
   ClientHeight    =   5025
   ClientLeft      =   1140
   ClientTop       =   2130
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
   ScaleHeight     =   5025
   ScaleWidth      =   7470
   Begin VB.ListBox lstSizeList 
      Height          =   3570
      Index           =   0
      Left            =   600
      TabIndex        =   15
      Top             =   1320
      Width           =   735
   End
   Begin VB.ListBox lstDeletionList 
      Height          =   3570
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   375
   End
   Begin VB.ComboBox cboConnection 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtRequest 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   12
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtSocket 
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtState 
      Height          =   315
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtTotalCount 
      Height          =   315
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtTotalSize 
      Height          =   315
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox lstFileList 
      Height          =   3570
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label lblConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Connection:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblTotalCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Count:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblTotalSize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Size:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmPOPServer"
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

Dim lngListenSocket As Long
Dim lngTransmitSocket As Long
Dim intStatus As Integer

Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Const POP3Port = 110
Dim intCurrentConnection As Integer
Const PostOfficeDirectory = "C:\Post Office\"
Dim CRLF As String

Private Sub Form_Load()

  intStatus = StartWinSock()
  
  CRLF = Chr(13) + Chr(10)
  
  cboConnection.AddItem "0"
    
  txtRequest(0).Text = ""
  txtSocket(0).Text = ""
  txtRequest(0).Text = ""
  txtState(0).Text = ""
  txtUser(0).Text = ""
  txtTotalCount(0).Text = ""
  txtTotalSize(0).Text = ""
  lstFileList(0).AddItem ""
    
  intStatus = CreateSocket(lngListenSocket, POP3Port)
  
  If listen(lngListenSocket, 5) Then
      MsgBox "Could not listen on Port " + Str$(POP3Port)
      End
  End If
  
  If WSAAsyncSelect(lngListenSocket, txtRequest(0).hwnd, WM_MBUTTONUP, FD_ACCEPT) Then
      MsgBox "Unable to set Asynch mode"
  End If

End Sub

Private Sub txtRequest_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim sotTransmitSocketAddress As OutputSocketDescriptor
  Dim intTransmitSocketAddressSize As Integer
  Dim intCurrentConnection As Integer
  
  intTransmitSocketAddressSize = LenB(sotTransmitSocketAddress)
  intTransmitSocketAddressSize = 20
  
  lngTransmitSocket = accept(lngListenSocket, sotTransmitSocketAddress, intTransmitSocketAddressSize)

  intCurrentConnection = cboConnection.ListCount
    
  cboConnection.AddItem Trim(Str(intCurrentConnection))
  Load txtSocket(intCurrentConnection)
  Load txtRequest(intCurrentConnection)
  Load txtState(intCurrentConnection)
  Load txtUser(intCurrentConnection)
  Load txtTotalCount(intCurrentConnection)
  Load txtTotalSize(intCurrentConnection)
  Load lstDeletionList(intCurrentConnection)
  Load lstSizeList(intCurrentConnection)
  Load lstFileList(intCurrentConnection)
  
  txtSocket(intCurrentConnection) = Trim(Str(lngTransmitSocket))
  txtState(intCurrentConnection).Text = "Authorisation"
  
  DisplayConnection intCurrentConnection
  
  intStatus = SendSocket(lngTransmitSocket, "+OK Pop server here")
    
  If WSAAsyncSelect(lngTransmitSocket, txtRequest(intCurrentConnection).hwnd, WM_MBUTTONDBLCLK, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If
    
End Sub

Private Sub txtRequest_DblClick(Index As Integer)
  Dim strSendLine As String
  Dim strClientRequest As String
  Dim strCommand As String
  Dim strParameters As String
  
  DisplayConnection Index
  intStatus = ReceiveSocket(lngTransmitSocket, strClientRequest)
  strClientRequest = Trim(strClientRequest)
  txtRequest(Index).Text = strClientRequest
  strCommand = Trim(UCase(Left$(strClientRequest, 4)))
  strParameters = Trim(Mid(strClientRequest, 6))
  
  Select Case strCommand
    Case "USER"
      txtUser(Index).Text = strParameters
      intStatus = SendSocket(lngTransmitSocket, "+OK")
    Case "PASS"
      If strParameters = txtUser(Index).Text Then
        intStatus = SendSocket(lngTransmitSocket, "+OK")
        txtState(Index).Text = "Transaction"
        LoadUserFiles Index
      Else
        intStatus = SendSocket(lngTransmitSocket, "-ERR")
      End If
    Case "STAT"
      strSendLine = "+OK " + txtTotalCount(Index).Text + " " + txtTotalSize(Index).Text
      intStatus = SendSocket(lngTransmitSocket, strSendLine)
    Case "LIST"
      ListMessages Index
    Case "RETR"
      RetrieveMessage Index, val(strParameters)
    Case "DELE"
      lstDeletionList(Index).List(val(strParameters) - 1) = "D"
      intStatus = SendSocket(lngTransmitSocket, "+OK")
    Case "QUIT"
      txtState(Index).Text = "Update"
      DeleteMessages Index
    Case Else
      intStatus = SendSocket(lngTransmitSocket, "-ERR " + strCommand + " not supported")
  End Select

End Sub

Private Sub LoadUserFiles(intCurrentConnection As Integer)
  Dim strCurrentUser As String
  Dim strCurrentDirectory As String
  Dim strCurrentFile As String
  Dim intCurrentFileSize As Integer
  Dim intTotalCount As Integer
  Dim intTotalSize As Integer

  strCurrentUser = txtUser(intCurrentConnection).Text
  strCurrentDirectory = PostOfficeDirectory + strCurrentUser + "\"
  
  strCurrentFile = Dir$(strCurrentDirectory + "*")
  intTotalCount = 0
  intTotalSize = 0
  
  Do Until strCurrentFile = ""
  
    intTotalCount = intTotalCount + 1
    intCurrentFileSize = FileLen(strCurrentDirectory + strCurrentFile)
    intTotalSize = intTotalSize + intCurrentFileSize
    
    lstDeletionList(intCurrentConnection).AddItem Space(1)
    lstSizeList(intCurrentConnection).AddItem Str(intCurrentFileSize)
    lstFileList(intCurrentConnection).AddItem strCurrentFile
    strCurrentFile = Dir$
    
  Loop
  
  txtTotalCount(intCurrentConnection).Text = Trim(Str(intTotalCount))
  txtTotalSize(intCurrentConnection).Text = Trim(Str(intTotalSize))
  
End Sub

Private Sub ListMessages(intCurrentConnection As Integer)
  Dim strSendLine As String
  Dim intMessageIndex As Integer
  
  strSendLine = "+OK " + txtTotalCount(intCurrentConnection).Text + " Messages"
  intStatus = SendSocket(lngTransmitSocket, strSendLine)
  
  For intMessageIndex = 0 To val(txtTotalCount(intCurrentConnection).Text) - 1
    strSendLine = Str(intMessageIndex + 1) + " " + lstSizeList(intCurrentConnection).List(intMessageIndex)
    intStatus = SendSocket(lngTransmitSocket, strSendLine)
  Next
  intStatus = SendSocket(lngTransmitSocket, ".")

End Sub

Private Sub RetrieveMessage(intCurrentConnection As Integer, MessageIndex As Integer)
  Dim strSendLine As String
  Dim strCurrentFile As String
  Dim intCurrentFileNumber As Integer
  Dim strCurrentUser As String
  Dim strCurrentDirectory As String

  intStatus = SendSocket(lngTransmitSocket, "+OK ")
  
  strCurrentUser = txtUser(intCurrentConnection).Text
  strCurrentDirectory = PostOfficeDirectory + strCurrentUser + "\"
  strCurrentFile = strCurrentDirectory + lstFileList(intCurrentConnection).List(MessageIndex - 1)
  
  intCurrentFileNumber = FreeFile
  Open strCurrentFile For Input As #intCurrentFileNumber
  
  Do While Not EOF(intCurrentFileNumber)
    Line Input #intCurrentFileNumber, strSendLine
    intStatus = SendSocket(lngTransmitSocket, strSendLine)
  Loop
  
  Close #intCurrentFileNumber
  
  intStatus = SendSocket(lngTransmitSocket, ".")

End Sub

Private Sub DeleteMessages(intCurrentConnection As Integer)
  Dim strCurrentUser As String
  Dim strCurrentDirectory As String
  Dim intMessageIndex As Integer
  
  strCurrentUser = txtUser(intCurrentConnection).Text
  strCurrentDirectory = PostOfficeDirectory + strCurrentUser + "\"
 
  For intMessageIndex = 0 To val(txtTotalCount(intCurrentConnection).Text) - 1
    If lstDeletionList(intCurrentConnection).List(intMessageIndex) = "D" Then
      Kill strCurrentDirectory + lstFileList(intCurrentConnection).List(intMessageIndex)
      Debug.Print "Kill " + strCurrentDirectory + lstFileList(intCurrentConnection).List(intMessageIndex)
    End If
  Next
  
  intStatus = SendSocket(lngTransmitSocket, "+OK ")

End Sub

Private Sub cboConnection_Click()
  DisplayConnection cboConnection.ListIndex
End Sub

Private Sub DisplayConnection(NewConnection As Integer)

  If NewConnection < 0 Or NewConnection >= cboConnection.ListCount Then
    NewConnection = 0
  End If
  
  If NewConnection = intCurrentConnection Then Exit Sub
  
  txtRequest(intCurrentConnection).Visible = False
  txtRequest(NewConnection).Visible = True
  txtSocket(intCurrentConnection).Visible = False
  txtSocket(NewConnection).Visible = True
  txtState(intCurrentConnection).Visible = False
  txtState(NewConnection).Visible = True
  txtUser(intCurrentConnection).Visible = False
  txtUser(NewConnection).Visible = True
  txtTotalCount(intCurrentConnection).Visible = False
  txtTotalCount(NewConnection).Visible = True
  txtTotalSize(intCurrentConnection).Visible = False
  txtTotalSize(NewConnection).Visible = True
  lstDeletionList(intCurrentConnection).Visible = False
  lstDeletionList(NewConnection).Visible = True
  lstSizeList(intCurrentConnection).Visible = False
  lstSizeList(NewConnection).Visible = True
  lstFileList(intCurrentConnection).Visible = False
  lstFileList(NewConnection).Visible = True
 
  cboConnection.ListIndex = NewConnection
  
  intCurrentConnection = NewConnection
  
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  intStatus = ReleaseSocket(lngListenSocket)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup intStatus " & SocketError(intStatus)
End Sub


