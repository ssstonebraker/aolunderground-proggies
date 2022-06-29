VERSION 5.00
Begin VB.Form frmSMTPServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "SMTP Server"
   ClientHeight    =   5340
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
   ScaleHeight     =   5340
   ScaleWidth      =   7470
   Begin VB.TextBox txtTo 
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   10
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtMessage 
      Height          =   3735
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   1440
      Width           =   7095
   End
   Begin VB.TextBox txtRequest 
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtConnect 
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox cboConnection 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtSocket 
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sender:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Connection:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
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
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "frmSMTPServer"
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
Const SMTPPort = 25
Const PostOfficeDirectory = "C:\Post Office\"

Dim intCurrentConnection As Integer
Dim CRLF As String

Private Sub Form_Load()
  Dim lngListenSocketAddress As InputSocketDescriptor

  intStatus = StartWinSock()
  
  CRLF = Chr(13) + Chr(10)
  
  cboConnection.AddItem ""
    
  txtRequest(0).Text = ""
  txtSocket(0).Text = ""
  txtMessage(0).Text = ""
  txtSender(0).Text = ""
  txtFrom(0).Text = ""
  txtTo(0).Text = ""
    
  
  intStatus = CreateSocket(lngListenSocket, SMTPPort)
  
  If listen(lngListenSocket, 5) Then
      MsgBox "Could not listen on Port " + Str$(SMTPPort)
      End
  End If
  
  If WSAAsyncSelect(lngListenSocket, txtConnect.hwnd, WM_MBUTTONUP, FD_ACCEPT) Then
      MsgBox "Unable to set Asynch mode"
  End If

End Sub

  ' Accept incoming connection
Private Sub txtConnect_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim lngTransmitSocketAddress As OutputSocketDescriptor
  Dim lngTransmitSocketAddressSize As Integer
  Dim intCurrentConnection As Integer
  
  lngTransmitSocketAddressSize = LenB(lngTransmitSocketAddress)
  lngTransmitSocketAddressSize = 20
  
  lngTransmitSocket = accept(lngListenSocket, lngTransmitSocketAddress, lngTransmitSocketAddressSize)

  intCurrentConnection = cboConnection.ListCount
    
  cboConnection.AddItem Trim(Str(intCurrentConnection))
  Load txtSocket(intCurrentConnection)
  Load txtRequest(intCurrentConnection)
  Load txtMessage(intCurrentConnection)
  Load txtSender(intCurrentConnection)
  Load txtFrom(intCurrentConnection)
  Load txtTo(intCurrentConnection)
  
  txtSocket(intCurrentConnection) = Trim(Str(lngTransmitSocket))
  
  DisplayConnection intCurrentConnection
  
  intStatus = SendSocket(lngTransmitSocket, "220 SMTP server here")
    
  If WSAAsyncSelect(lngTransmitSocket, txtRequest(intCurrentConnection).hwnd, WM_MBUTTONDBLCLK, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If

    
End Sub

  ' Process incoming request
Private Sub txtRequest_DblClick(Index As Integer)
  Dim strClientRequest As String
  
  intStatus = ReceiveSocket(lngTransmitSocket, strClientRequest)
  DisplayConnection Index
  txtRequest(Index).Text = strClientRequest
  
  strClientRequest = Trim(strClientRequest)
  Select Case Trim(UCase(Left$(strClientRequest, 4)))
    Case "HELO"
      txtSender(Index).Text = Trim(Mid$(strClientRequest, 5))
      intStatus = SendSocket(lngTransmitSocket, "250 olive.tree")
    Case "MAIL"
      txtFrom(Index).Text = Trim(Mid$(strClientRequest, 11))
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
    Case "RCPT"
      txtTo(Index).Text = Trim(Mid$(strClientRequest, 9))
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
    Case "DATA"
      ReceiveMessage Index
    Case "QUIT"
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
      intStatus = ReleaseSocket(lngTransmitSocket)
    Case Else
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
  End Select

End Sub

Private Sub ReceiveMessage(Index As Integer)
  Dim strClientRequest As String
  Dim strFileName As String
  Dim intFileNumber As Integer
  Dim strUserName As String
  
  txtMessage(Index).Text = ""

  strUserName = Left(txtTo(Index).Text, InStr(1, txtTo(Index).Text, "@") - 1)
  strUserName = Mid(strUserName, 2)
  
  intFileNumber = FreeFile
  strFileName = PostOfficeDirectory + strUserName + "\" + RandomName + ".MIME"
  Open strFileName For Output As #intFileNumber

  intStatus = SendSocket(lngTransmitSocket, "354 Start sending...")
  Do
    intStatus = ReceiveSocket(lngTransmitSocket, strClientRequest)
    If intStatus Then
      If strClientRequest = "." Then Exit Do
      txtMessage(Index).Text = txtMessage(Index).Text + strClientRequest + CRLF
      Print #intFileNumber, strClientRequest
    End If
  Loop 'While Right(strClientRequest, 3) <> CRLF + "."

  Close #intFileNumber
  intStatus = SendSocket(lngTransmitSocket, "250 OK")

End Sub

Private Function RandomName() As String
  RandomName = "M" + Format(Now, "yymmddhhnnss") + Trim(Str(Int(Rnd * 10000)))
End Function


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
  txtMessage(intCurrentConnection).Visible = False
  txtMessage(NewConnection).Visible = True
  txtSender(intCurrentConnection).Visible = False
  txtSender(NewConnection).Visible = True
  txtFrom(intCurrentConnection).Visible = False
  txtFrom(NewConnection).Visible = True
  txtTo(intCurrentConnection).Visible = False
  txtTo(NewConnection).Visible = True
      
  cboConnection.ListIndex = NewConnection
  
  intCurrentConnection = NewConnection
  
End Sub

Private Sub Form_Unload(CANCEL As Integer)
   
  intStatus = ReleaseSocket(lngListenSocket)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup intStatus " & SocketError(intStatus)

End Sub
