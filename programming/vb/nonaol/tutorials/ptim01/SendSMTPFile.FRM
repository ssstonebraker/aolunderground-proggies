VERSION 5.00
Begin VB.Form frmSendSMTPFile 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Send SMTP File"
   ClientHeight    =   3390
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   7155
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
   ScaleHeight     =   3390
   ScaleWidth      =   7155
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox txtToRecipient 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   5775
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   5775
   End
   Begin VB.CommandButton btnDisplay 
      Caption         =   "&Display"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblTRecipient 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblFilename 
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmSendSMTPFile"
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

Dim intStatus As Integer
Const SMTPPort = 25

Private Sub Form_Load()
  intStatus = StartWinSock()
  txtFilename.Text = Command()
End Sub

Private Sub btnDisplay_Click()
  Dim intInputFileNumber As Integer
  Dim strInputLine As String
  
  Load frmReadMIMEMessages
  
  intInputFileNumber = FreeFile
  Open txtFilename.Text For Input As #intInputFileNumber
  Do While Not EOF(intInputFileNumber)
    Line Input #intInputFileNumber, strInputLine
    frmReadMIMEMessages.DisplayMessage.Lines.Add strInputLine
  Loop
  Close #intInputFileNumber

  frmReadMIMEMessages.Show

End Sub

Private Sub btnSend_Click()
  Dim lngSMTPSocket As Long
  Dim lngServerAddress As Long
  Dim strServerName As String
  
  Dim strServerResponse As String
  Dim strInputFile As String
  Dim intInputFileNumber As Integer
  Dim strInputLine As String
  
  lngSMTPSocket = 0
  
  strServerName = txtServer.Text
  intStatus = GetIPAddress(lngServerAddress, strServerName)
  intStatus = CreateSocket(lngSMTPSocket, 0)
  intStatus = ConnectSocket(lngSMTPSocket, lngServerAddress, SMTPPort)
  
  intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
  intStatus = SendSocket(lngSMTPSocket, "mail from:<" + txtFrom.Text + ">")
  
  intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
  intStatus = SendSocket(lngSMTPSocket, "rcpt to:<" + txtToRecipient.Text + ">")
  intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
  intStatus = SendSocket(lngSMTPSocket, "data")
  intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
  
  intInputFileNumber = FreeFile
  Open txtFilename.Text For Input As #intInputFileNumber
  Do While Not EOF(intInputFileNumber)
    Line Input #intInputFileNumber, strInputLine
    intStatus = SendSocket(lngSMTPSocket, strInputLine)
  Loop
  Close #intInputFileNumber
  
  intStatus = SendSocket(lngSMTPSocket, ".")
  intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
  intStatus = SendSocket(lngSMTPSocket, "quit")
  intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
  intStatus = ReleaseSocket(lngSMTPSocket)
  
End Sub

Private Sub btnQuit_Click()
  Unload Me
  End
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup intStatus " & SocketError(intStatus)
End Sub

