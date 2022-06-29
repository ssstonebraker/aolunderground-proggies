VERSION 5.00
Begin VB.Form frmGetPOPMessages 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "POP Client"
   ClientHeight    =   4545
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   7755
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
   ScaleHeight     =   4545
   ScaleWidth      =   7755
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton btnExtract 
      Caption         =   "&Extract"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtMailSize 
      Height          =   285
      Left            =   5880
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtMailCount 
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton btnQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton btnShow 
      Caption         =   "&Show"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ListBox lstMessageList 
      Height          =   1425
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label lblMailSize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Size (Bytes)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblMailCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Count"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblMessages 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Messages"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmGetPOPMessages"
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

Dim lngPOPSocket As Long
Const POP3Port = 110
Dim CRLF As String
Dim intStatus As Integer


Private Sub Form_Activate()
  txtPassword.SetFocus
End Sub

Private Sub Form_Load()
  CRLF = Chr(13) + Chr(10)
  intStatus = StartWinSock()
End Sub

Private Sub btnConnect_Click()
  Dim lngServerAddress As Long
  Dim strServerName As String
  Dim strSocketResponse As String
  Dim intMailCountEndposition As Integer
  Dim intMailSizeEndPosition As Integer
  
  lngPOPSocket = 0

  strServerName = txtServer.Text
  intStatus = GetIPAddress(lngServerAddress, strServerName)
  intStatus = CreateSocket(lngPOPSocket, 0)
  
  frmVBWinSockDialog.Left = Me.Left
  frmVBWinSockDialog.Top = Me.Height
  
  intStatus = ConnectSocket(lngPOPSocket, lngServerAddress, POP3Port)
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  intStatus = SendSocket(lngPOPSocket, "USER " + txtUsername.Text)
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  intStatus = SendSocket(lngPOPSocket, "PASS " + txtPassword.Text)
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  intStatus = SendSocket(lngPOPSocket, "STAT ")
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  intMailCountEndposition = InStr(5, strSocketResponse, " ")
  txtMailCount.Text = Mid(strSocketResponse, 5, intMailCountEndposition - 5)
  txtMailSize.Text = Mid(strSocketResponse, intMailCountEndposition + 1)
  
  
  intStatus = SendSocket(lngPOPSocket, "LIST ")
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  Do While InStr(strSocketResponse, ".") = 0
    lstMessageList.AddItem strSocketResponse
    intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  Loop
  
  btnShow.Enabled = True
  btnExtract.Enabled = True
  btnDelete.Enabled = True
  btnConnect.DEFAULT = False
  btnShow.DEFAULT = True
    
End Sub

Private Sub btnShow_Click()
  Dim strSocketResponse As String
 
  Load frmReadMIMEMessages
  
  intStatus = SendSocket(lngPOPSocket, "RETR " + Str(lstMessageList.ListIndex + 1))
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  Do While strSocketResponse <> "."
    Debug.Print strSocketResponse
    frmReadMIMEMessages.DisplayMessage.Lines.Add strSocketResponse
    intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  Loop
  
  frmReadMIMEMessages.Show
    
End Sub

Private Sub btnExtract_Click()
  
  Dim strSocketResponse As String
  Dim strFileSpecification As String
  
  strFileSpecification = InputBox("Please enter file: ")
  Open strFileSpecification For Output As #1
  
  intStatus = SendSocket(lngPOPSocket, "RETR " + Str(lstMessageList.ListIndex + 1))
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  Do While strSocketResponse <> "."
    Debug.Print strSocketResponse
    Print #1, strSocketResponse
    intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  Loop

  Close #1

End Sub

Private Sub btnDelete_Click()
  Dim strSocketResponse As String
  Dim intRetrieveIndex As Integer
  
  intRetrieveIndex = lstMessageList.ListIndex + 1

  intStatus = SendSocket(lngPOPSocket, "DELE " + Str(intRetrieveIndex))
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  Debug.Print strSocketResponse

End Sub

Private Sub btnQuit_Click()
  Dim strSocketResponse As String
  
  intStatus = SendSocket(lngPOPSocket, "QUIT ")
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  Unload Me
  End
  
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  
  intStatus = ReleaseSocket(lngPOPSocket)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(intStatus)

End Sub
