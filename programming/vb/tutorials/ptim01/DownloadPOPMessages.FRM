VERSION 5.00
Begin VB.Form frmDownloadPOPMessages 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "POP Downloader"
   ClientHeight    =   4635
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
   ScaleHeight     =   4635
   ScaleWidth      =   7755
   Begin VB.CheckBox chkDelete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete messages on server"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      TabIndex        =   8
      Text            =   "c:\Post Office\"
      Top             =   1320
      Width           =   5775
   End
   Begin VB.CommandButton btnDownload 
      Caption         =   "&Download"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ListBox lstMessageList 
      Height          =   2010
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   6975
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
   Begin VB.Label lblPath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   9
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
Attribute VB_Name = "frmDownloadPOPMessages"
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

Private Sub btnDownload_Click()
  Dim lngServerAddress As Long
  Dim strServerName As String
  Dim strSocketResponse As String
  Dim intMailCountEndposition As Integer
  Dim intMailCount As Integer
  Dim intMessageIndex As Integer
  Dim strFilename As String
  
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
  intMailCount = val(Mid(strSocketResponse, 5, intMailCountEndposition - 5))
  
  For intMessageIndex = 1 To intMailCount
    strFilename = "M" + Format(Now, "yymmddhhnnss") + Trim(Str(Int(Rnd * 1000))) + ".MIME"
    ExtractMessage intMessageIndex, strFilename
    If chkDelete.Value = 1 Then
      intStatus = SendSocket(lngPOPSocket, "DELE " + Str(intMessageIndex))
      intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
    End If
    lstMessageList.AddItem strFilename
  Next

  intStatus = SendSocket(lngPOPSocket, "QUIT ")
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)

End Sub

Private Sub ExtractMessage(MessageIndex As Integer, OutputFilename As String)
  Dim strSocketResponse As String
  Dim strFileSpecification As String

  Open txtPath + OutputFilename For Output As #1
  
  intStatus = SendSocket(lngPOPSocket, "RETR " + Str(MessageIndex))
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  
  intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  Do While strSocketResponse <> "."
    Debug.Print strSocketResponse
    Print #1, strSocketResponse
    intStatus = ReceiveSocket(lngPOPSocket, strSocketResponse)
  Loop

  Close #1

End Sub

Private Sub Form_Unload(CANCEL As Integer)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(intStatus)
End Sub
