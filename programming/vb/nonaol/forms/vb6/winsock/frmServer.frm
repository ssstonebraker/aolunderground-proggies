VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   5280
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   5895
      Begin VB.Frame Frame2 
         Height          =   500
         Left            =   60
         TabIndex        =   6
         Top             =   115
         Width           =   5775
         Begin VB.CheckBox chkListen 
            Caption         =   "Listen For Connections"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   295
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   150
            Width           =   5535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Chat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2610
         Left            =   60
         TabIndex        =   4
         Top             =   650
         Width           =   5775
         Begin VB.TextBox txtChat 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2275
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   5590
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   60
         TabIndex        =   1
         Top             =   3250
         Width           =   5775
         Begin VB.TextBox txtMessage 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   3
            Top             =   140
            Width           =   4815
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5000
            TabIndex        =   2
            Top             =   150
            Width           =   655
         End
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'_____________________________________'
'                                     '
'this is a vb example on how to use   '
'winsock to make a simple client &    '
'server based application...for your  '
'questions & comments please email me '
'at webmaster@plastik.zzn.com..Thanks!'
'_____________________________________'
'

Private Sub chkListen_Click()
Select Case chkListen.Value
 Case "1"
  tcpServer.LocalPort = 2223
  tcpServer.Listen
 Case "0"
  tcpServer.Close
End Select
End Sub

Private Sub cmdSend_Click()
   ' send data to the client
   Call tcpServer.SendData("SERVER >>> " & txtMessage.Text)
   txtChat.Text = txtChat.Text & _
      "SERVER >>> " & txtMessage.Text & vbCrLf & vbCrLf
   txtMessage.Text = ""
   txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub Form_Terminate()
tcpServer.Close 'if form terminates then close the server
End Sub

Private Sub tcpServer_Close()
'if the client closes connection so should the server
cmdSend.Enabled = False
txtMessage.Enabled = False
Call tcpServer.Close
txtChat.Text = txtChat.Text & "Client closed the connection." & vbCrLf
txtChat.SelStart = Len(txtChat.Text)
tcpServer.Listen 'listen for another connection
End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
'close the server if it is open before accepting
'the connection request
If tcpServer.State <> sckClosed Then
 tcpServer.Close
End If
 cmdSend.Enabled = True
 txtMessage.Enabled = True
 Call tcpServer.Accept(requestID) 'accepts connection
 'display connection status
 txtChat.Text = "Connection from IP address: " & _
      tcpServer.RemoteHostIP & vbCrLf & "Port #: " & _
      tcpServer.RemotePort & vbCrLf & vbCrLf
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
Dim strMessage As String
'converts the data into a string called strMessage$
Call tcpServer.GetData(strMessage$)
txtChat.Text = txtChat.Text & strMessage$ & vbCrLf
txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub tcpServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   Dim result As Integer
   result = MsgBox(Source & ": " & Description, _
      vbOKOnly, "TCP/IP Error")
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call cmdSend_Click
End If
End Sub
