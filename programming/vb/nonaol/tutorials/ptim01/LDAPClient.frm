VERSION 5.00
Begin VB.Form frmLDAPClient 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "LDAP Client"
   ClientHeight    =   5835
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   8535
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
   ScaleHeight     =   5835
   ScaleWidth      =   8535
   Begin VB.TextBox txtASNDialog 
      Height          =   5535
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton btnSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtSearchString 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   3135
   End
   Begin VB.ListBox lstAttributes 
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   4695
   End
   Begin VB.ListBox lstMatches 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblAttributes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attributes:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblMatches 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Matches:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblSearchString 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search String:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmLDAPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Option Compare Text

Dim intStatus As Integer

Const LDAPPort = 389
Dim lngLDAPSocket As Long
Const LDAP_SIZELIMIT = 10
Const LDAP_TIMELIMIT = 30
Const LDAP_VERSION = 3
Const LDAP_BASEOBJECT = "" '"C=*"

Dim CRLF As String
Dim intRequestMessageID As Integer

Private Sub Form_Load()
  intStatus = StartWinSock()
  CRLF = Chr(13) + Chr(10)
End Sub

Private Sub btnSearch_Click()
  Dim lngServerAddress As Long
  Dim strServerName As String
  
  lngLDAPSocket = 0
  
  strServerName = txtServer.Text
  intStatus = GetIPAddress(lngServerAddress, strServerName)
  intStatus = CreateSocket(lngLDAPSocket, 0)
  intStatus = ConnectSocket(lngLDAPSocket, lngServerAddress, LDAPPort)
  
  intStatus = SendBindRequest()
  intStatus = SendSearchRequest(txtSearchString.Text)
  intStatus = SendUnBindRequest()
  
  intStatus = ReleaseSocket(lngLDAPSocket)
  
End Sub

Private Function SendBindRequest() As Integer
  Dim asnBindRequest As New ASN1
  Dim asnBindResponse As New ASN1
  Dim strServerResponse As String

  intRequestMessageID = intRequestMessageID + 1
      
  asnBindRequest.Compose ( _
    Array(ASN_SEQUENCE_TAG, _
      Array(ASN_INTEGER_TAG, Chr(intRequestMessageID)), _
      Array(LDAP_BINDREQUEST_TAG, _
        Array(ASN_INTEGER_TAG, Chr(LDAP_VERSION)), _
        Array(ASN_OCTETSTRING_TAG, ""), _
        Array(ASN_CONTEXT_SPECIFIC, ""))))

  txtASNDialog.Text = txtASNDialog.Text + asnBindRequest.Dump

  intStatus = SendSocketBinary(lngLDAPSocket, asnBindRequest.TransferString)
  intStatus = ReceiveSocketBinary(lngLDAPSocket, strServerResponse)
  
  asnBindResponse.Parse (strServerResponse)
  txtASNDialog.Text = txtASNDialog.Text + asnBindResponse.Dump
  txtASNDialog.Refresh
  
End Function

Private Function SendSearchRequest(SearchString As String) As Integer
  Dim asnRequest As New ASN1
  Dim asnResponse As New ASN1
  Dim strServerResponse As String
  Dim strServerNextResponse As String
  Dim intMatchIndex As Integer
  Dim intAttributeIndex As Integer
  Dim asnAttribute As New ASN1
  

  intRequestMessageID = intRequestMessageID + 1
      
  asnRequest.Compose ( _
    Array(ASN_SEQUENCE_TAG, _
      Array(ASN_INTEGER_TAG, Chr(intRequestMessageID)), _
      Array(LDAP_SEARCHREQUEST_TAG, _
        Array(ASN_OCTETSTRING_TAG, LDAP_BASEOBJECT), _
        Array(ASN_ENUMERATED_TAG, Chr(LDAP_SCOPE_WHOLESUBTREE)), _
        Array(ASN_ENUMERATED_TAG, Chr(LDAP_DEREFALIASES_ALWAYS)), _
        Array(ASN_INTEGER_TAG, Chr(LDAP_SIZELIMIT)), _
        Array(ASN_INTEGER_TAG, Chr(LDAP_TIMELIMIT)), _
        Array(ASN_BOOLEAN_TAG, Chr(LDAP_FALSE)), _
        Array(LDAP_SEARCH_SUBSTRINGS, _
          Array(ASN_OCTETSTRING_TAG, "cn"), _
          Array(ASN_SEQUENCE_TAG, _
            Array(ASN_CONTEXT_SPECIFIC, SearchString))), _
        Array(ASN_SEQUENCE_TAG, _
          Array(ASN_OCTETSTRING_TAG, "cn"), _
          Array(ASN_OCTETSTRING_TAG, "mail")))))

  txtASNDialog.Text = txtASNDialog.Text + asnRequest.Dump

  intStatus = SendSocketBinary(lngLDAPSocket, asnRequest.TransferString)
  
  intStatus = ReceiveSocketBinary(lngLDAPSocket, strServerResponse)
  Debug.Print Len(strServerResponse)
  asnResponse.Parse strServerResponse
  
  Debug.Print Len(strServerResponse)
  txtASNDialog.Text = txtASNDialog.Text + asnResponse.Dump
  
  For intMatchIndex = 1 To lstAttributes.UBound
    Unload lstAttributes(intMatchIndex)
  Next
  lstMatches.Clear
  
  intMatchIndex = 0
  Do While asnResponse.SubItem(2).Tag = LDAP_SEARCHRESPONSEENTRY_TAG
  
    DoEvents
    Debug.Print asnResponse.SubItem(2).SubItem(1).Value
    lstMatches.AddItem asnResponse.SubItem(2).SubItem(1).Value
  
    intMatchIndex = intMatchIndex + 1
    Load lstAttributes(intMatchIndex)
    Set asnAttribute = asnResponse.SubItem(2).SubItem(2)
    For intAttributeIndex = 1 To asnAttribute.SubItemCount
      lstAttributes(intMatchIndex).AddItem _
        asnAttribute.SubItem(intAttributeIndex).SubItem(1).Value + " " + _
        asnAttribute.SubItem(intAttributeIndex).SubItem(2).SubItem(1).Value
    Next
  
    asnResponse.Parse strServerResponse
    
    txtASNDialog.Text = txtASNDialog.Text + asnResponse.Dump
  
  Loop
  
  ShowMatch (intMatchIndex)
  lstMatches.ListIndex = intMatchIndex - 1
  txtASNDialog.Refresh

End Function

Private Function SendUnBindRequest() As Integer

End Function

Private Sub lstMatches_Click()
  ShowMatch (lstMatches.ListIndex + 1)
End Sub

Private Sub lstMatches_KeyPress(KeyAscii As Integer)
  ShowMatch (lstMatches.ListIndex + 1)
End Sub

Private Sub lstMatches_LostFocus()
  lstMatches.ToolTipText = ""
End Sub

Private Sub lstMatches_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  ShowMatch (lstMatches.ListIndex + 1)
End Sub

Private Sub ShowMatch(MatchIndex As Integer)
  Static intOldMatchIndex As Integer

  If intOldMatchIndex > lstAttributes.UBound Then intOldMatchIndex = lstAttributes.UBound

  lstAttributes(intOldMatchIndex).Visible = False
  lstAttributes(MatchIndex).Visible = True
  
  intOldMatchIndex = MatchIndex
  
  lstMatches.ToolTipText = lstMatches.Text
  
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  
  intStatus = ReleaseSocket(lngLDAPSocket)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(intStatus)

End Sub


