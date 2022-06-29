VERSION 5.00
Begin VB.Form frmReadMIMEMessages 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "MIME Reader"
   ClientHeight    =   5925
   ClientLeft      =   2175
   ClientTop       =   1875
   ClientWidth     =   11670
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
   ScaleHeight     =   5925
   ScaleWidth      =   11670
   Begin VB.Frame grpComponents 
      BackColor       =   &H00C0C0C0&
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtCC 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   12
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   5775
      End
      Begin VB.ListBox lstAttachmentList 
         Height          =   840
         Left            =   120
         TabIndex        =   5
         Top             =   4680
         Width           =   6975
      End
      Begin VB.TextBox txtContent 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   2040
         Width           =   6975
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   3
         Top             =   1680
         Width           =   5775
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label lblCC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CC:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblSubject 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtSource 
      Height          =   5565
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmReadMIMEMessages"
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
Dim CRLF As String

Public DisplayMessage As New MIMEMessage

Private Sub Form_Load()
  CRLF = Chr(13) + Chr(10)
End Sub

Private Sub Form_Activate()
  LoadMessage
End Sub

Private Sub LoadMessage()
  Dim strCurrentLine
  Dim mhdrHeaderLine
  Dim mattAttachment
  Dim strBodyLine
  
  txtSource.Text = ""
  For Each strCurrentLine In DisplayMessage.Lines
    txtSource.Text = txtSource.Text + CRLF + strCurrentLine
  Next

  If Len(txtSource.Text) > 2 Then txtSource.Text = Mid(txtSource.Text, 3)

  DisplayMessage.Parse
  
  For Each mhdrHeaderLine In DisplayMessage.Headers
    Select Case LCase(mhdrHeaderLine.Label)
      Case "from":
        txtFrom.Text = mhdrHeaderLine.Value
      Case "to":
        txtTo.Text = mhdrHeaderLine.Value
      Case "cc":
        txtCC.Text = mhdrHeaderLine.Value
      Case "subject":
        txtSubject.Text = mhdrHeaderLine.Value
      Case "date":
        txtDate.Text = mhdrHeaderLine.Value
    End Select
  Next
  
  For Each strBodyLine In DisplayMessage.BodyLines
    txtContent.Text = txtContent.Text + CRLF + strBodyLine
  Next
  
  If Len(txtContent.Text) > 2 Then txtContent.Text = Mid(txtContent.Text, 3)
  
  For Each mattAttachment In DisplayMessage.Attachments
    Debug.Print mattAttachment.Filename
    lstAttachmentList.AddItem mattAttachment.Filename
  Next
  
End Sub

Private Sub lstAttachmentList_DblClick()
  DisplayMessage.Attachments(lstAttachmentList.ListIndex + 1).CallAttachment
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  Set DisplayMessage = Nothing
End Sub
