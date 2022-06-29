VERSION 5.00
Begin VB.Form frmComposeMIMEMessage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Compose MIME Message"
   ClientHeight    =   6495
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
   ScaleHeight     =   6495
   ScaleWidth      =   7155
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Text            =   "Fri, 31 Dec 99 23:30 -0000"
      Top             =   240
      Width           =   5775
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   285
      Left            =   2160
      TabIndex        =   18
      Text            =   "jrr.mime"
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtContent 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   2040
      Width           =   6615
   End
   Begin VB.TextBox txtCCRecipient 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   1680
      Width           =   5775
   End
   Begin VB.TextBox txtToRecipient 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   600
      Width           =   5775
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtAttachment1 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   4440
      Width           =   3015
   End
   Begin VB.ComboBox cmbTransferEncoding1 
      Height          =   315
      Left            =   4680
      TabIndex        =   2
      Text            =   "Text"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtAttachment2 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   4800
      Width           =   3015
   End
   Begin VB.ComboBox cmbTransferEncoding2 
      Height          =   315
      Left            =   4680
      TabIndex        =   0
      Text            =   "Base64"
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblOutputFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Output File:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   6720
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblRecipient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Recipients:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblCC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblAttachment1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attachments:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblTransferEncoding 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Encoding:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
End
Attribute VB_Name = "frmComposeMIMEMessage"
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

Private Sub Form_Load()
    
  cmbTransferEncoding1.AddItem "Text"
  cmbTransferEncoding1.AddItem "Base64"
  cmbTransferEncoding1.AddItem "Quoted-Printable"
  
  cmbTransferEncoding2.AddItem "Text"
  cmbTransferEncoding2.AddItem "Base64"
  cmbTransferEncoding2.AddItem "Quoted-Printable"

End Sub

Private Sub btnSave_Click()
  Dim intOutputFileNumber As Integer
  Dim mmsgCurrentMessage As New MIMEMessage
  Dim strCurrentLine
  
  mmsgCurrentMessage.AddHeader "Date", txtDate.Text
  mmsgCurrentMessage.AddHeader "From", txtFrom.Text
  mmsgCurrentMessage.AddHeader "To", txtToRecipient.Text
  If txtCCRecipient.Text <> "" Then
    mmsgCurrentMessage.AddHeader "CC", txtCCRecipient.Text
  End If
  mmsgCurrentMessage.AddHeader "Subject", txtSubject.Text
  
  mmsgCurrentMessage.BodyLines.Add txtContent.Text
  
  If txtAttachment1.Text <> "" And Dir(txtAttachment1.Text) <> "" Then
    mmsgCurrentMessage.AddAttachment txtAttachment1.Text, cmbTransferEncoding1.Text
  End If
      
  If txtAttachment2.Text <> "" And Dir(txtAttachment2.Text) <> "" Then
    mmsgCurrentMessage.AddAttachment txtAttachment2.Text, cmbTransferEncoding2.Text
  End If
  
  mmsgCurrentMessage.Compose
  
  intOutputFileNumber = FreeFile
  Open txtOutputFile.Text For Output As #intOutputFileNumber
  For Each strCurrentLine In mmsgCurrentMessage.Lines
    Print #intOutputFileNumber, CStr(strCurrentLine)
  Next
  Close #intOutputFileNumber
  
End Sub
