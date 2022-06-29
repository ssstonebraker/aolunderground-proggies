VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmComposeRFC822Message 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Compose RFC822 Message"
   ClientHeight    =   5520
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
   ScaleHeight     =   5520
   ScaleWidth      =   7155
   Begin VB.CommandButton btnBrowse 
      Caption         =   "&Browse..."
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   4560
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog dlgOutputFile 
      Left            =   5880
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   240
      Width           =   5775
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Text            =   "jrr.rfc822"
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox txtContent 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2040
      Width           =   6615
   End
   Begin VB.TextBox txtCCRecipient 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Text            =   "Test message"
      Top             =   1680
      Width           =   5775
   End
   Begin VB.TextBox txtToRecipient 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   5775
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblOutputFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Output File:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   6720
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblRecipient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Recipients:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmComposeRFC822Message"
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


Private Sub btnSave_Click()
  Dim intOutputFileNumber As Integer
  Dim rmsgCurrentMessage As New RFC822Message
  Dim strCurrentLine
  
  rmsgCurrentMessage.AddHeader "Date", txtDate.Text
  rmsgCurrentMessage.AddHeader "From", txtFrom.Text
  rmsgCurrentMessage.AddHeader "To", txtToRecipient.Text
  If txtCCRecipient.Text <> "" Then
    rmsgCurrentMessage.AddHeader "CC", txtCCRecipient.Text
  End If
  rmsgCurrentMessage.AddHeader "Subject", txtSubject.Text
  
  rmsgCurrentMessage.BodyLines.Add txtContent.Text
  
  rmsgCurrentMessage.Compose
  
  intOutputFileNumber = FreeFile
  Open txtOutputFile.Text For Output As #intOutputFileNumber
  For Each strCurrentLine In rmsgCurrentMessage.Lines
    Print #intOutputFileNumber, CStr(strCurrentLine)
  Next
  Close #intOutputFileNumber
  
End Sub

Private Sub btnBrowse_Click()
  dlgOutputFile.DefaultExt = ".rfc822"
  dlgOutputFile.InitDir = CurDir
  dlgOutputFile.Filter = "*.rfc822"
  dlgOutputFile.ShowSave
  txtOutputFile.Text = dlgOutputFile.filename
End Sub


