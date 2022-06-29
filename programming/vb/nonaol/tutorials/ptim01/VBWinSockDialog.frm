VERSION 5.00
Begin VB.Form frmVBWinSockDialog 
   BackColor       =   &H00C0C0C0&
   Caption         =   "VB WinSock Dialog"
   ClientHeight    =   4245
   ClientLeft      =   9600
   ClientTop       =   7095
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5535
   Begin VB.CommandButton Save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtDialog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmVBWinSockDialog"
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

Private Sub Form_Resize()

  txtDialog.Top = 100
  txtDialog.Left = 100
  
  If Me.Height > 1000 Then txtDialog.Height = Me.Height - 1000
  If Me.Width > 600 Then txtDialog.Width = Me.Width - 300

End Sub

Private Sub btnClear_Click()
  txtDialog.Text = ""
End Sub

Private Sub Save_Click()
  Dim FileSpecification As String
  Dim intOutputFileNumber As Integer

  intOutputFileNumber = FreeFile
  

  FileSpecifcation = InputBox("Please enter file: ")
  Open FileSpecifcation For Output As #intOutputFileNumber
  Print #intOutputFileNumber, txtDialog.Text
  Close #intOutputFileNumber

End Sub
