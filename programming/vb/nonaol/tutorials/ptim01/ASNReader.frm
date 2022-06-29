VERSION 5.00
Begin VB.Form frmASNReader 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "ASN Reader"
   ClientHeight    =   6105
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   6645
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
   ScaleHeight     =   6105
   ScaleWidth      =   6645
   Begin VB.CommandButton btnRead 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Read"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "test.asn"
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtBinary 
      Height          =   4095
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtASN 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lblFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblBinary 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Binary:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblASN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ASN:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmASNReader"
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
Option Base 1
Option Compare Text

Dim asnMain As New ASN1

Private Sub btnRead_Click()
  Dim intInputFileNumber As Integer
  Dim strInputString As String * 10000

  intInputFileNumber = FreeFile
  Open txtFilename.Text For Binary As #intInputFileNumber
  Get #intInputFileNumber, , strInputString
  Close #intInputFileNumber

  asnMain.Parse strInputString

  RefreshForm
  
End Sub

Private Sub RefreshForm()
  Dim strTransferString As String
  Dim intScanIndex As Integer

  txtASN.Text = asnMain.Dump
  strTransferString = asnMain.TransferString
  
  txtBinary.Text = ""
  For intScanIndex = 1 To Len(strTransferString)
    Debug.Print Right(Space(3) + Str(Asc(Mid(strTransferString, intScanIndex, 1))), 4)
    txtBinary.Text = txtBinary.Text + Right(Space(3) + Str(Asc(Mid(strTransferString, intScanIndex, 1))), 4)
    If Int(intScanIndex / 4) = intScanIndex / 4 Then
      txtBinary.Text = txtBinary.Text + Chr(13) + Chr(10)
    End If
  Next
  
  txtASN.Refresh
  txtBinary.Refresh
  
End Sub

