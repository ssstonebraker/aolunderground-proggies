VERSION 5.00
Begin VB.Form frmASNComposer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "ASN Composer"
   ClientHeight    =   6420
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   9660
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
   ScaleHeight     =   6420
   ScaleWidth      =   9660
   Begin VB.CommandButton btnSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Text            =   "test.asn"
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   960
      TabIndex        =   15
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox cboValueType 
      Height          =   315
      Left            =   1320
      TabIndex        =   14
      Text            =   "Composite"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton btnUpdate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton btnDelete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtBinary 
      Height          =   5175
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtASN 
      Height          =   5175
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton btnAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ListBox lstRecords 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtDepth 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblBinary 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Binary:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblASN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ASN:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Value:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblValueType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Value Type:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblTag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tag:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblDepth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Depth:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmASNComposer"
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
Dim MainArray()

Private Sub Form_Load()

  cboValueType.AddItem "Numeric"
  cboValueType.AddItem "String"
  cboValueType.AddItem "Composite"
  
End Sub

Private Sub btnAdd_Click()
  Dim strAddLine As String

  strAddLine = Right(Space(4) + txtDepth.Text, 4) + _
            Right(Space(4) + txtTag.Text, 4) + _
            Right(Space(2) + Left(cboValueType.Text, 1), 2) + _
            " " + txtValue.Text
  
  If lstRecords.ListIndex = -1 Then
    lstRecords.AddItem strAddLine
  Else
    lstRecords.AddItem strAddLine, lstRecords.ListIndex
  End If
  
  RefreshForm
  
End Sub

Private Sub btnDelete_Click()
  
  If lstRecords.ListIndex <> -1 Then
    lstRecords.RemoveItem lstRecords.ListIndex
  End If
  
  RefreshForm
  
End Sub

Private Sub btnUpdate_Click()
  If lstRecords.ListIndex <> -1 Then
    lstRecords.List(lstRecords.ListIndex) = _
            Right(Space(4) + txtDepth.Text, 4) + _
            Right(Space(4) + txtTag.Text, 4) + _
            Right(Space(2) + Left(cboValueType.Text, 1), 2) + _
            " " + txtValue.Text
  End If

  RefreshForm

End Sub

Private Sub btnSave_Click()
  Dim intOutputFileNumber As Integer

  intOutputFileNumber = FreeFile
  Open txtFilename.Text For Binary As #intOutputFileNumber
  Put #intOutputFileNumber, , asnMain.TransferString
  Close #intOutputFileNumber
  
End Sub

Private Sub RefreshForm()
  Dim strTransferString As String
  Dim intScanIndex As Integer

  FillArray MainArray, 0
  
  asnMain.Compose MainArray
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

Private Sub FillArray(ByRef WorkArray, ByRef LineIndex As Integer)
  Dim intArrayIndex As Integer
  Dim intLineDepth As Integer
  Dim varSubordinateArray() As Variant

  ReDim Preserve WorkArray(2)
  WorkArray(1) = val(Mid(lstRecords.List(LineIndex), 5, 4))
  
  intArrayIndex = 1
  intLineDepth = GetDepth(LineIndex)
  
  Select Case Mid(lstRecords.List(LineIndex), 10, 1)
    Case "N"
      WorkArray(2) = Chr(Mid(lstRecords.List(LineIndex), 12))
    Case "S"
      WorkArray(2) = Mid(lstRecords.List(LineIndex), 12)
    Case "C"
      Do While GetDepth(LineIndex + 1) > intLineDepth
        intArrayIndex = intArrayIndex + 1
        LineIndex = LineIndex + 1
        ReDim Preserve WorkArray(intArrayIndex)
        FillArray varSubordinateArray, LineIndex
        WorkArray(intArrayIndex) = varSubordinateArray
      Loop
  End Select
  
End Sub

Private Function GetDepth(LineIndex As Integer) As Integer
  If LineIndex <= lstRecords.ListCount Then
    GetDepth = val(Left(lstRecords.List(LineIndex), 4))
  Else
    GetDepth = -1
  End If
End Function

Private Sub lstRecords_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    lstRecords.ListIndex = -1
  End If
End Sub
