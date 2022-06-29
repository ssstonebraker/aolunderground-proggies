VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto-Correct By Dexter"
   ClientHeight    =   3990
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   5595
   End
   Begin VB.Label Label5 
      Caption         =   "Info:"
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   3060
      Width           =   2565
   End
   Begin VB.Label Label4 
      Caption         =   "Special Chars"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3330
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Current Word"
      Height          =   255
      Left            =   1650
      TabIndex        =   3
      Top             =   3330
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1650
      TabIndex        =   1
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Item List"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Auto Text Correct ßy Dexter
'Correct text while you type.  The same as Word does
'Created: December 05, 2001

'Email:  DexterXCT@aol.com

'AutoCorrect.msf is the word list
'Word list is setup like:
'(Total Number in list)  chr$(3) wrong word chr$(1) correct word chr$(2) wrong chr$(1) correct chr$(2) ...

Private Function LoadAutoFormat()
  Dim lngLength As Long, lngFind As Long, lngStart As Long
  Dim strBuffer As String, strFind As String, strReplace As String
  Dim i As Integer, intCount As Integer
  Dim ItmX As ListItem
  Open App.Path & "\AutoFormat.msf" For Binary As #1
  lngLength& = LOF(1)
  If lngLength& = 0 Then
    Close #1
    Exit Function
  End If
  strBuffer$ = String(lngLength&, 1)
  Get #1, 1, strBuffer$
  Close #1
  lngFind& = InStr(strBuffer$, Chr$(3))  'Get total number of words in list.
  intCount% = CInt(Mid(strBuffer$, 1, lngFind& - 1))
  strBuffer$ = Mid(strBuffer$, lngFind& + 1, Len(strBuffer$)) 'Set first word at start of file
  Form2.Label3.Caption = "Total Words:  " & CStr(intCount)
  For i = 1 To intCount 'Loop through entire list
    lngFind& = InStr(strBuffer$, Chr$(1)) 'Find end of first word
    lngStart& = InStr(strBuffer$, Chr$(2)) 'find end of second word
    strFind = Mid(strBuffer$, 1, lngFind& - 1) 'Set as first word
    strReplace = Mid(strBuffer$, lngFind& + 1, lngStart& - lngFind& - 1) 'Set as second word
    strBuffer$ = Mid(strBuffer$, lngStart& + 1, Len(strBuffer$)) 'Remove word from list
    Set ItmX = Form2.ListView1.ListItems.Add(, , strFind) 'Set the first word in the first position
            'of the list view
    ItmX.SubItems(1) = strReplace 'Set second word in the second position
            'of the list view
  Next i
End Function

Private Sub Form_Load()
  Call LoadAutoFormat
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  End
End Sub

Private Sub mnuAbout_Click()
   MsgBox "Written By: " & Chr$(9) & "Dexter" & vbCrLf & _
              "Email: " & Chr$(9) & Chr$(9) & "DexterXCT@aol.com" & vbCrLf & _
              "Version: " & Chr$(9) & Chr$(9) & "0.1" & vbCrLf & _
              "            " & Chr$(9) & Chr$(9) & "December 05, 2001", _
              vbInformation, _
              "Auto-Correct By Dexter"
End Sub

Private Sub mnuFileItem_Click()
   Form2.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim FndX As ListItem
  If KeyAscii = 32 Or KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii = 58 Or KeyAscii = 59 Or KeyAscii = 33 Or KeyAscii = 63 Or KeyAscii = 44 Then
    'If next letter is any type of word break.
    'Such as a space ' ', '.' , '!', ect.
    
    If Len(Label1.Caption) < 1 Then Exit Sub
    'If the buffer label is empty, don't do anything
    
    If Not Form2.ListView1.FindItem(Label1.Caption & Chr$(KeyAscii), , , 1) Is Nothing Then
      'If buffer label is partcialy found in the list
      'then keep adding the the word
      Label1.Caption = Label1.Caption & Chr$(KeyAscii)
    Else
      Set FndX = Form2.ListView1.FindItem(Label1.Caption) 'search through wrong words to see if there is a match
      If Not FndX Is Nothing Then 'If theres a match then replace word.
      'FndX is returned as nothing if the word is not found
      'I used Not/Nothing so i didn't have to use an Else
        Me.Text1.SelStart = Me.Text1.SelStart - Len(Label1.Caption)
        Me.Text1.SelLength = Len(Label1.Caption)
        Me.Text1.SelText = FndX.ListSubItems(1)
      End If
      Label1.Caption = ""  'Clear word
    End If
  ElseIf KeyAscii = 8 And Not Len(Label1.Caption) = 0 Then
  'If keypress is a backspace, delete the last letter in
  'the buffer label.  But don't delete if the label is empty
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 1)
  ElseIf KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
    'do something...
  
  ElseIf KeyAscii < 32 Then
    Label2.Caption = ""
    'do nothing
  Else
      Label1.Caption = Label1.Caption & Chr$(KeyAscii)
      'Keep adding letters to the buffer label until
      'the word is found or until the buffer label
      'doesn't match anything
  End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim FndX As ListItem
  'This is for correcting special characters.
  'Such as a "."
  '... is changed to …
  If KeyCode = 190 Or KeyCode = 189 Then ' Or KeyCode = 46 Or KeyCode = 58 Or KeyCode = 59 Or KeyCode = 33 Or KeyCode = 63 Then
    Label2.Caption = Label2.Caption & Chr$(KeyCode - 144)
    If Not Form2.ListView1.FindItem(Label2.Caption) Is Nothing Then
      Set FndX = Form2.ListView1.FindItem(Label2.Caption)
      Me.Text1.SelStart = Me.Text1.SelStart - Len(Label2.Caption)
      Me.Text1.SelLength = Len(Label2.Caption)
      Me.Text1.SelText = FndX.ListSubItems(1)
      Label2.Caption = ""
    End If
  Else
    Label2.Caption = ""
  End If
End Sub

