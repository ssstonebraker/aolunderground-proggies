VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "View List"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   3660
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2985
      Left            =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   5595
   End
   Begin VB.Label Label3 
      Caption         =   "Current Word"
      Height          =   255
      Left            =   2460
      TabIndex        =   3
      Top             =   3390
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   930
      TabIndex        =   2
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   3660
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function LoadAutoFormat()
  Dim lngLength As Long, lngFind As Long, lngStart As Long
  Dim strBuffer As String, strFind As String, strReplace As String
  Dim i As Integer, intCount As Integer
  Dim ItmX As ListItem
  Open "E:\MacroShop\Program Files\AutoFormat.msf" For Binary As #1
  lngLength& = LOF(1)
  If lngLength& = 0 Then
    Close #1
    Exit Function
  End If
  strBuffer$ = String(lngLength&, 1)
  Get #1, 1, strBuffer$
  Close #1
  lngFind& = InStr(strBuffer$, Chr$(3))
  intCount% = CInt(Mid(strBuffer$, 1, lngFind& - 1))
  strBuffer$ = Mid(strBuffer$, lngFind& + 1, Len(strBuffer$))
  For i = 1 To intCount
    lngFind& = InStr(strBuffer$, Chr$(1))
    lngStart& = InStr(strBuffer$, Chr$(2))
    strFind = Mid(strBuffer$, 1, lngFind& - 1)
    strReplace = Mid(strBuffer$, lngFind& + 1, lngStart& - lngFind& - 1)
    strBuffer$ = Mid(strBuffer$, lngStart& + 1, Len(strBuffer$))
    Set ItmX = Form2.ListView1.ListItems.Add(, , strFind)
    ItmX.SubItems(1) = strReplace
  Next i
End Function

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Form_Load()
  Call LoadAutoFormat
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim FndX As ListItem
  If KeyAscii = 32 Or KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii = 58 Or KeyAscii = 59 Or KeyAscii = 33 Or KeyAscii = 63 Then
    If Len(Label1.Caption) < 1 Then Exit Sub
    If Not Form2.ListView1.FindItem(Label1.Caption & Chr$(KeyAscii), , , 1) Is Nothing Then
      Label1.Caption = Label1.Caption & Chr$(KeyAscii)
      Exit Sub
    End If
    Set FndX = Form2.ListView1.FindItem(Label1.Caption)
    If Not FndX Is Nothing Then
      Me.Text1.SelStart = Me.Text1.SelStart - Len(Label1.Caption)
      Me.Text1.SelLength = Len(Label1.Caption)
      Me.Text1.SelText = FndX.ListSubItems(1)
    End If
    Label1.Caption = ""
  ElseIf KeyAscii = 8 And Not Len(Label1.Caption) = 0 Then
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 1)
  ElseIf KeyAscii < 32 Then
    Label2.Caption = ""
    'do nothing
  Else
    If Form2.ListView1.FindItem(Label1.Caption & Chr$(KeyAscii), , , 1) Is Nothing Then
      Label1.Caption = Mid(Label1.Caption, InStr(Label1.Caption, " ") + 1, Len(Label1.Caption)) & Chr$(KeyAscii)
    Else
      Label1.Caption = Label1.Caption & Chr$(KeyAscii)
    End If
  End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim FndX As ListItem
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

