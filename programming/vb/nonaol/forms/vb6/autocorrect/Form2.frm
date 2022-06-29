VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Format"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   3000
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   255
      Left            =   330
      TabIndex        =   6
      Top             =   3030
      Width           =   645
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2115
      Left            =   180
      TabIndex        =   3
      Top             =   780
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   3731
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Find"
         Object.Width           =   5146
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Replace With"
         Object.Width           =   5145
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   315
      Left            =   5130
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1860
      TabIndex        =   1
      Top             =   330
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   330
      TabIndex        =   8
      Top             =   3420
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "With:"
      Height          =   225
      Left            =   1860
      TabIndex        =   5
      Top             =   90
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Replace:"
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  'If the wrong word is already in the list
  'don't add it.
  'Else add it.
  Dim ItmX As ListItem
  If Not ListView1.FindItem(Text1.Text) Is Nothing Then
    Exit Sub
  End If
  Set ItmX = ListView1.ListItems.Add(, , Text1.Text)
  ItmX.SubItems(1) = Text2.Text
  Text1.Text = ""  'Clear textboxs
  Text2.Text = ""
  Text1.SetFocus  'And set text1 with the focus for
                  'easy word entering
End Sub


Private Sub Command2_Click()
  'Save list
  Dim Str As String
  Dim i As Integer
  For i = 1 To ListView1.ListItems.Count 'loop through making the list
    Str$ = Str$ & ListView1.ListItems(i)
    Str$ = Str$ & Chr$(1) & ListView1.ListItems(i).SubItems(1) & Chr$(2)
  Next i
  Str$ = ListView1.ListItems.Count & Chr$(3) & Str$
  Open App.Path & "\AutoFormat.msf" For Output As #1
  Print #1, Str$
  Close #1
  
End Sub

Private Sub Command3_Click()
  'Remove a word from the list
  If Not ListView1.FindItem(Text1.Text) Is Nothing Then
    Call ListView1.ListItems.Remove(ListView1.SelectedItem.Index)
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
  End If

End Sub

Private Sub Command4_Click()
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'don't close the form or the list will be removed
  Cancel = True
  Me.Hide
End Sub

Private Sub ListView1_Click()
  'If a word is clicked in the listview
  'Add them to the textbox
  If ListView1.ListItems.Count = 0 Then Exit Sub
  Text1.Text = ListView1.SelectedItem
  Text2.Text = ListView1.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub Text1_Change()
  Dim FndX As ListItem
  'As the word is typed, search for it in the list
  Set FndX = ListView1.FindItem(Text1.Text, , , lvwPartial)
  If Not FndX Is Nothing Then
    FndX.EnsureVisible
    FndX.Selected = True
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  'If enter is pressed and there is no
  'wrong word or correct word then don't do anything
  'else add it to the list
  If KeyAscii = 13 Then
    If Not Text1.Text = "" And Not Text2.Text = "" Then
      Command1_Click
    End If
  End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Not Text1.Text = "" And Not Text2.Text = "" Then
      Command1_Click
    End If
  End If
End Sub
