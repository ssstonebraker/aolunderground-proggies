VERSION 2.00
Begin Form formChild 
   Caption         =   "Child Form"
   ClientHeight    =   2700
   ClientLeft      =   615
   ClientTop       =   2055
   ClientWidth     =   4290
   Height          =   3420
   Left            =   540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2700
   ScaleWidth      =   4290
   Top             =   1410
   Width           =   4440
   Begin TextBox Text1 
      Height          =   852
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   1152
   End
   Begin Menu menuFile 
      Caption         =   "&File"
      Begin Menu menuFileSave 
         Caption         =   "&Save"
      End
      Begin Menu menuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin Menu menuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin Menu menuWindowArrange 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin Menu menuWindowArrange 
         Caption         =   "Tile &Horizontal"
         Index           =   1
      End
      Begin Menu menuWindowArrange 
         Caption         =   "Tile &Vertical"
         Index           =   2
      End
      Begin Menu menuWindowArrange 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
      Begin Menu menuWindowArrange 
         Caption         =   "-"
         Index           =   4
      End
      Begin Menu menuWindowArrange 
         Caption         =   "&New Child"
         Index           =   5
      End
      Begin Menu menuWindowArrange 
         Caption         =   "Mo&dal"
         Index           =   6
      End
      Begin Menu menuWindowArrange 
         Caption         =   "Mode&less"
         Index           =   7
      End
   End
End

Sub Form_Activate ()
     MDIForm1!pictStatus.Cls
     MDIForm1!pictStatus.Print "Child form #" & Me.Tag
End Sub

Sub Form_Paint ()
     Text1 = "This is instance # " & Me.Tag
     Me.Caption = "Child Form " & Me.Tag
End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
    Dim msg, docTitle, saveChange
    If DocState(Me.Tag) Then          ' DocState flags changes to this child's text1
        docTitle = Me.Caption         ' Identifies which document is being saved
        msg = "Document has changed.  Save Changes?"
        saveChange = MsgBox(msg, 51, docTitle)' Exclamation with Yes/No/Cancel
        Select Case saveChange
        Case 2                        ' User chooses cancel; stop unloading process
            Cancel = True
        Case 6                        ' Save the file and continue unloading
            menuFileSave_Click        ' Save the document!
            Cancel = False
        Case 7                        ' User chooses no save; continue unloading
            Cancel = False
        End Select
    End If
End Sub

Sub Form_Resize ()
     'Expand the text box to completely fill
     'Child form's internal area
     Text1.Height = ScaleHeight
     Text1.Width = ScaleWidth
End Sub

Sub menuFileExit_Click ()
     Quit
End Sub

Sub menuFileSave_Click ()
     'write whatever code to save the file here...
     DocState(Me.Tag) = False    'and show no changes to text
End Sub

Sub menuWindowArrange_Click (Index As Integer)
     Select Case Index
          Case 0: MDIForm1.Arrange CASCADE
          Case 1: MDIForm1.Arrange TILE_HORIZONTAL
          Case 2: MDIForm1.Arrange TILE_VERTICAL
          Case 3: MDIForm1.Arrange ARRANGE_ICONS
          Case 5: DisplayNewChild
          Case 6: DisplayModal
          Case 7: DisplayModeless
     End Select
End Sub

Sub Text1_Change ()
     DocState(Me.Tag) = True   'Show that text1 has changed
End Sub

Sub Text1_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = RIGHT_BUTTON Then
          PopupMenu menuWindow, 2  'pop up menu, OK to choose item w/right mouse button
     End If
End Sub

