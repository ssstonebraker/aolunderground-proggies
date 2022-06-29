VERSION 2.00
Begin Form formMain 
   Caption         =   "Text Box Project"
   ClientHeight    =   3915
   ClientLeft      =   915
   ClientTop       =   1875
   ClientWidth     =   6840
   Height          =   4635
   Icon            =   CHP13A.FRX:0000
   Left            =   840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6840
   Top             =   1230
   Width           =   6990
   Begin TextBox textDocument 
      Height          =   1212
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   1452
   End
   Begin Label lablStatus 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   0 Characters"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   4932
   End
   Begin Menu menuBar 
      Caption         =   "&File"
      Index           =   0
      Begin Menu menuFile 
         Caption         =   "&New"
         Index           =   0
      End
      Begin Menu menuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin Menu menuFile 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin Menu menuBar 
      Caption         =   "&Edit"
      Index           =   1
      Begin Menu menuEdit 
         Caption         =   "&Undo"
         Index           =   0
      End
      Begin Menu menuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin Menu menuEdit 
         Caption         =   "Cu&t"
         Index           =   2
      End
      Begin Menu menuEdit 
         Caption         =   "&Copy"
         Index           =   3
      End
      Begin Menu menuEdit 
         Caption         =   "&Paste"
         Index           =   4
      End
      Begin Menu menuEdit 
         Caption         =   "&Delete"
         Index           =   5
      End
   End
End
Dim TxtBeforeChange As String
Dim PosBeforeChange As Integer
Dim LenBeforeChange As Integer

Sub Form_Resize ()
     textDocument.Height = formMain.ScaleHeight - 250
     textDocument.Width = formMain.ScaleWidth
     lablStatus.Top = textDocument.Height
     lablStatus.Height = 250
     lablStatus.Width = formMain.ScaleWidth
End Sub

Sub menuBar_Click (Index As Integer)
     Select Case Index
     Case 0  ' File
     Case 1  ' Edit
          menuEdit(0).Enabled = False     'undo
          menuEdit(2).Enabled = False     'cut
          menuEdit(3).Enabled = False     'copy
          menuEdit(4).Enabled = False     'paste
          menuEdit(5).Enabled = False     'delete
          If Len(TxtBeforeChange) > 0 Then menuEdit(0).Enabled = True 'undo
          If textDocument.SelLength > 0 Then
               menuEdit(2).Enabled = True 'cut
               menuEdit(3).Enabled = True 'copy
               menuEdit(5).Enabled = True 'delete
          End If
          If Clipboard.GetFormat(1) Then menuEdit(4).Enabled = True    'paste
     End Select
End Sub

Sub menuEdit_Click (Index As Integer)
     Select Case Index
          Case 0 'undo
               textDocument.Text = TxtBeforeChange
               textDocument.SelStart = PosBeforeChange
               textDocument.SelLength = LenBeforeChange
          Case 2 'cut
               Clipboard.SetText textDocument.SelText
               textDocument.SelText = ""
          Case 3 'copy
               Clipboard.SetText textDocument.SelText
          Case 4 'paste
               textDocument.SelText = Clipboard.GetText(1)
          Case 5 'delete
               textDocument.SelText = ""
     End Select
End Sub

Sub menuFile_Click (Index As Integer)
     Select Case Index
          Case 0 'new
               formPassword!textPassword.Text = ""
               formPassword.Show 1                           'modal
               If formPassword!textPassword.Text <> "" Then  'all passwords work!
                    textDocument.Text = ""
               Else                                          'user didn’t even try, so
                    Beep                                     'Bronx cheer
                    menuFile_Click 0                         'bad password, try again
               End If
          Case 2 'exit
               End
     End Select
End Sub

Sub textDocument_Change ()
     temp = Format$(Len(textDocument.Text), "  ##,###,##0")
     temp = temp + " Character"
     If Len(textDocument.Text) <> 1 Then temp = temp + "s"
     lablStatus.Caption = temp
End Sub

Sub textDocument_KeyDown (KeyCode As Integer, Shift As Integer)
     If Shift < 2 Then
          TxtBeforeChange = textDocument.Text
          PosBeforeChange = textDocument.SelStart
          LenBeforeChange = textDocument.SelLength
     End If
End Sub

