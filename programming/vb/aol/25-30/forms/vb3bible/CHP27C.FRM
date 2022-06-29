VERSION 2.00
Begin Form formFind 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Find"
   ClientHeight    =   3156
   ClientLeft      =   672
   ClientTop       =   3012
   ClientWidth     =   4092
   Height          =   3600
   Left            =   612
   LinkTopic       =   "Form3"
   ScaleHeight     =   3156
   ScaleWidth      =   4092
   Top             =   2628
   Width           =   4212
   Begin CommandButton cmndClose 
      Caption         =   "&Close"
      Height          =   372
      Left            =   2880
      TabIndex        =   12
      Top             =   2640
      Width           =   1092
   End
   Begin CommandButton cmndClearHistory 
      Caption         =   "Clear &History"
      Height          =   372
      Left            =   1440
      TabIndex        =   11
      Top             =   2640
      Width           =   1332
   End
   Begin CommandButton cmndGoto 
      Caption         =   "&Goto"
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   1092
   End
   Begin ListBox listFindHistory 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   984
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   3732
   End
   Begin CommandButton cmndFind 
      Caption         =   "&Last"
      Height          =   372
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   960
      Width           =   852
   End
   Begin CommandButton cmndFind 
      Caption         =   "&First"
      Height          =   372
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   852
   End
   Begin CommandButton cmndFind 
      Caption         =   "&Next"
      Height          =   372
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   852
   End
   Begin CommandButton cmndFind 
      Caption         =   "&Previous"
      Height          =   372
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   852
   End
   Begin CommandButton cmndAndOr 
      Caption         =   "OR"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   228
      Left            =   372
      TabIndex        =   2
      Top             =   600
      Width           =   684
   End
   Begin TextBox textFind 
      Height          =   288
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   1932
   End
   Begin TextBox textFind 
      Height          =   288
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1932
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "this in path:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   1092
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Find this in path name:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1692
   End
End

Dim FindHistory() As String               'module level array to hold bookmarks

Sub cmndAndOr_Click ()
     If cmndAndOr.Caption = "OR" Then     'switch states
          cmndAndOr.Caption = "AND"       '(this caption is used to
     Else                                 ' build a SQL WHERE clause)
          cmndAndOr.Caption = "OR"
     End If
End Sub

Sub cmndClearHistory_Click ()
     listFindHistory.Clear                'clear out old history
     ReDim FindHistory(0)                 'clear out old bookmarks
     cmndGoto.Enabled = False             'disable goto bookmark
End Sub

Sub cmndClose_Click ()
     formFind.Hide                        'hide this form
     formMain.Show                        'make sure main form is current
End Sub

Sub cmndFind_Click (Index As Integer)
     'build SQL WHERE clause: FileName Like '*name*' AND/OR PathName Like '*path*'
     'this finds any embedded text, not just exact matches
     criteria$ = "FileName Like '*" & textFind(0).Text & "*' " & cmndAndOr.Caption
     criteria$ = criteria$ & " PathName Like '*" & textFind(1).Text & "*'"
     
     cmndFind(0).Enabled = True           'start by enabling all buttons
     cmndFind(1).Enabled = True
     Select Case Index
          Case 0 'Previous
               formMain!dataMain.Recordset.FindPrevious criteria$
               If formMain!dataMain.Recordset.NoMatch Then cmndFind(0).Enabled = False
          Case 1 'Next
               formMain!dataMain.Recordset.FindNext criteria$
               If formMain!dataMain.Recordset.NoMatch Then cmndFind(1).Enabled = False
          Case 2 'First
               formMain!dataMain.Recordset.FindFirst criteria$
               cmndFind(0).Enabled = False
          Case 3 'Last
               formMain!dataMain.Recordset.FindLast criteria$
               cmndFind(1).Enabled = False
     End Select
     If Not formMain!dataMain.Recordset.NoMatch Then  'if we found a record, then
          item$ = formMain!textPathName.Text & "\" & formMain!textFileName.Text
          listFindHistory.AddItem item$               'add it to the history
          If formMain!dataMain.Recordset.Bookmarkable Then
               ReDim Preserve FindHistory(listFindHistory.ListCount - 1)
               FindHistory(listFindHistory.ListCount - 1) = formMain!dataMain.Recordset.Bookmark
               cmndGoto.Enabled = True                 'enable Goto bookmark
          End If
     End If
End Sub

Sub cmndGoto_Click ()
     book = FindHistory(listFindHistory.ListIndex)     'extract bookmark
     formMain!dataMain.Recordset.Bookmark = book       'and go to it
     formMain!dataMain.UpdateControls                  'make sure display is current
End Sub

Sub Form_Load ()
     textFind(0).Text = ""                             'blank out beginning text
     textFind(1).Text = ""
End Sub

Sub listFindHistory_DblClick ()
     cmndGoto_Click                                    'goto bookmark
End Sub

Sub textFind_Change (Index As Integer)
     cmndFind(0).Enabled = False             'changed criteria; no more previous
     cmndFind(1).Enabled = False             'changed criteria; no more next
End Sub

Sub textFind_GotFocus (Index As Integer)
     textFind(Index).SelStart = 0            'start at the beginning of text
     textFind(Index).SelLength = 100         'select all text
End Sub

Sub textFind_LostFocus (Index As Integer)
     textFind(Index).SelLength = 0           'deselect all text
End Sub

