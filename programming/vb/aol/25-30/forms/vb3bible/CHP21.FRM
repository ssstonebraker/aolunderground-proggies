VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Focus"
   ClientHeight    =   3390
   ClientLeft      =   1245
   ClientTop       =   2010
   ClientWidth     =   6465
   Height          =   3825
   Icon            =   CHP21.FRX:0000
   Left            =   1170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   6465
   Top             =   1650
   Width           =   6615
   Begin CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   372
      Left            =   4560
      TabIndex        =   15
      Top             =   2760
      Width           =   1572
   End
   Begin CommandButton Command1 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   372
      Left            =   4560
      TabIndex        =   14
      Top             =   1800
      Width           =   1572
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type"
      Height          =   1212
      Left            =   4560
      TabIndex        =   16
      Top             =   240
      Width           =   1572
      Begin OptionButton Option 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Personal"
         Height          =   372
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1092
      End
      Begin OptionButton Option 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Business"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1092
      End
   End
   Begin TextBox Text 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   5
      Left            =   1440
      TabIndex        =   13
      Text            =   " "
      Top             =   2760
      Width           =   2772
   End
   Begin TextBox Text 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   4
      Left            =   1440
      TabIndex        =   11
      Text            =   " "
      Top             =   2280
      Width           =   2772
   End
   Begin TextBox Text 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   3
      Left            =   1440
      TabIndex        =   9
      Text            =   " "
      Top             =   1800
      Width           =   2772
   End
   Begin TextBox Text 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Text            =   " "
      Top             =   1320
      Width           =   2772
   End
   Begin TextBox Text 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Text            =   " "
      Top             =   840
      Width           =   2772
   End
   Begin TextBox Text 
      BackColor       =   &H00C0C0C0&
      Height          =   312
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Text            =   " "
      Top             =   360
      Width           =   2772
   End
   Begin Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C&losing"
      Height          =   192
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   636
   End
   Begin Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City, ST, &Zip"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1140
   End
   Begin Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Address"
      Height          =   192
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   708
   End
   Begin Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Company"
      Height          =   192
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   792
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salu&tation"
      Height          =   192
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   840
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Name"
      Height          =   192
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   504
   End
End

Sub Command1_Click ()
     On Error GoTo OpenNotepad                    'not open yet, so open it
     AppActivate "Notepad - (Untitled)"           'switch focus to notepad
     On Error GoTo 0                              'success!
     For i = 0 To 5                               'step through each text box
          SendKeys Text(i).Text & Chr$(13), True  'and send it to the notepad
     Next i
     AppActivate "Focus"                          'go back to this project
Exit Sub

OpenNotepad:                                      'need to open the notepad
     If tried Then                                'if we've already tried, then
          MsgBox "Failed to open notepad"         'no way
          Exit Sub                                'bye bye
     End If
     tried = True                                 'ok, we're trying now!
     x = Shell("Notepad")                         'attempt to open notepad
     Resume                                       'and start back where we left off (AppActivate)
End Sub

Sub Command2_Click ()
     End                 'end the program
End Sub

Sub Option_Click (Index As Integer)
     If Index = 0 Then '************  Business
          Text(1).Enabled = True     'enable company
          Label3.Enabled = True      'turn on company labek
          Text(4).TabStop = False    'make keyboard entry flow past the Salutation
          Text(5).Text = "Sincerely," 'set default closing
     Else               '***********  Personal
          Text(1).Enabled = False    'Turn off company box
          Label3.Enabled = False     'and company box label
          Text(4).TabStop = True     'set keyboard flow to include Salutation
          Text(5).Text = "Ciao!"     'Set default closing
     End If
     Text(0).SetFocus                'go to the contact name
End Sub

Sub Text_GotFocus (Index As Integer)
     Text(Index).BackColor = QBColor(15)           'highlight box with white background
     Text(Index).SelStart = 0                      'select all the text
     Text(Index).SelLength = Len(Text(Index).Text)
End Sub

Sub Text_LostFocus (Index As Integer)
     Text(Index).BackColor = QBColor(7)                'unhighlight text box
     If Index = 0 Then                                 'if editing contact name,
          n = Text(0).Text                             'get the name
          p = InStr(n, ".")                            'find a period (indicating honorific like Mr., Mrs., Dr., etc)
          If Text(1).Enabled Then                      'if the company name is turned on (ie, business letter requiring formality)
               If p = 0 Then                           'and the user hasn't entered honorific, then
                    Beep                               'bronx cheer
                    Text(0).SetFocus                   'and set the focus back to name to give them another chance
                    Exit Sub
               End If
               For i = Len(n) To 1 Step -1             'step backwards through the name to isolate the last name
                    If Mid$(n, i, 1) = " " Then Exit For  'found the beginning of the last work
               Next i
               Text(4).Text = "Dear " & Left$(n, p) & Mid$(n, i) & ":"  'formal business salutation: "Dear Mr. Hardwick:"
          Else                                         'personal letter!
               If p = 0 Then                           'if there is no honorific
                    i = 1                              'then start at the beginning of the name
               Else                                    'there is an honorific, so
                    i = InStr(n, " ") + 1              'start just after the honorific
               End If
               j = InStr(i, n, " ") - i                'isolate the first name
               If j < 1 Then j = Len(n)                'if the user entered only a first name, choose the entire thing
               Text(4).Text = "Dear " & Mid$(n, i, j) & "," 'and set the default salutation "Dear John,"
          End If
     End If
End Sub

