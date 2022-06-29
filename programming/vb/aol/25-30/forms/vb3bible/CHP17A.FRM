VERSION 2.00
Begin Form formMain 
   Caption         =   "Dialog Project"
   ClientHeight    =   3570
   ClientLeft      =   885
   ClientTop       =   1890
   ClientWidth     =   3495
   Height          =   4290
   Icon            =   CHP17A.FRX:0000
   Left            =   810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   3495
   Top             =   1245
   Width           =   3645
   Begin CommonDialog cdlgBox 
      Left            =   2760
      Top             =   3120
   End
   Begin TextBox textEntry 
      Height          =   3132
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   180
      Width           =   3012
   End
   Begin Menu menuBar 
      Caption         =   "&File"
      Index           =   0
      Begin Menu menuFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin Menu menuFile 
         Caption         =   "&Save"
         Index           =   1
      End
      Begin Menu menuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin Menu menuFile 
         Caption         =   "&Print Setup"
         Index           =   3
      End
      Begin Menu menuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin Menu menuFile 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin Menu menuBar 
      Caption         =   "F&ormat"
      Index           =   1
      Begin Menu menuFormat 
         Caption         =   "&Color"
         Index           =   0
      End
      Begin Menu menuFormat 
         Caption         =   "&Font"
         Index           =   1
      End
   End
   Begin Menu menuBar 
      Caption         =   "&Help"
      Index           =   2
      Begin Menu menuHelp 
         Caption         =   "&VB MsgBox"
         Index           =   0
      End
      Begin Menu menuHelp 
         Caption         =   "&Search"
         Index           =   1
      End
      Begin Menu menuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin Menu menuHelp 
         Caption         =   "&About"
         Index           =   3
      End
   End
End
'MsgBox constants
 Const MB_ABORTRETRYIGNORE = 2
 Const MB_ICONEXCLAMATION = 48
 Const MB_DEFBUTTON3 = 512
 Const IDABORT = 3
 Const IDRETRY = 4
 Const IDIGNORE = 5

'Common Dialog Control
'Action Property
 Const DLG_FILE_OPEN = 1
 Const DLG_FILE_SAVE = 2
 Const DLG_COLOR = 3
 Const DLG_FONT = 4
 Const DLG_PRINT = 5
 Const DLG_HELP = 6

'Individual flags
 Const CC_RGBINIT = &H1&
 Const CF_BOTH = &H3&
 Const CF_EFFECTS = &H100&
 Const PD_ALLPAGES = &H0&
 Const HELP_KEY = &H101
 Const HELP_PARTIALKEY = &H105
 Const HELP_QUIT = &H2

Sub Form_Resize ()
     textEntry.Top = 0                            'make the text box fill the whole form
     textEntry.Left = 0
     textEntry.Height = formMain.ScaleHeight
     textEntry.Width = formMain.ScaleWidth
End Sub

Sub menuFile_Click (Index As Integer)
     Select Case Index
          Case 0 'Open
               Dim lines As Integer, newLine As String, bigLine As String
               Dim mbType As Integer, buttonPressed As Integer, Filename As String
               
               cdlgBox.DialogTitle = "Open Text File"
               cdlgBox.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
               cdlgBox.FilterIndex = 1                      'set default to text files (*.txt)
               cdlgBox.Action = DLG_FILE_OPEN               'Show the File Open dialog
               Filename = cdlgBox.Filename                  'store the filename
               
               On Error GoTo DinnaOpen                      'set error trap
               Open Filename For Input As #1                'open up the file
               On Error GoTo 0                              'successful! turn off trap
               Do While (Not EOF(1)) And lines < 10         'read in up to 10 lines
                    Line Input #1, newLine                  'get the line
                    bigLine = bigLine & newLine & Chr$(13) & Chr$(10) 'append the line
                    lines = lines + 1
               Loop
               textEntry.Text = bigLine                     'assign the line to the text box
               Close                                        'close the text file
          Case 1 'Save
               cdlgBox.DefaultExt = "TXT"
               cdlgBox.DialogTitle = "Save Text File"
               cdlgBox.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
               cdlgBox.FilterIndex = 1
               cdlgBox.Action = DLG_FILE_SAVE               'Show the File Open dialog
               Rem We're not really going to write anything...
          Case 2 '(separator bar)
          Case 3 'Print Setup
               cdlgBox.PrinterDefault = True
               cdlgBox.Flags = PD_ALLPAGES
               cdlgBox.Action = DLG_PRINT                   'Display Print dialog box
               Printer.Print textEntry.Text
          Case 4 '(separator bar)
          Case 5 'Exit
               cdlgBox.HelpFile = "VB.HLP"                  'we're going to call up VB help
               cdlgBox.HelpCommand = HELP_QUIT              'to close down the help file
               cdlgBox.Action = DLG_HELP                    'do it
               End                                          'end program
     End Select
Exit Sub

DinnaOpen:                                                  'oops! bad file name
     mbType = MB_ABORTRETRYIGNORE Or MB_DEFBUTTON3 Or MB_ICONEXCLAMATION
     buttonPressed = MsgBox("Disk Error", mbType, "Cannot Open File") 'what do you want to do?
     Select Case buttonPressed
          Case IDABORT                                      'stop the world, I want to get off
               MsgBox "Text is unchanged"                   'confirm nothing happened
               Exit Sub                                 'and leave
          Case IDRETRY                                      'let's do it again!
               Filename = InputBox$("Enter new filename:")  'plug in a new filename by hand
               Resume                                       'and try to open it again
          Case IDIGNORE                                     'no big deal, just ignore it
               Exit Sub                                 'and leave
     End Select
End Sub

Sub menuFormat_Click (Index As Integer)
     Select Case Index
          Case 0 ' color
               cdlgBox.Color = textEntry.BackColor                'default color is what's already in text box
               cdlgBox.Flags = CC_RGBINIT                         'set for rgb values
               cdlgBox.Action = DLG_COLOR                         'Display color box
               textEntry.BackColor = cdlgBox.Color                'set the new color
          Case 1 ' font
               cdlgBox.Flags = CF_BOTH Or CF_EFFECTS               'both printer and screen + special effects
               cdlgBox.FontBold = textEntry.FontBold               'set the dialog box's default
               cdlgBox.FontItalic = textEntry.FontItalic           'values to what's already in
               cdlgBox.FontUnderline = textEntry.FontUnderline     'the text box
               cdlgBox.FontStrikethru = textEntry.FontStrikethru
               cdlgBox.Color = textEntry.ForeColor
               cdlgBox.FontName = textEntry.FontName
               cdlgBox.FontSize = textEntry.FontSize
               
               cdlgBox.Action = DLG_FONT                            'Display the Font dialog box
               
               textEntry.FontBold = cdlgBox.FontBold                'and set the text box's properties
               textEntry.FontItalic = cdlgBox.FontItalic
               textEntry.FontStrikethru = cdlgBox.FontStrikethru
               textEntry.FontUnderline = cdlgBox.FontUnderline
               textEntry.ForeColor = cdlgBox.Color
               textEntry.FontName = cdlgBox.FontName
               textEntry.FontSize = cdlgBox.FontSize
     End Select
End Sub

Sub menuHelp_Click (Index As Integer)
     Select Case Index
          Case 0 ' VB Help
               cdlgBox.HelpFile = "VB.HLP"          'we're going to call up VB help
               cdlgBox.HelpCommand = HELP_KEY       'search for a keyword
               cdlgBox.HelpKey = "MsgBox"           'ask about MsgBox!
               cdlgBox.Action = DLG_HELP            'do it
          Case 1 'Search
               cdlgBox.HelpFile = "VB.HLP"          'we're going to call up VB help
               cdlgBox.HelpCommand = HELP_PARTIALKEY 'bring up search box
               cdlgBox.Action = DLG_HELP            'do it
          Case 2 '(separator bar)
          Case 3 'About
               formAbout.Show 1                     'show about box modally
     End Select

End Sub

