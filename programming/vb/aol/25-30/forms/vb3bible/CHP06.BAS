Global CurrentIndex As Integer    'Used to track which letter option we're on
Global TypeIndex As Integer       'Used to track if we're actively changing letter type

Global Const RTN_ADDRESS = 0      'These are the seven different letter options
Global Const ADDRESSEE = 1        '        (they are NOT constant.txt values!)
Global Const GREETING = 2
Global Const BODY = 3
Global Const CLOSING = 4
Global Const ENCLOSURE = 5
Global Const CARBON = 6

Global Const MODAL = 1             'Modal Dialog box
Global Const CHECKED = 1           'Checked Check Box
Global Const GRAYED = 2            'Grayed Check Box

Global Const BLACK = &H0           'Colors used in ForeColor property
Global Const BLUE = &HFF0000

Sub GenerateLetter ()
     BLANK = Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)        'Two CR LF combos
     L = ""                                                   'Erase the letter
     L = L & formEntry!textEntry(RTN_ADDRESS).Text & BLANK    'Add return address
     L = L & formEntry!textEntry(ADDRESSEE).Text & BLANK      'Add Addressee
     If formEntry!textEntry(GREETING).Text = "" Then          'If there is no greeting,
          L = L & "To Whom it may concern:" & BLANK           'Add a default greeting
     Else                                                     '...otherwise
          L = L & formEntry!textEntry(GREETING).Text & BLANK  'Add the user's greeting
     End If
     L = L & formEntry!textEntry(BODY).Text & BLANK           'Add body text
     L = L & formEntry!textEntry(CLOSING).Text & BLANK        'Add the closing statement
     If formEntry!textEntry(ENCLOSURE).Text <> "" Then        'If there are enclosures
          L = L & "Encl: " & formEntry!textEntry(ENCLOSURE).Text & BLANK 'Add them
     End If
     If formEntry!textEntry(CARBON).Text <> "" Then           'If there are Carbon Copies
          L = L & "CC:   " & formEntry!textEntry(CC).Text     'Add them
     End If
     formLetter!textPrintOut.Text = L                         'Transfer letter for display
End Sub

Sub HiLite (controlName As Control)
     controlName.ForeColor = BLUE                 'Make ForeColor text blue
End Sub

Sub UnHiLite (controlName As Control)
     controlName.ForeColor = BLACK                'make ForeColor text back to default Black
End Sub

