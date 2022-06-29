VERSION 2.00
Begin Form Form1 
   Caption         =   "Printer Project"
   ClientHeight    =   2160
   ClientLeft      =   2520
   ClientTop       =   2625
   ClientWidth     =   3735
   Height          =   2880
   Icon            =   CHP24.FRX:0000
   Left            =   2445
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   3735
   Top             =   1980
   Width           =   3885
   Begin TextBox Text2 
      Height          =   1215
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin Label Label2 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin Menu menuFile 
      Caption         =   "&File"
      Begin Menu menuPrintForm 
         Caption         =   "Print &Form"
      End
      Begin Menu menuPrintGraphics 
         Caption         =   "Print &Graphics"
      End
      Begin Menu menuPrintText 
         Caption         =   "Print &Text File"
      End
      Begin Menu menuSeparator 
         Caption         =   "-"
      End
      Begin Menu menuExit 
         Caption         =   "E&xit"
      End
   End
End

Sub menuExit_Click ()
     End                                     'End the program
End Sub

Sub menuPrintForm_Click ()
    PrintForm                                'print the form
End Sub

Sub menuPrintGraphics_Click ()
    Form1.MousePointer = 11                            'hourglass; we're going to be a while
    PrintBox "This text is in a box", False            'print this text inside an unshaded box
    PrintBox "This text is in a shaded box", True      'print this text inside a shaded box
    PrintLine "This text is crossed out"               'print this text with a line through the middle
    PrintCircle "This text is in a circle"             'print this text inside a circle
    Printer.EndDoc                                     'end the printing, send to print routines
    Form1.MousePointer = 0                             'cursor back to normal
    MsgBox "Done Printing Graphics"                    'inform the user
End Sub

Sub menuPrintText_Click ()
    FileNumber = FreeFile                                   'get next available file number
    FileName = "\VB\CONSTANT.TXT"                           'this line will cause an error if path or file doesn't exist!
    Open FileName For Input As #FileNumber                  'open text file
    Form1.MousePointer = 11                                 'hourglass; we're going to be a while
    Do
        If Printer.CurrentX = Printer.ScaleLeft And Printer.CurrentY = Printer.ScaleTop Then
            PrintHeader CStr(FileName)                   'If at beginning of page, print header
        End If
        Line Input #FileNumber, InputLine                   'get a new line
        SpaceLeft = Printer.ScaleHeight - Printer.ScaleTop - Printer.CurrentY   'figure out how much space is left at bottom
        Printer.Print InputLine                             'print the line
        If SpaceLeft < (3 * Printer.TextHeight(InputLine)) Then PrintFooter     'if not much space left, print footer and start a new page
    Loop Until EOF(FileNumber) Or Printer.Page = 3          'do it until we're done with file or have printed 3 pages
    Close FileNumber                                        'close the file
    Printer.EndDoc                                          'and send the completed document to the printer
    Form1.MousePointer = 0                                  'cursor back to normal
End Sub

Sub PrintBox (Msg As String, Shaded As Integer)
    leftX = Printer.CurrentX                      'left side of box
    topY = Printer.CurrentY                       'top of box
    rightX = Printer.CurrentX + Printer.TextWidth(Msg) + 200     'a bit wider than the text
    bottomY = Printer.CurrentY + Printer.TextHeight(Msg) + 200   'a bit higher than the text
    Printer.CurrentX = Printer.CurrentX + 100     'indent text in from side of box
    Printer.CurrentY = Printer.CurrentY + 100     'indent text down from top of box
    Printer.Print Msg                             'Print the text
    Printer.Line (leftX, topY)-(rightX, bottomY), , B  'draw the box around the text
    If Shaded = True Then                         'supposed to shade?
        For yPos = topY To bottomY Step 50        'go from top to bottom, every 50 dots
            For xPos = leftX To rightX Step 50    'go from left to right, every 50 dots
                Printer.PSet (xPos, yPos)         'and make a dot
            Next
        Next
    End If
    Printer.CurrentX = Printer.ScaleLeft          'go all the way back to the left
    Printer.CurrentY = bottomY + 100              'and go down a bit from where we're at
End Sub

Sub PrintCircle (Msg As String)
    radius = (Printer.TextWidth(Msg) / 2) + 100        'radius is a bit bigger than half the text
    xPos = Printer.CurrentX + radius                   'figure out where the center of the
    yPos = Printer.CurrentY + radius                   'circle will be
    Printer.CurrentX = Printer.CurrentX + 100          'leave a bit of space horizontally,
    Printer.CurrentY = Printer.CurrentY + radius       'and leave enough room for the circle;
    Printer.Print Msg                                  'now print the message
    Printer.Circle (xPos, yPos), radius                'and draw the circle around the message
    Printer.CurrentX = Printer.ScaleLeft               'back to the left side of the page
    Printer.CurrentY = yPos + radius + 100             'and a bit below the bottom of the circle
End Sub

Sub PrintFooter ()
    Printer.Print                                           'Print a blank line
    pageNumber = "Page: " + Format$(Printer.Page, "###")    'figure out and format page number
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(pageNumber)) / 2  'middle of line
    Printer.Print pageNumber;                               'and print the page number
    Printer.NewPage                                         'start a new page
End Sub

Sub PrintHeader (Header As String)
    Printer.CurrentX = Printer.ScaleLeft                                    'all the way to the left
    Printer.Print Date$;                                    'print current system date
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Header)) / 2  ' middle of line
    Printer.Print Header;                                   'print header string
    pageNumber = "Page: " + Format$(Printer.Page, "###")    'figure out page number
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(pageNumber)  'set up for flush right
    Printer.Print pageNumber                                'and print the page number
    Printer.Print " "                                       'force a new line
End Sub

Sub PrintLine (Msg As String)
    yPos = (Printer.TextHeight(Msg) / 2) + Printer.CurrentY 'figure out center of text
    xStart = Printer.CurrentX                               'remember where text starts
    xEnd = xStart + Printer.TextWidth(Msg)                  'this is where text will end
    Printer.Print Msg                                       'print the text
    Printer.Line (xStart, yPos)-(xEnd, yPos)                'and print a line through it
    Printer.CurrentX = Printer.ScaleLeft   'all the way to the left
    Printer.CurrentY = Printer.CurrentY + 100               'and down a bit
End Sub

