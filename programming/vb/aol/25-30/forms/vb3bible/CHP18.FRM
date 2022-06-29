VERSION 2.00
Begin Form formMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Lists, Combos, and Grids Project"
   ClientHeight    =   5430
   ClientLeft      =   735
   ClientTop       =   1380
   ClientWidth     =   5310
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Terminal"
   FontSize        =   4.5
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   5865
   Icon            =   CHP18.FRX:0000
   Left            =   660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   5310
   Top             =   1020
   Width           =   5460
   Begin ComboBox cmboInflation 
      Height          =   288
      Left            =   2160
      TabIndex        =   1
      Text            =   "cmboInflation"
      Top             =   120
      Width           =   1452
   End
   Begin ListBox listHistory 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   930
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   960
      Width           =   3495
   End
   Begin ComboBox cmboIncome 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "cmboIncome"
      Top             =   120
      Width           =   1932
   End
   Begin ComboBox cmboGridlines 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1932
   End
   Begin CommandButton cmndSave 
      Caption         =   "&Save"
      Height          =   372
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   1452
   End
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   852
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   1452
   End
   Begin CommandButton cmndOpen 
      Caption         =   "&Open"
      Height          =   372
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1452
   End
   Begin ComboBox cmboAlign 
      Height          =   288
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   1452
   End
   Begin Grid gridSheet 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   3012
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   5052
   End
   Begin Label lablCellSelected 
      BackColor       =   &H00C0C0C0&
      Caption         =   "lablCellSelected"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   4320
      TabIndex        =   11
      Top             =   5160
      Width           =   852
   End
   Begin Label lablSelection 
      BackColor       =   &H00C0C0C0&
      Caption         =   "lablSelection"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   2640
      TabIndex        =   10
      Top             =   5160
      Width           =   1572
   End
   Begin Label lablActive 
      BackColor       =   &H00C0C0C0&
      Caption         =   "lablActive"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   2412
   End
End

Function clean (number As Variant) As Single
     'this cleans up a number, removing garbage like commas and dollar signs
     On Error GoTo woops                          'just in case....
     For i = 1 To Len(number)                     'iterate through entire string
	  If InStr("1234567890", Mid$(number, i, 1)) = 0 Then    'if character isn't a number,
	       Mid$(number, i, 1) = " "           'replace with an inoffensive blank
	  End If
     Next i
     clean = Val(number)                          'and make string into a number
Exit Function                                     'done

woops:                                            'something went wrong
     clean = 0                                    'so fake it
Exit Function                                     'and leave

End Function

Sub cmboAlign_Click ()
     For i = gridSheet.SelStartCol To gridSheet.SelEndCol   'iterate through selected columns
	  gridSheet.ColAlignment(i) = cmboAlign.ListIndex   'and set alignment
     Next i
End Sub

Sub cmboGridlines_Click ()
     gridSheet.GridLines = cmboGridlines.ItemData(cmboGridlines.ListIndex) 'set gridlines
End Sub

Sub cmboIncome_DropDown ()
     If InStr(cmboIncome.Text, "  ") Then         'If this is a previously formatted entry
	  value = clean(Left$(cmboIncome.Text, InStr(cmboIncome.Text, "  "))) 'take just the income value on the left
     Else                                         'otherwise
	  value = clean((cmboIncome.Text))        'income value is all there is
     End If
     value = Format$(value, "$###,###,###")       'pretty up the value
     display = value & Space$(12 - Len(value) - Len(cmboInflation.Text)) & cmboInflation.Text 'add inflation to the right
     For i = 0 To cmboIncome.ListCount - 1        'run through the existing list
	  If display = cmboIncome.List(i) Then    'if we've already got this scenario,
	       Exit Sub                           'leave
	  End If
     Next i
     cmboIncome.AddItem display                   'new scenario, so add it to list
     cmboIncome.ItemData(cmboIncome.NewIndex) = cmboInflation.Text 'remember the inflation rate
     cmboIncome.ListIndex = cmboIncome.NewIndex   'and set box to this newest entry
End Sub

Sub cmboIncome_GotFocus ()
     cmboIncome.SelLength = Len(cmboIncome.Text)  'highlight the whole entry
End Sub

Sub cmndExit_Click ()
     End                 'end the program
End Sub

Sub cmndOpen_Click ()
     On Error GoTo badopen                   'set error trap
     Open "chp18.tmp" For Input As #1        'open the data file
     On Error GoTo 0                         'success!
     Do While Not EOF(1)                     'while there's still data,
	Line Input #1, rawdata$              'get a line
	clipdata$ = clipdata$ & rawdata$ & Chr$(13)  'add the carriage return back in
     Loop
     Close                                   'close data file
     NoHilite True                           'turn off highlighting, select whole grid
     gridSheet.Clip = clipdata$              'and put the saved data in the grid
     NoHilite False                          'turn back highlighting, select whole grid
Exit Sub

badopen:                                     'oops! bad file
     MsgBox "Failed to open"                 'inform user
Exit Sub                                     'and leave

End Sub

Sub cmndSave_Click ()
     On Error GoTo badsave                   'set error trap
     Open "chp18.tmp" For Output As #1       'open data file
     On Error GoTo 0                         'success!
     NoHilite True                           'turn off highlighting, select entire grid
     Print #1, gridSheet.Clip                'save the grid's data
     NoHilite False                          'turn on highlighting, select old selection
     Close                                   'close data file
Exit Sub                                     'done

badsave:                                     'oops! bad file
     MsgBox "Failed to save"                 'inform user
Exit Sub                                     'leave

End Sub

Sub Form_Load ()
     gridSheet.Rows = 22                          'dimension grid
     gridSheet.Cols = 5
     gridSheet.RowHeight(0) = gridSheet.RowHeight(1) * 1.3 'make top row a bit bigger
     gridSheet.Row = 0                            'add in the labels
     gridSheet.Col = 1
     gridSheet.Text = "Yearly"
     gridSheet.Col = 2
     gridSheet.Text = "Monthly"
     gridSheet.Col = 3
     gridSheet.Text = "Weekly"
     gridSheet.Col = 4
     gridSheet.Text = "Hourly"
     For i = 1 To gridSheet.Cols - 1              'set column widths and alignments
	  gridSheet.ColWidth(i) = 1000
	  gridSheet.ColAlignment(i) = 1           'right align
	  gridSheet.FixedAlignment(i) = 1         'right align
     Next i
     cmboGridlines.AddItem "No Gridlines"         'set up gridlines combo box
     cmboGridlines.ItemData(0) = 0                'associate correct settings
     cmboGridlines.AddItem "Gridlines"            'with the displayed items
     cmboGridlines.ItemData(1) = -1
     cmboGridlines.ListIndex = 1                  'set the default to "Gridlines"
     cmboAlign.AddItem "Left Align"               'add the alignment options in the correct order
     cmboAlign.AddItem "Right Align"
     cmboAlign.AddItem "Center Align"
     cmboAlign.ListIndex = 1                      'set default alignment to right align
     For i = 2 To 20                              'set up inflation box with
	  cmboInflation.AddItem i                 'inflation rates from 2 to 20
     Next i
     cmboInflation.ListIndex = 4                  'set default inflation rate
     gridSheet_RowColChange                       'force update the active cell label
     cmboIncome.Text = "$50,000"                  'put in default value
     cmboIncome_DropDown                          'force update the displayed value
End Sub

Sub gridSheet_GotFocus ()
     cmboIncome_DropDown                'force the display's update
     bucks = clean(Left$(cmboIncome.Text, InStr(cmboIncome.Text, "  ")))   'take just the income
     inflation = 1 + Val(cmboIncome.ItemData(cmboIncome.ListIndex)) / 100  'and the inflation rate
     NewValues bucks, inflation         'and compute new scenario
End Sub

Sub gridSheet_RowColChange ()
     currentRow$ = "Row " & gridSheet.Row              'update the label every time
     currentCol$ = "Column " & gridSheet.Col           'a new cell becomes active
     lablActive.Caption = currentRow$ & " " & currentCol$
     If gridSheet.CellSelected Then                    'and figure out if the active
	  lablCellSelected.Caption = "[selected]"      'cell is selected
     Else
	  lablCellSelected.Caption = ""
     End If
End Sub

Sub gridSheet_SelChange ()
     numRows = Abs(gridSheet.SelEndRow - gridSheet.SelStartRow) + 1 'number of selected rows
     numCols = Abs(gridSheet.SelEndCol - gridSheet.SelStartCol) + 1 'number of selected columns
     If numCols + numRows = 0 Then                          'if there isn't a selection, then
	  lablSelection.Caption = "(no selection)"          'say so
     Else                                                   'otherwise
	  lablSelection.Caption = numRows & " x " & numCols 'update the display with selection size
     End If
     If gridSheet.CellSelected Then                         'selection has changed, so the
	  lablCellSelected.Caption = "[selected]"           'active cell may or may not be
     Else                                                   'selected, so check and update
	  lablCellSelected.Caption = ""                     'display
     End If
End Sub

Sub listHistory_DblClick ()
     listHistory.Clear                  'clear all history entries
End Sub

Sub listHistory_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then                                     'right mouse button
	  For i = (listHistory.ListCount - 1) To 0 Step -1  'step through from last to first
	       If listHistory.Selected(i) Then              'if user selected item,
		    listHistory.RemoveItem i                'delete it
	       End If
	  Next i
     End If
End Sub

Sub NewValues (bucks, inflation)
     For years = 1 To 20                          'figure out values for the next 20 years
	  gridSheet.Row = years                   'go to the correct row
	  gridSheet.Col = 0                       'go to label column
	  gridSheet.Text = Str$(1992 + years)     'and create label
	  gridSheet.Col = 1                       'go to yearly income column
	  If years = 1 Then                       'first year?
	       gridSheet.Text = Format$(Val(bucks), "$###,###,###") 'just give it base figure
	  Else                                    'otherwise
	       gridSheet.Row = gridSheet.Row - 1  'get last year's data
	       oldVal = clean((gridSheet.Text))   'clean it up
	       gridSheet.Row = gridSheet.Row + 1  'go back to current year
	       gridSheet.Text = Format$(oldVal * inflation, "$###,###,###") 'and add in inflation factor
	  End If
	  newBucks = clean((gridSheet.Text))      'this year's inflated value
	  gridSheet.Col = 2                       'Monthly wage
	  gridSheet.Text = Format$(newBucks / 12, "$###,###,###")
	  gridSheet.Col = 3                       'Weekly wage
	  gridSheet.Text = Format$(newBucks / 52, "$###,###,###")
	  gridSheet.Col = 4                       'Hourly wage, for those of you who only work 40 hours a week
	  gridSheet.Text = Format$(newBucks / 2000, "$###,###,###.00")
     Next years
     gridSheet.Col = 1                            'Yearly wage
     gridSheet.Row = 10                           'for year 10
     year10 = gridSheet.Text
     listHistory.AddItem cmboIncome.Text & Space$(14 - Len(year10)) & year10 'add in the history entry
     gridSheet.TopRow = 1                         'and scroll to the top
     gridSheet.LeftCol = 1                        'left corner of the grid
End Sub

Sub NoHilite (onOff As Integer)
     Static sCol%, sRow%, eCol%, eRow%            'will remember current selection
     If onOff = True Then                         'turn hilight off
	  sCol% = gridSheet.SelStartCol           'remember current selection
	  sRow% = gridSheet.SelStartRow
	  eCol% = gridSheet.SelEndCol
	  eRow% = gridSheet.SelEndRow
	  gridSheet.HighLight = False             'turn off highlight
	  gridSheet.SelStartCol = 1               'and select entire grid
	  gridSheet.SelStartRow = 1
	  gridSheet.SelEndCol = gridSheet.Cols - 1
	  gridSheet.SelEndRow = gridSheet.Rows - 1
     Else                                         'back to normal
	  gridSheet.SelStartCol = sCol%           'reset old selection
	  gridSheet.SelStartRow = sRow%
	  gridSheet.SelEndCol = eCol%
	  gridSheet.SelEndRow = eRow%
	  gridSheet.HighLight = True              'and turn on highlighting
     End If
End Sub

