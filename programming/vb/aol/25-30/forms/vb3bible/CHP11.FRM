VERSION 2.00
Begin Form Form1 
   Caption         =   "Text Project (Click the form)"
   ClientHeight    =   3945
   ClientLeft      =   1560
   ClientTop       =   1950
   ClientWidth     =   5295
   Height          =   4380
   Icon            =   CHP11.FRX:0000
   Left            =   1485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5295
   Top             =   1590
   Width           =   5445
End

Sub Example1 ()
     Const MSG1 = "Top Left"                           'define the text to display
     Const MSG2 = "Center Center"
     Const MSG3 = "Bottom Right"
     
     Cls                                               'clear the form
     AutoReDraw = True                                 'graphics automatically repaint
     FontName = "Arial"                                'select the font
     FontSize = 20                                     'select the font size
     
     Print MSG1                                        'This goes in upper left corner

     CurrentX = (ScaleWidth / 2) - (TextWidth(MSG2) / 2)    'position this exactly
     CurrentY = (ScaleHeight / 2) - (TextHeight(MSG2) / 2)  'in the center
     Print MSG2

     CurrentX = ScaleWidth - TextWidth(MSG3)           'position this flush right and
     CurrentY = ScaleHeight - TextHeight(MSG3)         'flush bottom
     Print MSG3
End Sub

Sub Example2 ()
     Cls                                               'Clear the form
     AutoReDraw = True                                 'graphics automatically repaint
     FontName = "Times New Roman"                      'Select font
     For size = 8 To 20 Step 4                         'step through font sizes
          num$ = ""                                    'will contain a variable length string
          FontSize = size                              'Set the font size
          For j = 1 To 5                               'five examples at each font size
               Print Spc(j); "Space1"; Spc(j * 2); "Space2";   'demonstrate spc()
               num$ = num$ & "*"                       'build up the variable length string
               Print Tab(35); "Left Tab" & num$;       'print at the tab stop
               y = CurrentY                            'remember what line we're on
               Print Tab(65); "";                      'set the print position to column 65
               CurrentX = CurrentX - TextWidth("Right Tab" & num$)  'right justify at column 65
               CurrentY = y                            'reset us back to the correct line
               Print "Right Tab" & num$                'and print the text
          Next j
     Next size
End Sub

Sub Form_Click ()
     Static example As Integer                         'track which example to show
     If example = 0 Then example = 1                   'start at the first example
     If example = 1 Then
          Example1                                     'demonstrate TextHeight and TextWidth
          example = 2                                  'next time will be example2
     Else
          Example2                                     'demonstrate Spc() and Tab()
          example = 1                                  'next time will be example1
     End If
End Sub

