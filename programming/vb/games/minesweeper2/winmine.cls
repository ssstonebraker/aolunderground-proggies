VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "clsWinMine"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
'***********************************************************************************'
'                                                                                   '
'   WINMINE: OVERVIEW                                                               '
'   -----------------                                                               '
'                                                                                   '
'   This is a Game Sample Application similar to MineSweeper that ships with        '
'   Windows 3.x \ Windows NT. This project comprises of the following files:        '
'                                                                                   '
'   winmine.cls:    This is a class module that implements the main functionality   '
'                   of different aspects of the game.                               '
'                                                                                   '
'   winmine.frm:    The main display form, that implements the user interface and   '
'                   instantiates an object of the above class.                      '
'                                                                                   '
'   coords.cls:     This is another class module that just implements a (X,Y)       '
'                   co-ordinate pair for wrong mine location markings.              '
'                                                                                   '
'   custdlg.frm:    This is the form that is shown modally when the custom game     '
'                   level is chosen from the Game\Custom menu.                      '
'                                                                                   '
'   instruct.frm:   This is the form that displays the rules and playing            '
'                   instructions when F1 is pressed.                                '
'                                                                                   '
'   about.frm       This is the form that displays info about the author etc.       '
'                                                                                   '
'***********************************************************************************'
Option Explicit

' left mouse button constant used by VB
Private Const LEFT_BUTTON As Byte = 1

' flag indicating empty square
Private Const NONE As Byte = 0
' flag indicating a square with a mine
Private Const MINE As Byte = 243
' flag indicating that square has already been opened
Private Const BEEN As Byte = 244
' flag indicating that square has been marked to be a mine
Private Const FLAGGED As Byte = 2
' flag indicating ambiguous square
Private Const QUESTION As Byte = 1

' Maximum\minimum # of mines, rows and columns
Private Const MIN_MINES As Byte = 10
Private Const MAX_MINES As Byte = 99
Private Const MIN_ROWS As Integer = 8
Private Const MAX_ROWS As Integer = 24
Private Const MIN_COLS As Integer = 8
Private Const MAX_COLS As Integer = 36

' Width of a square in pixels
Private Const mintButtonWidth As Byte = 16
' Height of a square in pixels
Private Const mintButtonHeight As Byte = 16

' Total number of Mines in current game level
Private mbytNumMines As Byte
' Number of squares correctly marked to indicate containing a mine
Private mbytCorrectHits As Byte
' Total number of marked squares (including wrong ones)
Private mbytTotalHits As Byte

' Total number of rows and columns in current game level
Private mintRows As Integer
Private mintCols As Integer

' row and column currently being processed
Private mintRow As Integer
Private mintCol As Integer

' flag indicating that its time for a new game
Public mblnNewGame As Boolean
' flag indicating that a mouse click is currently being processed
Private mblnHitTestBegun As Boolean
' variable to hold the main display form
Private mfrmDisplay As Form

' dynamic 2D array to keep track of which squares contain mines,
' which ones indicate mines surround them, which ones have
'already been opened, etc.
Private mbytMineStatus() As Byte

' dynamic 2D array to keep track of the current marking status of a square
' -- whether it is unmarked, ambiguous, flagged correctly, or incorrectly
Private mbytMarked() As Byte

' dynamic 2D array to keep track of the X and Y co-ords
' of the mbytNumMines mine locations in the minefield
Private mbytMineLocations() As Byte

' A collection of clsCoords objects to hold the
' X and Y co-ords of the squares marked wrongly
' to contain mines
Private mcolWrongLocations As New Collection
'***********************************************************************************'
'                                                                                   '
' Purpose:  Determines which square was clicked and with which mouse button, and    '
'           takes action accordingly. Called from the MouseDown event of the main   '
'           display form.                                                           '
'                                                                                   '
' Inputs:   intButton:  The mouse button clicked (left or right\middle)             '
'           inX:        X co-ordinate of mouse cursor position                      '
'           inY:        Y co-ordinate of mouse cursor position                      '
'                                                                                   '
' Returns:  None                                                                    '
'                                                                                   '
'***********************************************************************************'
Public Sub BeginHitTest(intButton As Integer, intX As Single, intY As Single)
    
    ' If the current game is over, start a new game
    ' when the minefield is clicked
    If mblnNewGame Then
        NewGame
    End If
    
    ' Indicate that a mouse click is currently in progress
    mblnHitTestBegun = True
    
    ' Calculate row and col grid co-ords from mouse co-ords
    intX = Int(intX / mintButtonWidth)
    intY = Int(intY / mintButtonHeight)

    ' abort, if  co-ords lie outside minefield
    If intX >= mintCols Or intY >= mintRows Or intX < 0 Or intY < 0 Then
        Exit Sub
    End If

    ' calculate exact square co-ords from grid co-ords
    mintCol = intX * mintButtonWidth
    mintRow = intY * mintButtonHeight

    ' abort, if square already opened
    If mbytMineStatus(intY, intX) >= BEEN Then Exit Sub

    Dim blnLeftDown As Boolean
    blnLeftDown = (intButton And LEFT_BUTTON) > 0

    ' If left mouse button clicked ...
    If blnLeftDown Then
        
        ' if square already marked, can't open, so abort
        If mbytMarked(intY, intX) >= FLAGGED Then Exit Sub
        
        ' temporarily display image control with appropriate bitmap
        If mbytMarked(intY, intX) = QUESTION Then
            mfrmDisplay.imgPressed.Visible = False
            mfrmDisplay.imgQsPressed.Visible = False
            mfrmDisplay.imgQsPressed.Left = mintCol
            mfrmDisplay.imgQsPressed.TOP = mintRow
            mfrmDisplay.imgQsPressed.Visible = True
        Else
            mfrmDisplay.imgQsPressed.Visible = False
            mfrmDisplay.imgPressed.Visible = False
            mfrmDisplay.imgPressed.Left = mintCol
            mfrmDisplay.imgPressed.TOP = mintRow
            mfrmDisplay.imgPressed.Visible = True
        End If
        
    Else    ' if right mouse button clicked ...
        
        Dim Msg As String
        Dim CRLF As String

        CRLF = Chr$(13) & Chr$(10)
        
        Select Case mbytMarked(intY, intX)
    
            Case NONE:      ' if you run out of squares to mark ...
                            If mbytTotalHits = mbytNumMines Then
                                Msg = "Can't Mark Any More Mines!" & CRLF
                                Msg = Msg & "One or more Mines have been wrongly marked." & CRLF
                                Msg = Msg & "UnMark one or more mines with the right mouse button."
                                    
                                MsgBox Msg, vbCritical, "WinMine: Error!"
                                Exit Sub
                            End If
                                    
                            ' if not marked, display a flag in the square to indicate marking
                            mfrmDisplay.PaintPicture mfrmDisplay.imgFlag, mintCol, mintRow
                            ' increment the total # of squares marked
                            mbytTotalHits = mbytTotalHits + 1

                            ' Update display of mines left
                            mfrmDisplay.lblMinesLeft = "Mines Left : " & mbytNumMines - mbytTotalHits
                            
                            ' if correctly marked ...
                            If mbytMineStatus(intY, intX) = MINE Then
                                mbytCorrectHits = mbytCorrectHits + 1
                                mbytMarked(intY, intX) = FLAGGED
                            Else    ' if wrongly marked ...
                                Dim objCoords As New clsCoords
                                    
                                ' store co-ords of wrong location in a new object
                                objCoords.mintX = intX
                                objCoords.mintY = intY
                                    
                                ' and add it to the collection
                                mcolWrongLocations.Add objCoords
                                
                                ' store the index in the collection, of this wrongly marked square
                                ' in the corresponding element of the mbytMarked array.
                                mbytMarked(intY, intX) = mbytTotalHits - mbytCorrectHits + 2
                            End If
                                
                            ' if all mines were correctly marked ...
                            If mbytCorrectHits = mbytNumMines Then
                                Msg = "Congratulations!" & CRLF
                                Msg = Msg & "You have won." & CRLF
                                
                                MsgBox Msg, vbInformation, "WinMine"
                                
                                ' prepare for new game
                                mblnNewGame = True
                            End If
            
            Case QUESTION:  ' if ambiguously marked, unmark it
                            mbytMarked(intY, intX) = NONE
                            ' and display the original square
                            mfrmDisplay.PaintPicture mfrmDisplay.imgButton, mintCol, mintRow

            Case Else:      ' if previously marked with a flag, mark it as ambiguous now
                            ' and display square with ?
                            mfrmDisplay.PaintPicture mfrmDisplay.imgQuestion, mintCol, mintRow
                            
                            ' Decrement total number of marked squares
                            mbytTotalHits = mbytTotalHits - 1
                                
                            ' Update display of mines left
                            mfrmDisplay.lblMinesLeft = "Mines Left : " & mbytNumMines - mbytTotalHits
                            
                            ' if previously marked square contained a mine...
                            If mbytMineStatus(intY, intX) = MINE Then
                                ' decrement the number of correctly marked squares as well
                                mbytCorrectHits = mbytCorrectHits - 1
                            Else    ' if it is a wrongly marked square ...
                                ' remove this wrongly marked co-ords from corresponding position in the collection
                                mcolWrongLocations.Remove mbytMarked(intY, intX) - 2
                                    
                                Dim intXwm As Integer   ' X co-ord of wrong location
                                Dim intYwm As Integer   ' Y co-ord of wrong location
                                Dim i As Integer        ' Loop counter
                                    
                                ' Update the index of the other wrong co-ords in the collection,
                                ' (that appear after the currently deleted item), in the mbytMarked array.
                                For i = mbytMarked(intY, intX) - 2 To mcolWrongLocations.Count
                                    intXwm = mcolWrongLocations(i).mintX
                                    intYwm = mcolWrongLocations(i).mintY
                                    mbytMarked(intYwm, intXwm) = mbytMarked(intYwm, intXwm) - 1
                                Next
                                    
                            End If

                            mbytMarked(intY, intX) = QUESTION
                
        End Select
    
    End If

End Sub
'***********************************************************************************'
'                                                                                   '
' Purpose:  Determines over which square the mouse curser is when the left mouse    '
'           button is released and takes action accordingly. Called from the        '
'           MouseUp event of the main display form                                  '
'                                                                                   '
' Inputs:   intButton:  The mouse button clicked (left or right\middle)             '
'           inX:        X co-ordinate of mouse cursor position                      '
'           inY:        Y co-ordinate of mouse cursor position                      '
'                                                                                   '
' Returns:  None                                                                    '
'                                                                                   '
'***********************************************************************************'
Public Sub EndHitTest(intButton As Integer, intX As Single, intY As Single)
    
    ' if currently processing a mouse click ...
    If mblnHitTestBegun Then
        ' reset the flag
        mblnHitTestBegun = False
    Else
        ' if not, abort. This makes sure that the code below is executed only
        ' when a mouse button is pressed.
        Exit Sub
    End If
    
    Dim blnLeftDown As Boolean
    blnLeftDown = (intButton And LEFT_BUTTON) > 0

    ' if left mouse button pressed ...
    If blnLeftDown Then
        
        ' Calculate row and col grid co-ords from the clicked mouse co-ords
        intX = Int(intX / mintButtonWidth)
        intY = Int(intY / mintButtonHeight)

        ' abort, if  co-ords lie outside minefield
        If intX >= mintCols Or intY >= mintRows Or intX < 0 Or intY < 0 Then
            Exit Sub
        End If

        ' abort, if current square, over which mouse is released has been
        ' marked with a flag
        If mbytMarked(intY, intX) >= FLAGGED Then Exit Sub

        ' Calculate the grid co-ords from last valid mouse cursor co-ords
        intX = mintCol \ mintButtonWidth
        intY = mintRow \ mintButtonHeight
    
        If mbytMarked(intY, intX) = QUESTION Then
            mfrmDisplay.imgQsPressed.Visible = False
        Else
            mfrmDisplay.imgPressed.Visible = False
        End If
        
        Select Case mbytMineStatus(intY, intX)

            Case Is >= BEEN:    ' abort, if current square has already been opened
                                Exit Sub
            
            Case NONE:  ' if current square is empty, iteratively open all
                        ' surrounding squares until non-empty squares are reached
                        OpenBlanks intX, intY
                        
            Case MINE:  ' if current square contains a mine, you blew it!
                        Dim intXm As Integer        ' X co-ord of mine location
                        Dim intYm As Integer        ' Y co-ord of mine location
                        Dim vntCoord As Variant     ' variant used in For Each loop
                        Dim i As Integer            ' Loop counter
                
                        ' reveal all the squares that contain mines
                        For i = 0 To mbytNumMines - 1
                            
                            intYm = mbytMineLocations(i, 0)
                            intXm = mbytMineLocations(i, 1)

                            If mbytMarked(intYm, intXm) < FLAGGED Then
                                mfrmDisplay.PaintPicture mfrmDisplay.imgMine, intXm * mintButtonWidth, intYm * mintButtonHeight
                            End If
                        
                        Next

                        ' display the current square as a blown mine
                        mfrmDisplay.PaintPicture mfrmDisplay.imgBlown, mintCol, mintRow
                        
                        ' reveal all the squares that were wrongly marked as mines
                        For Each vntCoord In mcolWrongLocations

                            intYm = vntCoord.mintY
                            intXm = vntCoord.mintX

                            mfrmDisplay.PaintPicture mfrmDisplay.imgWrongMine, intXm * mintButtonWidth, intYm * mintButtonHeight
                        
                        Next
                        
                        ' prepare for new game
                        mblnNewGame = True
                        
                        Dim CRLF As String
                        CRLF = Chr$(13) & Chr$(10)
                        
                        MsgBox "You Lose!", vbExclamation, "WinMine"

            Case Else:  ' if current square surrounds one or more squares that contains a mine
                        ' reveal the number of such mines that surrounds it
                        mfrmDisplay.PaintPicture mfrmDisplay.imgPressed, mintCol, mintRow
                        mfrmDisplay.CurrentX = mintCol
                        mfrmDisplay.CurrentY = mintRow
                        mfrmDisplay.ForeColor = QBColor(mbytMineStatus(intY, intX))
                        mfrmDisplay.Print mbytMineStatus(intY, intX)
                        
                        ' and mark it as being opened
                        mbytMineStatus(intY, intX) = mbytMineStatus(intY, intX) + BEEN
    
        End Select
    
    End If

End Sub
'***********************************************************************************'
'                                                                                   '
' Purpose:  Does other important stuff like calculate the size of the display form  '
'           when the form object used for display is passed to the class by         '
'           assigning it to this property procedure. Called in the form load event  '
'           of the main display form.                                               '
'                                                                                   '
' Inputs:   frmDisplay: The form object that is used as the main display form       '
' Returns:  None                                                                    '
'                                                                                   '
'***********************************************************************************'
Public Property Set frmDisplay(frmDisplay As Form)
    
    Set mfrmDisplay = frmDisplay
    mfrmDisplay.FontBold = True
    
    ' Resize the form to fit the chosen game level's minefield dimensions
    ResizeDisplay
    
End Property
'***********************************************************************************'
'                                                                                   '
' Purpose:  Displays the values for the number of rows, columns and mines from the  '
'           current game level, in the textboxes of the custom dialog box           '
'                                                                                   '
' Inputs:   frmDialog:  The modal form object that is used as the custom dialog box '
' Returns:  None                                                                    '
'                                                                                   '
'***********************************************************************************'
Public Sub GetMineFieldDimensions(frmDialog As Form)

    frmDialog.txtRows = mintRows
    frmDialog.txtColumns = mintCols
    frmDialog.txtMines = mbytNumMines
    
    frmDialog.txtRows.SelLength = Len(frmDialog.txtRows)
    frmDialog.txtColumns.SelLength = Len(frmDialog.txtColumns)
    frmDialog.txtMines.SelLength = Len(frmDialog.txtMines)

End Sub
'***********************************************************************************'
'                                                                                   '
' Purpose:  Allocates memory for dynamic arrays according to current minefield      '
'           dimensions and sets up the mine locations in the minefield              '
'                                                                                   '
' Inputs:   None                                                                    '
' Returns:  None                                                                    '
'                                                                                   '
'***********************************************************************************'
Private Sub InitializeMineField()
    
    ' allocate space for the 2D dynamic arrays to fit the current
    ' minefield size
    ReDim mbytMineStatus(mintRows - 1, mintCols - 1)
    ReDim mbytMarked(mintRows - 1, mintCols - 1)
    ReDim mbytMineLocations(mbytNumMines - 1, 1)

    ' Generate random mine locations in the minefield and store them in
    ' the mbytMineLocations array. Also fill the mbytMineStatus array with
    ' info as to which squares contain mines and which are the ones that
    ' indicate surrounding mines
    Randomize

    Dim i As Integer    ' Loop counter
    Dim r As Integer    ' Loop counter
    Dim c As Integer    ' Loop counter

    For i = 0 To mbytNumMines - 1

        Dim intX As Integer
        Dim intY As Integer

        intX = Int(Rnd * mintCols)
        intY = Int(Rnd * mintRows)

        While mbytMineStatus(intY, intX) = MINE
            intX = Int(Rnd * mintCols)
            intY = Int(Rnd * mintRows)
        Wend

        mbytMineStatus(intY, intX) = MINE
        mbytMineLocations(i, 0) = intY
        mbytMineLocations(i, 1) = intX

        For r = -1 To 1
            For c = -1 To 1

                Dim blnDx As Boolean
                Dim blnDy As Boolean

                blnDy = intY + r >= 0 And intY + r < mintRows
                blnDx = intX + c >= 0 And intX + c < mintCols

                If blnDy And blnDx Then
                    If mbytMineStatus(intY + r, intX + c) <> MINE Then
                        mbytMineStatus(intY + r, intX + c) = mbytMineStatus(intY + r, intX + c) + 1
                    End If
                End If

            Next
        Next

    Next

End Sub
'***********************************************************************************'
'                                                                                   '
' Purpose:  Prepares for a new game                                                 '
' Inputs:   None                                                                    '
' Returns:  None                                                                    '
'                                                                                   '
'***********************************************************************************'
Public Sub NewGame()
    
    ' clear the current display in the main form
    mfrmDisplay.Cls
 
    ' reset game variables and flags
    mbytCorrectHits = 0
    mbytTotalHits = 0

    mintRow = -1
    mintCol = -1

    mblnNewGame = False
    mblnHitTestBegun = False
    
    Dim i As Integer            ' Loop counter

    ' empty the collection of wrong co-ords
    For i = 1 To mcolWrongLocations.Count
        mcolWrongLocations.Remove 1
    Next
    
    ' re-calculate new mine locations
    InitializeMineField
    
    ' Reset the display of number of mines left
    mfrmDisplay.lblMinesLeft = "Mines Left : " & mbytNumMines
    
End Sub
'***************************************************************************************'
'                                                                                       '
' Purpose:  If the square that was clicked was empty, then this function iteratively    '
'           opens all squares surrounding it, until non-empty squares are encountered   '
'           A general fill algorithm is adopted here, wherein, the current square       '
'           location keeps moving left, until it comes across a non-empty square. From  '
'           here onwards, it tries to trace out a border of non-empty squares by moving '
'           clockwise in the directions where it can go. At the same time it stores the '
'           pairs of starting and ending X co-ords on each scanline that crosses the    '
'           region enclosed by the traced out border.                                   '
'                                                                                       '
' Inputs:   inX:    X grid co-ordinate of the square where the mouse was clicked        '
'           inY:    Y grid co-ordinate of the square where the mouse was clicked        '
'                                                                                       '
' Returns:  None                                                                        '
'                                                                                       '
'***********************************************************************************'***'
Private Sub OpenBlanks(ByVal intX As Single, ByVal intY As Single)

    ' flags to keep track of the direction in which current square moves
    Dim blnGoUp As Boolean
    Dim blnGoRight As Boolean
    Dim blnGoDown As Boolean
    Dim blnGoLeft As Boolean
    
    ' stores the X and Y co-ords of the square from which the trace for
    ' the border starts
    Dim intXStart As Integer
    Dim intYStart As Integer
    
    ' index used for collection item
    Dim intPos As Integer
    ' variant used in For Each loop
    Dim element As Variant
    
    ' Loop counters
    Dim y As Integer
    Dim x As Integer
    Dim i As Integer
    
    ' A dynamic array of collections of intgers. Each element in the
    ' array is a collection of starting and ending X co-ord pairs that
    ' define a part of the scan line going through the region enclosed
    ' by the traced out border
    Dim colX() As New Collection
    
    ' Size of this array is the same as the number of rows in the minefield
    ReDim colX(mintRows - 1)
    
    ' keep going left, until you reach a non-empty mine
    While mbytMineStatus(intY, intX) = NONE
        
        intX = intX - 1

        If intX < 0 Then
            intX = 0
            intXStart = intX
            intYStart = intY
            GoTo LFT
        End If
    
    Wend

    ' first direction to go is up
    blnGoUp = True
    
    ' store this first non-empty mine location as the starting point.
    intXStart = intX
    intYStart = intY

    ' trace out a border iteratively, until you return back to the
    ' starting point
    Do
        If mbytMineStatus(intY, intX) = NONE Then
            
            If blnGoUp Then
                intX = intX - 1
                intY = intY + 1
                colX(intY).Remove (colX(intY).Count)
                blnGoUp = False
                blnGoLeft = True
            ElseIf blnGoRight Then
                intX = intX - 1
                intY = intY - 1
                blnGoRight = False
                blnGoUp = True
            ElseIf blnGoDown Then
                intX = intX + 1
                intY = intY - 1
                colX(intY).Remove (colX(intY).Count)
                blnGoDown = False
                blnGoRight = True
            ElseIf blnGoLeft Then
                intX = intX + 1
                intY = intY + 1
                blnGoLeft = False
                blnGoDown = True
            End If

            If (intXStart = intX And intYStart = intY) Then Exit Do
        
        Else

            If blnGoUp Then

                colX(intY).Add intX

                If mbytMineStatus(intY, intX + 1) = NONE Then
                    
                    If intY = 0 Then
                        blnGoUp = False
UP:                     intX = intX + 1
                        If (intXStart = intX And intYStart = intY) Then Exit Do
                        While mbytMineStatus(intY, intX) = NONE
                            If intX = mintCols - 1 Then GoTo RIGHT
                            intX = intX + 1
                            If (intXStart = intX And intYStart = intY) Then Exit Do
                        Wend
                        blnGoDown = True
                    Else
                        intY = intY - 1
                        If (intXStart = intX And intYStart = intY) Then Exit Do
                    End If
                
                Else
                    
                    blnGoUp = False
                    blnGoRight = True
                    intX = intX + 1
                    If (intXStart = intX And intYStart = intY) Then
                        If colX(intY).Count Mod 2 <> 0 Then
                            intPos = 1
                            For Each element In colX(intY)
                                If element = intXStart Then
                                    colX(intY).Remove (intPos)
                                    Exit Do
                                End If
                                intPos = intPos + 1
                            Next
                        End If
                        Exit Do
                    End If
                
                End If
            
            ElseIf blnGoRight Then
                
                If mbytMineStatus(intY + 1, intX) = NONE Then
                    
                    If intX = mintCols - 1 Then
                        blnGoRight = False
RIGHT:                  colX(intY).Add intX
                        intY = intY + 1
                        If (intXStart = intX And intYStart = intY) Then Exit Do
                        While mbytMineStatus(intY, intX) = NONE
                            colX(intY).Add intX
                            If intY = mintRows - 1 Then GoTo DOWN
                            intY = intY + 1
                            If (intXStart = intX And intYStart = intY) Then Exit Do
                        Wend
                        colX(intY).Add intX
                        blnGoLeft = True
                    Else
                        intX = intX + 1
                        If (intXStart = intX And intYStart = intY) Then
                            If colX(intY).Count Mod 2 <> 0 Then
                                intPos = 1
                                For Each element In colX(intY)
                                    If element = intXStart Then
                                        colX(intY).Remove (intPos)
                                        Exit Do
                                    End If
                                    intPos = intPos + 1
                                Next
                            End If
                            Exit Do
                        End If
                    End If
                
                Else
                    
                    blnGoRight = False
                    blnGoDown = True
                    
                    colX(intY).Add intX
                    intY = intY + 1
                    If (intXStart = intX And intYStart = intY) Then Exit Do
                
                End If
            
            ElseIf blnGoDown Then
               
                colX(intY).Add intX
               
                If mbytMineStatus(intY, intX - 1) = NONE Then
                    
                    If intY = mintRows - 1 Then
                        blnGoDown = False
DOWN:                   intX = intX - 1
                        If (intXStart = intX And intYStart = intY) Then Exit Do
                        While mbytMineStatus(intY, intX) = NONE
                            If intX = 0 Then GoTo LFT
                            intX = intX - 1
                            If (intXStart = intX And intYStart = intY) Then Exit Do
                        Wend
                        blnGoUp = True
                    Else
                        intY = intY + 1
                        If (intXStart = intX And intYStart = intY) Then Exit Do
                    End If
                
                Else
                    
                    blnGoDown = False
                    blnGoLeft = True

                    intX = intX - 1
                    If (intXStart = intX And intYStart = intY) Then Exit Do
                
                End If
            
            ElseIf blnGoLeft Then
                
                If mbytMineStatus(intY - 1, intX) = NONE Then
                    
                    If intX = 0 Then
                        blnGoLeft = False
LFT:                    colX(intY).Add intX
                        If intY = 0 Then GoTo UP
                        intY = intY - 1
                        If (intXStart = intX And intYStart = intY) Then Exit Do
                        While mbytMineStatus(intY, intX) = NONE
                            colX(intY).Add intX
                            If intY = 0 Then GoTo UP
                            intY = intY - 1
                            If (intXStart = intX And intYStart = intY) Then Exit Do
                        Wend
                        colX(intY).Add intX
                        blnGoRight = True
                    Else
                        intX = intX - 1
                        If (intXStart = intX And intYStart = intY) Then Exit Do
                    End If
                
                Else
                    
                    blnGoLeft = False
                    blnGoUp = True

                    colX(intY).Add intX
                    intY = intY - 1
                    If (intXStart = intX And intYStart = intY) Then Exit Do
                
                End If
            
            End If
        
        End If

    Loop

    ' iterate through the collection for each scanline and paint
    ' the opened squares in one go
    For y = 0 To mintRows - 1
        
        If colX(y).Count > 0 Then
        
            ' Sort the X co-ord pairs in ascending order, by using
            ' a standard Listbox control
            For x = 1 To colX(y).Count
                
                Dim intXValue As Integer
                intXValue = colX(y)(x)
                
                If intXValue < 10 Then
                    intXValue = intXValue + 48
                ElseIf intXValue >= 10 Then
                    intXValue = intXValue + 55
                End If
                
                mfrmDisplay.lstSortedX.AddItem Chr$(intXValue)
            
            Next
            
            ' Display opened squares between and including each X co-ord
            ' pair for the collection in the current scanline
            For x = 0 To mfrmDisplay.lstSortedX.ListCount - 1 Step 2
            
                Dim intR1 As Integer
                Dim intC1 As Integer
                Dim intColStart As Integer
                Dim intColEnd As Integer
                Dim intDx As Integer
                Dim intWidth As Integer
                
                intR1 = y * mintButtonHeight
                
                intColStart = Asc(mfrmDisplay.lstSortedX.List(x))
                If intColStart <= 57 Then
                    intColStart = intColStart - 48
                ElseIf intColStart >= 65 Then
                    intColStart = intColStart - 55
                End If
                
                intColEnd = Asc(mfrmDisplay.lstSortedX.List(x + 1))
                If intColEnd <= 57 Then
                    intColEnd = intColEnd - 48
                ElseIf intColEnd >= 65 Then
                    intColEnd = intColEnd - 55
                End If
                
                intC1 = intColStart * mintButtonWidth
                intDx = intColEnd - intColStart + 1
                intWidth = intDx * mintButtonWidth

                mfrmDisplay.PaintPicture mfrmDisplay.imgOpenBlocks, intC1, intR1, , , 0, 0, intWidth, mintButtonHeight
        
                For i = 0 To intDx - 1
                    
                    If mbytMarked(y, intColStart + i) > NONE Then
                        
                        If mbytMarked(y, intColStart + i) = QUESTION Then
                            mfrmDisplay.PaintPicture mfrmDisplay.imgQuestion, intC1 + i * mintButtonWidth, intR1
                        Else
                            mfrmDisplay.PaintPicture mfrmDisplay.imgFlag, intC1 + i * mintButtonWidth, intR1
                        End If
                    
                    ElseIf mbytMineStatus(y, intColStart + i) > NONE Then
                        
                        mfrmDisplay.CurrentX = intC1 + i * mintButtonWidth
                        mfrmDisplay.CurrentY = intR1
                        
                        If mbytMineStatus(y, intColStart + i) >= BEEN Then
                            mfrmDisplay.ForeColor = QBColor(mbytMineStatus(y, intColStart + i) - BEEN)
                            mfrmDisplay.Print mbytMineStatus(y, intColStart + i) - BEEN
                        ElseIf mbytMineStatus(y, intColStart + i) = MINE Then
                             mfrmDisplay.PaintPicture mfrmDisplay.imgButton, intC1 + i * mintButtonWidth, intR1
                        Else
                            mfrmDisplay.ForeColor = QBColor(mbytMineStatus(y, intColStart + i))
                            mfrmDisplay.Print mbytMineStatus(y, intColStart + i)
                            mbytMineStatus(y, intColStart + i) = mbytMineStatus(y, intColStart + i) + BEEN
                        End If
                    
                    End If
                    
                Next
                
            Next
        
            ' clear the listbox for the next scanline collection
            mfrmDisplay.lstSortedX.Clear
            
        End If
        
    Next

End Sub
'***************************************************************************************'
'                                                                                       '
' Purpose:  Resizes the main display form to fit the currently chosen game level's      '
'           minefield dimensions.                                                       '
'                                                                                       '
' Inputs:   None                                                                        '
' Returns:  None                                                                        '
'                                                                                       '
'***********************************************************************************'***'
Private Sub ResizeDisplay()
    
    ' set the form dimensions
    mfrmDisplay.ScaleMode = 1
    mfrmDisplay.Width = mfrmDisplay.Width - mfrmDisplay.ScaleWidth + mintCols * mintButtonWidth * Screen.TwipsPerPixelX
    mfrmDisplay.Height = mfrmDisplay.Height - mfrmDisplay.ScaleHeight + mintRows * mintButtonHeight * Screen.TwipsPerPixelY + mfrmDisplay.lblMinesLeft.Height
    
    ' set the label (that displays the number of mines left) dimensions
    mfrmDisplay.lblMinesLeft.Left = 0
    mfrmDisplay.lblMinesLeft.TOP = mfrmDisplay.ScaleHeight - mfrmDisplay.lblMinesLeft.Height
    mfrmDisplay.lblMinesLeft.Width = mfrmDisplay.ScaleWidth
    mfrmDisplay.lblMinesLeft = "Mines Left : " & mbytNumMines
    
    mfrmDisplay.ScaleMode = 3

End Sub
'***********************************************************************************'
'                                                                                   '
' Purpose:  Determines over which square the mouse curser is, at present, while the '
'           left mouse button is pressed, and takes action accordingly. Called from '
'           the MouseMove event of the main display form                            '
'                                                                                   '
' Inputs:   intButton:  The mouse button clicked (left or right\middle)             '
'           inX:        X co-ordinate of mouse cursor position                      '
'           inY:        Y co-ordinate of mouse cursor position                      '
'                                                                                   '
' Returns:  None                                                                    '
'                                                                                   '
'***********************************************************************************'
Public Sub TrackHitTest(intButton As Integer, intX As Single, intY As Single)

    Dim blnLeftDown As Boolean
    blnLeftDown = (intButton And LEFT_BUTTON) > 0
    
    ' If left mouse button pressed ...
    If blnLeftDown Then
        
        ' abort, if not currently processing a mouse click
        If Not mblnHitTestBegun Then Exit Sub

        ' calculate the grid co-ords from the mouse co-ords
        intX = Int(intX / mintButtonWidth)
        intY = Int(intY / mintButtonHeight)

        ' abort, if the square over which the mouse cursor is currently
        ' over is outside the minefield
        If intX >= mintCols Or intY >= mintRows Or intX < 0 Or intY < 0 Then
            mfrmDisplay.imgQsPressed.Visible = False
            mfrmDisplay.imgPressed.Visible = False
            Exit Sub
        End If

        ' abort, if current square has been marked with a flag
        If mbytMarked(intY, intX) >= FLAGGED Then
            mfrmDisplay.imgQsPressed.Visible = False
            mfrmDisplay.imgPressed.Visible = False
            Exit Sub
        End If

        Dim intRowOld As Integer
        Dim intColOld As Integer
        
        ' store previous grid location of cursor
        intRowOld = mintRow
        intColOld = mintCol

        ' calculate current grid co-ords of mouse cursor
        mintCol = intX * mintButtonWidth
        mintRow = intY * mintButtonHeight

        ' Display current square as pressed, only if previous grid co-ords
        ' are not same as current grid co-ords
        If intRowOld = mintRow And intColOld = mintCol Then
            If mfrmDisplay.imgPressed.Visible Or mfrmDisplay.imgQsPressed.Visible Then
                Exit Sub
            End If
        End If
        
        ' abort, if current square already opened
        If mbytMineStatus(intY, intX) >= BEEN Then
            mfrmDisplay.imgPressed.Visible = False
            mfrmDisplay.imgQsPressed.Visible = False
            Exit Sub
        End If
        
        ' if current square is unmarked, or ambiguously marked display the
        ' corresponding square as pressed while the mouse cursor is over it
        ' and the left mouse button is pressed
        If mbytMarked(intY, intX) = QUESTION Then
            mfrmDisplay.imgPressed.Visible = False
            mfrmDisplay.imgQsPressed.Visible = False
            mfrmDisplay.imgQsPressed.Left = mintCol
            mfrmDisplay.imgQsPressed.TOP = mintRow
            mfrmDisplay.imgQsPressed.Visible = True
        Else
            mfrmDisplay.imgQsPressed.Visible = False
            mfrmDisplay.imgPressed.Visible = False
            mfrmDisplay.imgPressed.Left = mintCol
            mfrmDisplay.imgPressed.TOP = mintRow
            mfrmDisplay.imgPressed.Visible = True
        End If
    
    End If
    
End Sub
'***************************************************************************************'
'                                                                                       '
' Purpose:  Called when an object of type clsWinMine is instantiated. Initializes       '
'           game variables and flags and sets up the minefield                          '
'                                                                                       '
' Inputs:   None                                                                        '
' Returns:  None                                                                        '
'                                                                                       '
'***********************************************************************************'***'
Private Sub Class_Initialize()

    mbytNumMines = 10
    mbytCorrectHits = 0
    mbytTotalHits = 0
    
    mintRows = 8
    mintCols = 8
    mintRow = -1
    mintCol = -1
    
    mblnNewGame = False
    mblnHitTestBegun = False
    Set mfrmDisplay = Nothing
    
    ' Calculate random mine locations
    InitializeMineField
    
End Sub
'***************************************************************************************'
'                                                                                       '
' Purpose:  Checks to see if the specified number of rows, columns and mines for the    '
'           currently chosen game level is within limits and stores them in the         '
'           appropriate class properties.                                               '
'                                                                                       '
' Inputs:   intRows:        Number of rows in the minefield                             '
'           intCols:        Number of columns in the minefield                          '
'           bytMines:       Number of mines in the minefield                            '
'           blnLevelCustom: True if game level is custom; False otherwise               '
'                                                                                       '
' Returns:  None                                                                        '
'                                                                                       '
'***********************************************************************************'***'
Public Sub SetMineFieldDimension(intRows As Integer, intCols As Integer, bytMines As Byte, blnLevelCustom As Boolean)
    
    mintRows = intRows
    If intRows < MIN_ROWS Then mintRows = MIN_ROWS
    If intRows > MAX_ROWS Then mintRows = MAX_ROWS
        
    mintCols = intCols
    If intCols < MIN_COLS Then mintCols = MIN_COLS
    If intCols > MAX_COLS Then mintCols = MAX_COLS

    mbytNumMines = bytMines
    If blnLevelCustom Then
        Dim intMines As Integer
        intMines = (mintRows * mintCols) \ 5
        If bytMines < intMines Then
            mbytNumMines = intMines
            bytMines = intMines
        ElseIf bytMines > (intMines * 4) \ 3 Then
            mbytNumMines = (intMines * 4) \ 3
            bytMines = mbytNumMines
        End If
    End If
    
    If bytMines < MIN_MINES Then mbytNumMines = MIN_MINES
    If bytMines > MAX_MINES Then mbytNumMines = MAX_MINES
    
    ' clear the current display to start new game
    mfrmDisplay.Cls
    
    ' Adjust the display form size according to the new minefield dimensions
    ResizeDisplay
    
End Sub
'***************************************************************************************'
'                                                                                       '
' Purpose:  Called when the instance of the clsWinMine object is set to nothing when    '
'           the program terminates. Frees memory used for dynamically allocated arrays  '
'           and empties the collection of Wrong Mine locations.                         '
'                                                                                       '
' Inputs:   None                                                                        '
' Returns:  None                                                                        '
'                                                                                       '
'***********************************************************************************'***'
Private Sub Class_Terminate()
    
    Erase mbytMineStatus
    Erase mbytMarked
    Erase mbytMineLocations
    
    Dim i As Integer            ' Loop counter
    
    For i = 1 To mcolWrongLocations.Count
        mcolWrongLocations.Remove 1
    Next

End Sub
