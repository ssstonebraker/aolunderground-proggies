VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'   Battleship-clone.
'   This class-file contains the game itself. Implementation is your
'   problem. A simple interface is provided
'   THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
'   ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
'   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS for A
'   PARTICULAR PURPOSE.
'   Copyright 1999, Tom Werner Halvorsroed, Trysil Norway

Const HUMAN_MISS = 0
Const HUMAN_HIT = 1
Const HUMAN_WINNER = 2
Const HUMAN_TAKEN = 3
Const HUMAN_WRONG = 4

Const COMPUTER_MISS = 0
Const COMPUTER_HIT = 1
Const COMPUTER_WINNER = 2
Const COMPUTER_SUNK = 3

Const human = 0
Const computer = 1

Const GRID_LOW = 1
Const GRID_HIGH = 10

Const HIT_ATTRACTION = 50

'valid numbers for m_nLevel
Const MATE = 0
Const ENSIGN = 1
Const LIEUTENANT = 2
Const COMMANDER = 3
Const CAPTAIN = 4
Const REAR_ADMIRAL = 5
Const ADMIRAL = 6

'player, you and the computer
'define values in player()()()
'1=vacant/unknown
'M = 2 = miss
'O=3=vicinity of the ships only used in auto-placing of the ships
'S>=12=ship that is 12=two-square-ship and so forth
'H=hit-ship>=22 that is 22=two-square-ship-hit
Private player() As Integer '(COMPUTER+1)(GRID_HIGH+2)(GRID_HIGH+2)

'hit-points, you and the computer
Private hit() As Integer '(COMPUTER+1)
    
'cell-attraction:
Private attract() As Integer '(GRID_HIGH+2)(GRID_HIGH+2)
    
'current level for computer
Private m_nLevel As Integer

'record over which ship is still floating
Private ship_record() As Integer '(4)

'init the board
'        Parameters
'        random: Use the Randomize statement when you release the game
'                If you exclude Randomize the game-progress will be
'                the same for every game, with the same level. This
'                is done for debug-purposes
'        level : Level is defined at the top of this file

'        return nothing
Public Sub init_board(ByVal level As Integer)
    Dim count_ship As Integer
    Dim pl As Integer
    Dim col As Integer
    Dim row As Integer
    Dim ship_size As Integer
    Dim i As Integer
    Dim j As Integer
    Dim deltacol  As Integer
    Dim deltarow  As Integer
    ReDim player(computer + 1, GRID_HIGH + 2, GRID_HIGH + 2)
    ReDim hit(computer + 1)
    ReDim attract(GRID_HIGH + 2, GRID_HIGH + 2)
    ReDim ship_record(4)
    
    'setting all cells to vacant
    For pl = human To computer
        For row = GRID_LOW To GRID_HIGH
            For col = GRID_LOW To GRID_HIGH
                player(pl, row, col) = 1
            Next
        Next
    Next

    'place the ships
    For pl = human To computer
        For count_ship = 5 To 1 Step -1
            If count_ship = 1 Then
                ship_size = 2
            Else
                ship_size = count_ship
            End If
again:
            deltacol = 1
            deltarow = 0
    
            'direction values
            If Int(Rnd() * 2) = 0 Then
                deltacol = 0
                deltarow = 1
            End If
    
            'find start-row and col
            col = Int(Rnd() * (10 - ship_size * deltacol)) + 1
            row = Int(Rnd() * (10 - ship_size * deltarow)) + 1
    
            'can we place the ship?
            For i = 1 To ship_size 'check whether those squares are taken
                If player(pl, row + i * deltarow, col + i * deltacol) <> 1 Then GoTo again
            Next
    
            'yes we can, let's do it
            For i = 1 To ship_size
                player(pl, row + i * deltarow, col + i * deltacol) = 10 + ship_size
            Next
    
            'to make sure that any ship is not placed to close to another ship
            'we'll surround all ship's with a value. This value will be removed
            'prior to any play of course
            For row = GRID_LOW To GRID_HIGH
                For col = GRID_LOW To GRID_HIGH
                    'if a ship is present
                    If player(pl, row, col) >= 12 Then
                        For i = (row - 1) To row + 1
                            For j = (col - 1) To col + 1
                                'if it is vacant
                                If player(pl, i, j) = 1 Then
                                    player(pl, i, j) = 3
                                End If
                            Next
                        Next
                    End If
                Next
            Next
        Next
    Next

    'remove all occurences of threes, no need for them anymore
    For pl = human To computer
        For row = GRID_LOW To GRID_HIGH
            For col = GRID_LOW To GRID_HIGH
                If player(pl, row, col) = 3 Then
                    player(pl, row, col) = 1
                End If
            Next
        Next
    Next

    'initialize the attraction-array
    For row = GRID_LOW To GRID_HIGH
        For col = GRID_LOW To GRID_HIGH
            attract(row, col) = 1
        Next
    Next

    'initialize the hit-array
    hit(human) = 0
    hit(computer) = 0

    m_nLevel = level

    'remaining "live" points for the ships
    ship_record(0) = 4
    ship_record(1) = 3
    ship_record(2) = 4
    ship_record(3) = 5
End Sub

'if the first occurrence is higher than zero then you found the smallest/
'biggest ship around
Private Function big_ship() As Integer
    Dim i As Integer
    
    For i = 3 To 0 Step -1
        If ship_record(i) > 0 Then
            big_ship = i + 2
            Exit Function
        End If
    Next

    big_ship = 0
End Function

Private Function small_ship() As Integer
    Dim i As Integer
    
    For i = 0 To 3
        If ship_record(i) > 0 Then
            small_ship = i + 2
            Exit Function
        End If
    Next

    small_ship = 0
End Function

Public Function get_hit(ByVal pl As Integer) As Integer
    get_hit = hit(pl)
End Function

Public Function get_player(ByVal pl As Integer, ByVal row As Integer, ByVal col As Integer) As Integer
    get_player = player(pl, row, col)
End Function

'execute the players choice
'    parameters
'    row and col on the board
'    return values are defined at the top of this file

Public Function human_turn(ByVal row As Integer, ByVal col As Integer) As Integer
    'silly-check
    If row < 1 Or row > 10 Or col < 1 Or col > 10 Then
        human_turn = HUMAN_WRONG
        Exit Function
    End If

    'is it available for a shot?
    If player(human, row, col) = 2 Or player(human, row, col) >= 20 Then
        human_turn = HUMAN_TAKEN
        Exit Function
    End If

    'did you hit? if so, set the cell to a hit-value
    'and update the number of hits for the player
    If player(human, row, col) >= 12 And player(human, row, col) < 16 Then
        player(human, row, col) = player(human, row, col) + 10

        'update hit-count
        hit(human) = hit(human) + 1

        If hit(human) = 16 Then
            human_turn = HUMAN_WINNER
            Exit Function
        End If

        human_turn = HUMAN_HIT
        Exit Function
    End If

    player(human, row, col) = 2
    human_turn = HUMAN_MISS
End Function

'
'the computer "AI"

'the computer computes an "attraction value" for each cell on the grid
'ranging from 0, either a previous miss or a non-attractive cell, to
'11, which signifies a possible hit-situation.

'Level MATE          : Plain stupid. All cell's are treated equal
'Level ENSIGN        : The computer discovers to shoot next to a hit
'Level LIEUTENANT    : The computer does not shoot a sixth shot on a five-ship etc
'Level COMMANDER     : The computer avoids cells that can't hold a ship
'Level CAPTAIN       : The computer shoots in the largest region somewhere
'Level REAR_ADMIRAL  : The computer shoots in the middle of the largest region
'Level ADMIRAL       : The computer shoots in the middle of the largest region, more intelligent
Public Function computer_turn(ByRef m_nRow As Integer, ByRef m_nCol As Integer) As Integer
    Dim r, c, k, count  As Integer           'counters
    Dim row, col  As Integer            'current shot
    Dim large_val  As Integer         'the highest number in the attraction-array
    Dim attract_count   As Integer    'and the count of these
    Dim rand_num  As Integer            'which one shall we shoot
    Dim sunk As Boolean      'did the computer sink another ship
    Dim the_big_one  As Integer         'the biggest ship floating
    Dim high As Integer
    Dim mycount As Integer
    Dim direction As Integer
    
    sunk = False
    'commence next shot
    'counts all cells with highest value in attract(r,c)
    For r = GRID_LOW To GRID_HIGH
        For c = GRID_LOW To GRID_HIGH
            If attract(r, c) = large_val Then attract_count = attract_count + 1
            If attract(r, c) > large_val Then
                attract_count = 1
                large_val = attract(r, c)
            End If
        Next
    Next

    '2. creating a random number among the highest cells
    If (attract_count > 1) Then
        rand_num = Int(Rnd() * attract_count) + 1
    Else
        rand_num = 1
    End If

    '3. find the cell this corresponds to
    attract_count = 0

    For row = GRID_LOW To GRID_HIGH
        For col = GRID_LOW To GRID_HIGH
            If (attract(row, col) = large_val) Then attract_count = attract_count + 1
            'found it
            'let's continue the show
            If (attract_count >= rand_num) Then GoTo found_it
        Next
    Next

'bad programming? who cares!!!
found_it:

    m_nRow = row
    m_nCol = col

    'if the cell we shot at was a miss, but close to a ship we'll have to bail
    'out immediately
    'HIT_ATTRACTION will only be set by ENSIGN or higher
    'this is a kind of forced play. a lot of array-trimming happens
    'when we miss or the computer finds out it has sunk another ship
    'and we can't/shouldn't/won't do that on this occasion
    If attract(row, col) = HIT_ATTRACTION And player(computer, row, col) = 1 Then
        'make sure this cell is not shot upon anymore
        attract(row, col) = 0
        player(computer, row, col) = 2
        computer_turn = COMPUTER_MISS
        Exit Function
    End If

    'make sure this cell is not shot upon anymore
    attract(row, col) = 0

    'Is this a hit?
    If player(computer, row, col) >= 12 And player(computer, row, col) < 16 Then
        'the computer has made a hit and therefore do what's appropriate.
        'by setting the cells over, under, left and right
        'to HIT_ATTRACTION (if they haven't been set to 0 on a previous pass)
        'and the diagonals to 0 will make the computer much smarter because
        'we will with two hits have found the direction of the ship
        
        If m_nLevel > MATE Then
            attract(row - 1, col - 1) = 0
            attract(row - 1, col + 1) = 0
            attract(row + 1, col - 1) = 0
            attract(row + 1, col + 1) = 0

            'if the cell above the current has not been shot upon
            'then we'll make sure that it's considered for a shot
            'the next time
            If attract(row - 1, col) > 0 Then attract(row - 1, col) = HIT_ATTRACTION
            If attract(row, col - 1) > 0 Then attract(row, col - 1) = HIT_ATTRACTION
            If attract(row, col + 1) > 0 Then attract(row, col + 1) = HIT_ATTRACTION
            If attract(row + 1, col) > 0 Then attract(row + 1, col) = HIT_ATTRACTION
        End If

        'update the hit-count
        hit(computer) = hit(computer) + 1
    
        'let's check if we sink another ship
        'this will only take place for LIEUTENANT's and higher
        If m_nLevel > ENSIGN Then
            the_big_one = big_ship()
        End If

        'reduce the ship-record over ship's in play
        ship_record(player(computer, row, col) - 12) = ship_record(player(computer, row, col) - 12) - 1

        'have we sunk another big-ship?
        If m_nLevel > ENSIGN Then
            If the_big_one <> big_ship() Then
                sunk = True
            End If
        End If

        '12 is a two-ship that has not been shot
        '22 is therefore a Hit on a two ship
        player(computer, row, col) = player(computer, row, col) + 10

        If hit(computer) = 16 Then
            computer_turn = COMPUTER_WINNER
            Exit Function
        End If

        'the ship is still floating, so another shot at this ship is
        'required
        If Not sunk Then
            computer_turn = COMPUTER_HIT
            Exit Function
        End If
    Else
        'it was a miss not close to a ship
        player(computer, row, col) = 2
    End If

    If m_nLevel > ENSIGN Then
        'setting the rim to 0, for sanity and
        'setting the rest to 1
        'thus resetting HIT_ATTRACTION to 1
        'this will ensure that LIEUTENANT is smarter than ENSIGN
        'removing the possibility of a sixth-shot and five-ship and so on
        For r = GRID_LOW - 1 To GRID_HIGH + 1
            For c = GRID_LOW - 1 To GRID_HIGH + 1
                If r = GRID_LOW - 1 Or r = GRID_HIGH + 1 Or c = GRID_LOW - 1 Or c = GRID_HIGH + 1 Then
                    attract(r, c) = 0
                End If
                If attract(r, c) > 0 Then
                    attract(r, c) = 1
                End If
            Next
        Next
    End If

    If m_nLevel > LIEUTENANT Then
        'the largest vertical ship each cell can hold
        For c = GRID_LOW To GRID_HIGH
            For r = GRID_LOW To GRID_HIGH
                count = 0
                If attract(r, c) > 0 Then
                    k = r
                    While attract(k, c) > 0
                        count = count + 1
                        k = k + 1
                    Wend
                    high = (count + count Mod 2) * 1.5 - 1 - count Mod 2
                    mycount = count
                    If count > 2 Then
                        direction = 1
                    Else
                        direction = 0
                    End If
                    
                    For k = r To r + count - 1
                        attract(k, c) = mycount
                        If count > 2 Then
                            If m_nLevel = ADMIRAL Then mycount = mycount + direction
                            If direction = 0 Then direction = -1
                            If mycount = high Then
                                If count Mod 2 = 0 And direction = 1 Then
                                    direction = 0
                                Else
                                    direction = -1
                                End If
                            End If
                        End If
                    Next
                    r = r + count
                End If
            Next
        Next

        'horizontal and largest
        For r = GRID_LOW To GRID_HIGH
            For c = GRID_LOW To GRID_HIGH
                count = 0
                If attract(r, c) > 0 Then
                    k = c
                    While attract(r, k) > 0
                        k = k + 1
                        count = count + 1
                    Wend
                    high = (count + count Mod 2) * 1.5 - 1 - count Mod 2
                    mycount = count
                    If count > 2 Then
                        direction = 1
                    Else
                        direction = 0
                    End If
                    For k = c To c + count - 1
                        If attract(r, k) < mycount Then attract(r, k) = mycount
                        If count > 2 Then
                            If m_nLevel = ADMIRAL Then mycount = mycount + direction
                            If direction = 0 Then direction = -1
                            If mycount = high Then
                                If count Mod 2 = 0 And direction = 1 Then
                                    direction = 0
                                Else
                                    direction = -1
                                End If
                            End If
                        End If
                    Next
                    c = c + count
                End If
            Next
        Next
    
        'Level Commander will avoid small space
        If m_nLevel = COMMANDER Then
            For r = GRID_LOW To GRID_HIGH
                For c = GRID_LOW To GRID_HIGH
                    If attract(r, c) < small_ship() Then
                        attract(r, c) = 0
                    End If
                    If attract(r, c) > 0 Then
                        attract(r, c) = 1
                    End If
                Next
            Next
        End If

        'Level Captain will shoot somewhere in the largest region
        'By setting any value above 5 to 5, will make the computer
        'smarter than Commander, but more stupid than the Rear Admiral
        If m_nLevel = CAPTAIN Then
            For r = GRID_LOW To GRID_HIGH
                For c = GRID_LOW To GRID_HIGH
                    If attract(r, c) > 5 Then
                        attract(r, c) = 5
                    End If
                Next
            Next
        End If

        'Level (Rear) Admiral
        'as you can see I haven't included any code for Admiral
        'because it isn't necessary!!!
        'There are a few  special conditions remaining
	'which I maybe will address on a later date, but for now
	'this is what I have come up with
    End If

    computer_turn = COMPUTER_MISS
    
    If sunk Then
        computer_turn = COMPUTER_SUNK
    End If
End Function
