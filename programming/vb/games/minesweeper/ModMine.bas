Attribute VB_Name = "ModMine"
Public currentindex As Integer 'dim the variable to hold the information as to which button is being checked
Public bombs As Integer 'dim the variable for the number of bombs surrounding a button

Sub winner()
Call bombfound 'call the bombfound sub.  This will make all the bombs visible to the user
MsgBox "Congratulations, You are a winner.  You discovered where all the bombs were without setting them off.", vbOKOnly + vbInformation, "You're a Winner!"
End Sub
Sub check4win()


Dim winindex As Integer 'declare the variable which will check the array of buttons to see if the user has won

For winindex = 1 To 100 Step 1 'this for..next loop will check each control and see if the game is won
    If frmMine.m(winindex).Caption = "" And frmMine.m(winindex).Tag <> "b" And frmMine.m(winindex).Visible = True Then 'check to see if a button has no caption.  If it doesn't and it doesn't have a bomb in it then the user still has a more buttons to uncover
        Exit Sub 'exit the sub.  This way you aren't checking every button.  If you find just one like the above description then the user still has to play.
    End If 'end the if function
Next 'end the for..next function

Call winner 'call the sub which tells the user they are a winner


End Sub

Sub bombfound()
Dim nextindex As Integer 'declare the variable which will hold the current button in the for..next loop
For nextindex = 1 To 100 Step 1 'move the array searching for any buttons with a "b" in the tag of it.  This will mark them as having a bomb.  It will also disable the controls as it goes so the user knows the game is over
    If frmMine.m(nextindex).Tag = "b" Then 'search for the tag that has a "b" in it
        frmMine.m(nextindex).Caption = "B" 'if it does then make the Caption of the current button have a "B" in it.  This is where you can put in a picture of a bomb or whatever you want.
        frmMine.m(nextindex).Visible = True 'make sure that the button with a bomb in it is visible to the user
    Else
        frmMine.m(nextindex).Visible = False 'if the button doesn't have a bomb then make it visible
    End If 'end the if function
    
    frmMine.m(nextindex).Enabled = False 'disable the button
Next 'end the for..next loop

End Sub
Sub babove()
On Error GoTo skip1: 'skip to the next step
If frmMine.m(currentindex - 10).Tag = "b" Then bombs = bombs + 1 'check the button below this one see if it has a bomb
skip1:

End Sub
Sub bbelow()
On Error GoTo skip2: 'skip to the next step
If frmMine.m(currentindex + 10).Tag = "b" Then bombs = bombs + 1 'check the button above this one see if it has a bomb
skip2:

End Sub
Sub bleft()

On Error GoTo skip4: 'skip to the next step
If frmMine.m(currentindex - 1).Tag = "b" Then bombs = bombs + 1 'check the button to the left of this one for a bomb
skip4:

End Sub
Sub bright()

On Error GoTo skip3: 'skip to the next step
If frmMine.m(currentindex + 1).Tag = "b" Then bombs = bombs + 1 'check the button to the right of this one for a bomb
skip3:

End Sub
Sub baboveleft()

On Error GoTo skip5: 'skip to the next step
If frmMine.m(currentindex - 11).Tag = "b" Then bombs = bombs + 1 'check the button above and to the left of this one see if it has a bomb
skip5:

End Sub
Sub baboveright()

On Error GoTo skip6: 'skip to the next step
If frmMine.m(currentindex - 9).Tag = "b" Then bombs = bombs + 1 'check the button above and to the right this one see if it has a bomb
skip6:
End Sub
Sub bbelowleft()

On Error GoTo skip8: 'skip to the next step
If frmMine.m(currentindex + 9).Tag = "b" Then bombs = bombs + 1 'check the button below and to the left this one see if it has a bomb
skip8:
End Sub
Sub bbelowright()

On Error GoTo skip7: 'skip to the next step
If frmMine.m(currentindex + 11).Tag = "b" Then bombs = bombs + 1 'check the button below and to the right of this one see if it has a bomb
skip7:
End Sub

