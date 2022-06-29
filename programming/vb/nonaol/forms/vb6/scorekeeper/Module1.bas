Attribute VB_Name = "Module1"
Sub AddScore(Lst As ListBox, HowMuch As Integer, Winner As String, Sort As Boolean)
Dim Score As String, ScoreInt As Integer
Dim NewScore As String, LastWinner As String
For i = 0 To Lst.ListCount - 1
For c = 1 To Len(Lst.List(i))
ab = Mid(Lst.List(i), c, 1)
If ab = "-" Then
LastWinner = Left(Lst.List(i), c - 2)
If LCase(LastWinner) = LCase(Winner) Then
ScoreInt = i
Score = Lst.List(i)
GoTo bone
End If
End If
Next c
Next i

GoTo bone2:

bone:
For i = 1 To Len(Score)
ab = Mid(Score, i, 1)
If ab = "-" Then
NewScore = Mid(Score, i + 2, Len(Score) - i)
NewScore = Val(NewScore) + HowMuch
Lst.RemoveItem (ScoreInt)
Lst.AddItem "" + Winner + " - " + NewScore + ""
If Sort = True Then
SortScore Lst
End If
Exit Sub
End If
Next i

bone2:
Lst.AddItem "" + Winner + " - " + CStr(HowMuch) + ""
If Sort = True Then
SortScore Lst
End If
End Sub
Sub SortScore(Lst As ListBox)
Dim TopScore As Integer, Score As Integer
Dim Person As String, Winners As String
Dim TopPerson As String, TopInt As Integer
Dim Coin As String
Winners = ""

restart:
TopScore = 0
For i = 0 To Lst.ListCount - 1
    For c = 1 To Len(Lst.List(i))
    ab = Mid(Lst.List(i), c, 1)
        If ab = "-" Then
        Score = Mid(Lst.List(i), c + 2, Len(Lst.List(i)) - c)
        Person = Left(Lst.List(i), c - 2)
            If Score >= TopScore Then
            TopPerson = Person
            TopScore = Score
            TopInt = i
            End If
        End If
    Next c
Next i
Lst.RemoveItem (TopInt)
Winners = Winners + "" + TopPerson + " - " + CStr(TopScore) + "" + ","
If Lst.ListCount <> 0 Then GoTo restart

For i = 1 To Len(Winners)
ab = Mid(Winners, i, 1)
If ab = "," Then
Lst.AddItem Coin
Coin = ""
Else
Coin = Coin + ab
End If
Next i
End Sub
