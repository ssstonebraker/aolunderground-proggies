Attribute VB_Name = "modScramble"
Public Function RandomNum(BottomNum As Integer, TopNum As Integer) As Integer
Dim a As Integer
Randomize
a% = Int((TopNum% - BottomNum% + 1) * Rnd + BottomNum%)
RandomNum% = a%
End Function

Public Sub Remove(lst As ListBox, What As String)
For i& = 0 To lst.ListCount - 1
strIt$ = lst.List(i&)
If LCase(strIt$) = LCase(What$) Then
lst.RemoveItem i&
End If
Next i&
End Sub

Public Sub Add(lst As ListBox, What As String)
For i& = 0 To lst.ListCount - 1
strIt$ = lst.List(i&)
If LCase(strIt$) = LCase(What$) Then
Exit Sub
End If
Next i&
lst.AddItem What$
End Sub

Public Function WhichTeam(lst1 As ListBox, lst2 As ListBox, Who As String) As String
For i& = 0 To lst1.ListCount - 1
strIt$ = lst1.List(i&)
If LCase(strIt$) = LCase(Who$) Then
WhichTeam$ = "1"
Exit Function
End If
Next i&
For x& = 0 To lst2.ListCount - 1
strIt$ = lst2.List(x&)
If LCase(strIt$) = LCase(Who$) Then
WhichTeam$ = "2"
Exit Function
Else
WhichTeam$ = "0"
End If
Next x&
End Function

Public Function Scramble(Word As String) As String
Word$ = ReplaceC(Word$)
For i& = 6 To 20
If Len(Word$) = i& Then
GoTo morethan5
End If
Next i&
If Len(Word$) = "0" Then
Exit Function
End If
If Len(Word$) = "3" Then
Scramble$ = Reverse(Word$)
Exit Function
End If
If Len(Word$) = "2" Then
Scramble$ = Reverse(Word$)
Exit Function
End If
If Len(Word$) = "1" Then
Scramble$ = Word$
Exit Function
End If
If Len(Word$) = "4" Then
  strL$ = Left(Word$, 1)
  strR$ = Right(Word$, 1)
  strMid$ = Mid(Word$, 2, Len(Word$) - 2)
  strNum$ = RandomNum(1, 3)
   If strNum$ = "1" Then
  strRev$ = Reverse(strMid$)
  Scramble$ = strR$ & strRev$ & strL$
   End If
   If strNum$ = "2" Then
   Scramble$ = strR$ & strL$ & strMid$
   End If
   If strNum$ = "3" Then
   Scramble$ = strMid$ & strL$ & strR$
   End If
End If
If Len(Word$) = "5" Then
  strL$ = Left(Word$, 1)
  strR$ = Right(Word$, 1)
  strMid$ = Mid(Word$, 2, Len(Word$) - 2)
  strRev$ = Reverse(strMid$)
  strNum$ = RandomNum(1, 3)
   If strNum$ = "1" Then
    strFin$ = strMid$ & strR$ & strL$
   End If
   If strNum$ = "2" Then
    strFin$ = strR$ & strMid$ & strL$
   End If
   If strNum$ = "3" Then
    strFin$ = strL$ & strR$ & strMid$
   End If
  Scramble$ = strFin$
  Exit Function
End If
morethan5:
  strL$ = Left(Word$, 2)
  strR$ = Right(Word$, 2)
  strR$ = Left(strR$, 1)
  strMid$ = Mid(Word$, 3, Len(Word$) - 3)
  strR$ = Right(Word$, 1)
  strNum$ = RandomNum(1, 2)
   If strNum$ = "1" Then
    strRev$ = Reverse(strMid$)
    strNum2$ = RandomNum(1, 4)
   If strNum2$ = "1" Then
    strFin$ = strR & strMid$ & strR2$ & strL$
    Scramble$ = strFin$
   End If
   If strNum2$ = "2" Then
    strFin$ = strR2$ & strR & strL$ & strMid$
    Scramble$ = strFin$
   End If
   If strNum2$ = "3" Then
    strFin$ = strMid$ & strR$ & strL$ & strR2$
    Scramble$ = strFin$
   End If
   If strNum2$ = "4" Then
    strFin$ = strR$ & strR2$ & strMid$ & strL$
    Scramble$ = strFin$
   End If
  End If
  If strNum$ = "2" Then
  strNum2$ = RandomNum(1, 4)
   If strNum2$ = "1" Then
    strFin$ = strR$ & strMid$ & strR2$ & strL$
    Scramble$ = strFin$
   End If
   If strNum2$ = "2" Then
    strFin$ = strR2$ & strR & strL$ & strMid$
    Scramble$ = strFin$
   End If
   If strNum2$ = "3" Then
    strFin$ = strMid$ & strR$ & strL$ & strR2$
    Scramble$ = strFin$
   End If
   If strNum2$ = "4" Then
    strFin$ = strR$ & strR2$ & strMid$ & strL$
    Scramble$ = strFin$
   End If
  End If
Exit Function
End Function

Public Function Reverse(Word As String) As String
For i& = 1 To Len(Word$)
strIt$ = Mid(Word$, i&, 1)
strRev$ = strIt$ & strRev$
Next i&
Reverse$ = strRev$
End Function

Public Function ReplaceC(TextToSearch As String) As String
If InStr(TextToSearch$, Find$) = 0 Then
ReplaceChr$ = TextToSearch$
End If
For i& = 1 To Len(TextToSearch$)
strMid$ = Mid(TextToSearch$, i&, 1)
If strMid$ = " " Then
strMid$ = ""
End If
strMid2$ = strMid2$ & strMid$
ReplaceC$ = strMid2$
Next i&
End Function
