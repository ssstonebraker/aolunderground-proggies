Attribute VB_Name = "Module1"
Option Explicit

Public Const CECIL = 0
Public Const KAIN = 1
Public Const EDGE = 2
Public Const ROSA = 3
Public Const RYDIA = 4
Public Const YANG = 5
Public Const CID = 6
Public Const PALOM = 7
Public Const POROM = 8
Public Const TALLAH = 9
Public Const EDWARD = 10
Public Const FUSOYA = 11

Public Characters(12) As Class1
Public GoldCount As Double
Public itemCount As Integer
Public items(30) As Integer
Public numOfItem(30) As Integer
Public char1 As Integer, char2 As Integer, char3 As Integer
Public char4 As Integer, char5 As Integer 'who is currently in the party

Sub Main2()
  Dim i As Integer
  Dim result As Integer
  For i = 0 To 11
    Set Characters(i) = New Class1
  Next i
  For i = 0 To 45
    Set theItemList(i) = New Class3
  Next i
  itemCount = 0
  Call LoadItem
  Load Form1
  result = InputBox("Do you want to start a new game? 1 for Yes.  0 for No.")
  If result = 0 Then
    Call LoadGame(0)
  Else
    Call LoadGame
  End If
  Call LoadList
  Form1.Show
  
End Sub

Public Function CharType(typeNum As Integer) As String
  Dim strType As String
  
  Select Case typeNum
  Case 0
    strType = "Knight"
  Case 1
    strType = "Dragoon"
  Case 2
    strType = "Ninja"
  Case 3
    strType = "White Wizard"
  Case 4
    strType = "Caller"
  Case 5
    strType = "Karate"
  Case 6
    strType = "Fat Guy"
  Case 7
    strType = "Black Wizard"
  Case 8
    strType = "Bard"
  End Select
  CharType = strType
    
End Function

Private Sub LoadList()
  Dim i As Integer
  For i = 0 To itemCount - 1
    Form1.List1.AddItem theItemList(items(i)).Get_NAME & "  " & numOfItem(i)
  Next i
End Sub
Private Function LoadItem()

  Dim n As Integer
  Dim nm As String, cd As String, ef As String
  Dim tp As Integer, cs As Integer
  Call theItemList(n).Set_All("Empty", 0, "Empty", "", 0)
  Open App.Path & "\itemList.txt" For Input As #1
    For n = 1 To 45
      Input #1, nm, tp, ef, cd, cs
      Call theItemList(n).Set_All(nm, tp, ef, cd, cs)
    Next n
  Close #1
  
End Function
Public Function LoadGame(Optional newGame As Integer)
  Dim fn As String
  Dim cT As Integer, cA As Integer, LV As Integer, cHP As Integer, tHP As Integer, cMP As Integer, tMP As Integer
  Dim cE As Double, nLE As Double
  Dim lHW As Integer, rHW As Integer, hA As Integer, bA As Integer, rA As Integer
  Dim mS As Integer, mEv As Integer, mMR As Integer, mSp As Integer
  Dim filePath As String
  Dim n As Integer 'counter
  If Not newGame = 0 Then
    filePath = App.Path & "\game1.txt"
  Else
    filePath = App.Path & "\newGameData.txt"
  End If
  Open filePath For Input As #1
    For n = 0 To 11
      Input #1, fn, cT, cA, LV, cE, nLE, cHP, tHP, cMP, tMP
      Input #1, lHW, rHW, hA, bA, rA
      Input #1, mS, mEv, mMR, mS
      Call Characters(n).LoadCharInfo(fn, cT, cA, LV, cE, nLE, cHP, tHP, cMP, tMP)
      Call Characters(n).LoadArmorInfo(lHW, rHW, hA, bA, rA)
      Call Characters(n).LoadStatusInfo(mS, mEv, mMR, mSp)
    Next n
    Input #1, char1, char2, char3, char4, char5
    Form1.lstCurrent.AddItem Form1.lstTotalChar.List(char1)
    Form1.lstCurrent.AddItem Form1.lstTotalChar.List(char2)
    Form1.lstCurrent.AddItem Form1.lstTotalChar.List(char3)
    Form1.lstCurrent.AddItem Form1.lstTotalChar.List(char4)
    Form1.lstCurrent.AddItem Form1.lstTotalChar.List(char5)
    Input #1, GoldCount
    Do Until EOF(1)
      DoEvents
      Input #1, items(itemCount), numOfItem(itemCount)
      itemCount = itemCount + 1
    Loop
  Close #1
  
  For n = 0 To 11
    Call Characters(n).CalcStatus
  Next n
  
End Function

Public Sub PrintInfo(char As Integer)
  Dim LHAND As Integer, RHAND As Integer
  Dim HELMET As Integer, armor As Integer, RING As Integer
    
  Form1.Text1.Text = Characters(char).Get_firstName & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Type: " & CharType(Characters(char).Get_charType) & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Level: " & Characters(char).Get_level & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Exp: " & Characters(char).Get_currentExp & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Next Level: " & Characters(char).Get_nextLevelExp & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "HP: " & Characters(char).Get_currentHP & "/" & Characters(char).Get_totalHP & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "MP: " & Characters(char).Get_currentMP & "/" & Characters(char).Get_totalMP & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "LeftHand: " & theItemList(Characters(char).Get_lHandWeapon).Get_NAME & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "RightHand: " & theItemList(Characters(char).Get_rHandWeapon).Get_NAME & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Helmet: " & theItemList(Characters(char).Get_headArmor).Get_NAME & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Armor: " & theItemList(Characters(char).Get_BODYARMOR).Get_NAME & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Ring: " & theItemList(Characters(char).Get_ringArmor).Get_NAME & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Strength: " & Characters(char).Get_finalStrength & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Evade: " & Characters(char).Get_finalEvade & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "MagicResist: " & Characters(char).Get_finalMagicResist & vbCrLf
  Form1.Text1.Text = Form1.Text1.Text & "Speed: " & Characters(char).Get_finalEvade & vbCrLf

End Sub

Public Sub powerWeapon(theWeapon As Integer, attackPower As Double, attackPercent As Double)
  Dim theCode As String
  Dim fnd As Integer
  theCode = theItemList(theWeapon).Get_code
  fnd = InStr(theCode, "%")
  attackPower = (Mid(theCode, 2, fnd - 2))
  attackPercent = Mid(theCode, fnd + 1, Len(theCode))
  
End Sub

Public Sub powerArmor(theArmor As Integer, defenseResist As Double, defensePercent As Double, magicResist As Double, magicPercent As Double)
  Dim theCode As String
  Dim fnd As Integer
  Dim fndNext As Integer
  'ex. D5%30M4%5
  theCode = theItemList(theArmor).Get_code
  fnd = InStr(theCode, "D")
  fndNext = InStr(theCode, "%")
  
  defenseResist = Mid(theCode, fnd + 1, fndNext - 2)
  
  fnd = InStr(theCode, "M")
  
  defensePercent = Mid(theCode, fndNext + 1, fnd - 4)
  theCode = Mid(theCode, fnd + 1, Len(theCode))
  
  fnd = InStr(theCode, "%")
  magicResist = Mid(theCode, 1, fnd - 1)
  magicPercent = Mid(theCode, fnd + 1, Len(theCode))
End Sub

Public Sub powerRing(theRing As Integer, fP As Double, wP As Double, eP As Double, aP As Double, st As Double, ev As Double, mp As Double, sp As Double)
'1+10|2+10|3+5|E+2
  Dim currentEffect As String
  Dim stat As String
  Dim doWhat As String
  Dim theCode As String
  Dim val As Double
  Dim fnd As Integer
  Dim pos As Integer
  theCode$ = theItemList(theRing).Get_code
  pos = InStr(theCode, "|")
  Do While pos <> 0
    DoEvents
    stat = Mid(theCode, 1, 1)
    doWhat = Mid(theCode, 2, 1)
    val = CDbl(Mid(theCode, 3, pos - 3))
    Call addCorrect(stat, doWhat, val, fP, wP, eP, aP, st, ev, mp, sp)
    theCode$ = Mid(theCode, pos + 1, Len(theCode$))
    pos = InStr(theCode, "|")
    
  Loop
    stat = Mid(theCode, 1, 1)
    doWhat = Mid(theCode, 2, 1)
    val = Mid(theCode, 3, Len(theCode))
    Call addCorrect(stat, doWhat, val, fP, wP, eP, aP, st, ev, mp, sp)

End Sub

Public Sub addCorrect(stat As String, doWhat As String, val As Double, fP As Double, wP As Double, eP As Double, aP As Double, st As Double, ev As Double, mp As Double, sp As Double)
  
  If doWhat = "-" Then
    val = val * -1
  End If
  Select Case stat
    Case "1"
      st = st + val
    Case "2"
      ev = ev + val
    Case "3"
      mp = mp + val
    Case "4"
      sp = sp + val
    Case "5"
      st = st + val
      ev = ev + val
      mp = mp + val
      sp = sp + val
    Case "F"
      fP = fP + val
    Case "W"
      wP = wP + val
    Case "A"
      aP = aP + val
    Case "E"
      eP = eP + val
    Case "X"
      fP = fP + val
      aP = aP + val
      wP = wP + val
      eP = eP + val
  End Select
End Sub
