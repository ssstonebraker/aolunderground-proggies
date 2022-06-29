Attribute VB_Name = "Module2"
Option Explicit

Public CharClass(MAXSIZE) As Class4  'Holds info for each character
Public itemList(MAXSIZE)  As Class5  'Holds info for all items in game
Public weaponList(MAXSIZE) As Class1 'Holds info for weapons
Public armorList(MAXSIZE) As Class3  'Holds info for armor
Public enemyList(MAXSIZE) As Class2
Public GoldCount  As Double     'Gold count
Public timeCount As Integer     'Playing time

Public inventory(MAXSIZE) As Integer 'Your inventory
Public numOfItem(MAXSIZE) As Integer 'Number of the items possesd
                                'for each item.
Public charPos1 As Integer      'Current character in top position
Public charPos2 As Integer      '
Public charPos3 As Integer      '
Public charPos4 As Integer      '
Public charPos5 As Integer      '

Public rowOrder As Integer      'Row order, 3 infront, 2 in back, or vise versa

'-------------------------------
'First thing executed when the program is run
'-------------------------------
Sub Main()
  Dim i As Integer
  For i = 0 To 11
    Set CharClass(i) = New Class4
  Next i
  For i = 0 To 11 'item count
    Set itemList(i) = New Class5
  Next i
  For i = 0 To 39
    Set weaponList(i) = New Class1
  Next i
  For i = 0 To 29
    Set armorList(i) = New Class3
  Next i
 ' For i = 0 To 9
  'Next i
  'Load Form2
  Call LoadItemList
  Call LoadCharData
  Form1.Show
  
End Sub

'----------------------------------
'Loads the data for all the weapons, items, etc...
'----------------------------------
Private Sub LoadItemList()
  Dim filePath As String
  Dim n        As Integer 'counter
  Dim intiType As Integer, intWHEN As Integer
  Dim strSPECIAL As String, strName As String, strWHO As String, strATTRIB As String, strINFO As String
  Dim intTargetType As Integer, intHPadd As Integer, intHPPercent As Integer, intMPAdd As Integer, intMPPercent As Integer, strCure As String, intIfHP As Integer
  Dim dblPRICE As Double
  Dim strELEMENT As String, strENEMY As String, intATTACK As Integer, intATTPER As Integer
  Dim intDEFENSE As Integer, intDEFPER As Integer, intMAGICDEF As Integer, intMAGPER As Integer
  Dim inteTYPE As Integer, strWELEMENT As String, strSELEMENT As String, dblEXP As Double, dblGOLD As Double, dblHP As Double
  filePath = App.Path & "\itemList.txt"
  Open filePath For Input As #1
    For n = 0 To 11             'Loads indivual items
      Input #1, strName, intiType, strWHO, intTargetType, intHPadd, intHPPercent, intMPAdd, intMPPercent, strCure, intIfHP, intWHEN, strINFO, dblPRICE, strSPECIAL
        Call itemList(n).LoadAll(strName, intiType, strWHO, intTargetType, intHPadd, intHPPercent, intMPAdd, intMPPercent, strCure, intIfHP, intWHEN, strINFO, dblPRICE, strSPECIAL)
    Next n
    For n = 0 To 39
      Input #1, strName, intiType, strWHO, strELEMENT, strENEMY, intATTACK, intATTPER, intWHEN, strINFO, dblPRICE, strSPECIAL
        Call weaponList(n).LoadAll(strName, intiType, strWHO, strELEMENT, strENEMY, intATTACK, intATTPER, intWHEN, strINFO, dblPRICE, strSPECIAL)
    Next n
    For n = 0 To 28
      Input #1, strName, intiType, strWHO, strELEMENT, strENEMY, intDEFENSE, intDEFPER, intMAGICDEF, intMAGPER, intWHEN, strINFO, dblPRICE, strSPECIAL
        Call armorList(n).LoadAll(strName, intiType, strWHO, strELEMENT, strENEMY, intDEFENSE, intDEFPER, intMAGICDEF, intMAGPER, intWHEN, strINFO, dblPRICE, strSPECIAL)
    Next n
  Close 1
  n = 0
  Open App.Path & "\enemyList.txt" For Input As #1
    While Not EOF(1)
      DoEvents
      Set enemyList(n) = New Class2
      Input #1, strName, inteTYPE, strWELEMENT, strSELEMENT, intDEFENSE, intDEFPER, intMAGICDEF, intMAGPER, intATTACK, intATTPER, strINFO, dblEXP, dblGOLD, dblHP
        Call enemyList(n).LoadAll(strName, inteTYPE, strWELEMENT, strSELEMENT, intDEFENSE, intDEFPER, intMAGICDEF, intMAGPER, intATTACK, intATTPER, strINFO, dblEXP, dblGOLD, dblHP)
        n = n + 1
    Wend
  Close 1

End Sub

'---------------------------------
'Loads the data for each character.
'---------------------------------
Private Sub LoadCharData()
  Dim filePath As String
  Dim n        As Integer 'counter
  Dim strName As String
  Dim intCC As Integer, intCID As Integer, intLV As Integer, dblEXP As Double, dblNLV As Double, intMHP As Integer, intRHP As Integer, intMMP As Integer, intRMP As Integer, intHND As Integer
  Dim intSTR As Integer, intAGI As Integer, intVIT As Integer, intWIS As Integer, intWILL As Integer
  Dim intLHAND As Integer, intRHAND As Integer, intHELMET As Integer, intBODYARMOR As Integer, intRING As Integer
  Dim strACTION As String, strCONDITION As String
  Dim strSKILL1 As String, strSKILL2 As String
  filePath = App.Path & "\newGame.txt"
  Open filePath For Input As #1
    For n = 0 To 11             'Loads indivual characters
      Input #1, strName, intCC, intCID, intLV, dblEXP, dblNLV, intMHP, intRHP, intMMP, intRMP, intHND
        Call CharClass(n).LoadCharacter(strName, intCC, intCID, intLV, dblEXP, dblNLV, intMHP, intRHP, intMMP, intRMP, intHND)
      Input #1, intSTR, intAGI, intVIT, intWIS, intWILL
        Call CharClass(n).LoadStat(intSTR, intAGI, intVIT, intWIS, intWILL)
      Input #1, intLHAND, intRHAND, intHELMET, intBODYARMOR, intRING
        Call CharClass(n).LoadEquipment(intLHAND, intRHAND, intHELMET, intBODYARMOR, intRING)
      Input #1, strACTION, strCONDITION, strSKILL1, strSKILL2
        Call CharClass(n).LoadAction(strACTION, strCONDITION)
        Call CharClass(n).LoadSkill1(strSKILL1)
        Call CharClass(n).LoadSkill2(strSKILL2)
      Call CharClass(n).SetAllValues
    Next n
    
  Close
  
End Sub

'----------------------------------------------
'Gets the character type
'i.e. Cecil is a KNIGHT
'     Kain  is a DRAGOON, etc...
'----------------------------------------------
Public Function CharacterTypeName(intValue As Integer, strReturnShort As String) As String
  Dim strType As String
  
  Select Case intValue
    Case 0
      strType = "Knight"
      strReturnShort = "K"
    Case 1
      strType = "Dragoon"
      strReturnShort = "D"
    Case 2
      strType = "Caller"
      strReturnShort = "C"
    Case 3
      strType = "Sage"
      strReturnShort = "S"
    Case 4
      strType = "Karate Man"
      strReturnShort = "T"
    Case 5
      strType = "Black Wizard"
      strReturnShort = "Z"
    Case 6
      strType = "Bard"
      strReturnShort = "B"
    Case 7
      strType = "Engineer"
      strReturnShort = "E"
    Case 8
      strType = "Ninja"
      strReturnShort = "N"
    Case 9
      strType = "Lunarian"
      strReturnShort = "L"
    Case 10
      strType = "White Wizard"
      strReturnShort = "W"
  End Select
  CharType = strType
  
End Function

'--------------------------------------
'Prints the information of the character into
'the text file
'--------------------------------------
Public Sub PrintCharInfo(cValue As Integer) 'The character id
  Dim strName As String, strACTION As String, strCONDITION As String, strSKILL1 As String, strSKILL2 As String
  Dim intLHAND As Integer, intRHAND As Integer, intHELMET As Integer, intBODYARMOR As Integer, intRING As Integer
  Dim intMHP As Integer, intRHP As Integer, intMMP As Integer, intRMP As Integer, intLV As Integer
  Dim dblEXP As Double, dblNLV As Double
  Dim intSTR As Integer, intAGI As Integer, intVIT As Integer, intWIS As Integer, intWILL As Integer
  Dim intATT As Integer, intxA As Integer, intApc As Integer, intDEF As Integer, intxD As Integer, intDpc As Integer, intMD As Integer, intxM As Integer, intMpc As Integer
  Dim intATTACK As Integer
  Dim strValue As String
  
  strName$ = CharClass(cValue).Get_NAME()
  strACTION$ = CharClass(cValue).Get_ACTION
  strCONDITION$ = CharClass(cValue).Get_CONDITION
  strSKILL1$ = CharClass(cValue).Get_SKILL1
  strSKILL2$ = CharClass(cValue).Get_SKILL2
  intLHAND% = CharClass(cValue).Get_LHAND
  intRHAND% = CharClass(cValue).Get_RHAND
  intHELMET% = CharClass(cValue).Get_HELMET
  intBODYARMOR% = CharClass(cValue).Get_BODYARMOR
  intRING% = CharClass(cValue).Get_RING
  intMHP% = CharClass(cValue).Get_MHP
  intRHP% = CharClass(cValue).Get_RHP
  intMMP% = CharClass(cValue).Get_MMP
  intRMP% = CharClass(cValue).Get_RMP
  intLV% = CharClass(cValue).Get_LV
  dblEXP# = CharClass(cValue).Get_EXP
  dblNLV# = CharClass(cValue).Get_NLV
  intATTACK = CharClass(cValue).ATTACK(10, 65, 35, 12, 10)
  Call CharClass(cValue).Get_BASICstat(intSTR, intAGI, intVIT, intWIS, intWILL)
  Call CharClass(cValue).Get_STATS(intATT, intxA, intApc, intDEF, intxD, intDpc, intMD, intxM, intMpc)
  With Form1.Text1
    .Text = "Name:         " & strName$ & vbCrLf
    .Text = .Text & "Level:        " & intLV% & vbCrLf
    .Text = .Text & "Hit Points:   " & intRHP% & "/" & intMHP% & vbCrLf
    .Text = .Text & "Magic Points: " & intRMP% & "/" & intMMP% & vbCrLf
    .Text = .Text & "Total EXP:    " & dblEXP# & vbCrLf
    .Text = .Text & "NextLevel At: " & dblNLV# & vbCrLf
    .Text = .Text & "--Basic-Stats-" & vbCrLf
    .Text = .Text & "Strength:     " & intSTR & vbCrLf
    .Text = .Text & "Agility:      " & intAGI & vbCrLf
    .Text = .Text & "Vitality:     " & intVIT & vbCrLf
    .Text = .Text & "Wisdom:       " & intWIS & vbCrLf
    .Text = .Text & "Will:         " & intWILL & vbCrLf
    .Text = .Text & "-Weapon-Related-Stats-" & vbCrLf
    .Text = .Text & "Attack:             " & intATT & vbCrLf
    .Text = .Text & "Attack Multiplier:  " & intxA & vbCrLf
    .Text = .Text & "Attack Percent:     " & (intApc * 0.01) & "%" & vbCrLf
    .Text = .Text & "-Armor-Related-Stats--" & vbCrLf
    .Text = .Text & "Defense:            " & intDEF & vbCrLf
    .Text = .Text & "Defense Multiplier: " & (intxD * 0.01) & "%" & vbCrLf
    .Text = .Text & "Defense Percent:    " & intDpc & vbCrLf
    .Text = .Text & "Magic Resist:       " & intMD & vbCrLf
    .Text = .Text & "Magic Multiplier:   " & intxM & vbCrLf
    .Text = .Text & "Magic Percent:      " & (intMpc * 0.01) & "%" & vbCrLf
    .Text = .Text & "----------------------Armor-Info-----------------------------" & vbCrLf
    .Text = .Text & "Average Attack: " & intATTACK & vbCrLf
  
  End With
  
End Sub
'----------------------------------------
'Used to format a string, so it prints out nice.
'i.e.  If your printing out a list and want all the
'data to be lined up, this will add in number of spaces
'needed.
'Ex.  FormatString("String", 10)  will return
'  "String    "
'-------------------------------------------------
Private Function FormatString(strString As String, intLength As Integer) As String
  Dim strNew As String
  Dim intLen As Integer 'len of string
  Dim i As Integer      'counter
  strNew = strString
  intLen = Len(strNew)

  For i = intLen To intLength
    strNew = strNew & " "
  Next i
  FormatString = strNew
End Function

