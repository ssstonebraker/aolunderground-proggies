Attribute VB_Name = "Module3"
Option Explicit

'------------------------------------------
'Gets the item type
'----------------------------------------
Public Function itemType(intType As Integer) As String
  Dim strReturn As String
  Select Case intType
    Case 0
      strReturn = "EMPTY"
    Case 1
      strReturn = "ITEM"
    Case 2
      strReturn = "SWORD"
    Case 3
      strReturn = "SPEAR"
    Case 4
      strReturn = "BLADE"
    Case 5
      strReturn = "CLAW"
    Case 6
      strReturn = "WHIP"
    Case 7
      strReturn = "WRENCH"
    Case 8
      strReturn = "BOW"
    Case 9
      strReturn = "ARROW"
    Case 10
      strReturn = "HARP"
    Case 11
      strReturn = "ROD"
    Case 12
      strReturn = "STAFF"
    Case 13
      strReturn = "ROBE"
    Case 14
      strReturn = "ARMOR"
    Case 15
      strReturn = "SHIELD"
    Case 16
      strReturn = "HELMET"
    Case 17
      strReturn = "RING"
    Case Else
      strReturn = "ERROR"
      
  End Select
  itemType = strReturn
End Function

'W = wall      S = swoon
'P = poison    M = muddle
'U = mute      T = stone
'p = pig       s = small
'A = all       F = float
Public Function GetStatus(strList As String) As String
  Dim intFnd As Integer
  Dim strCurrent As String
  Dim strReturn As String
  intFnd% = InStr(strList, "|")
  Do While intFnd%
    DoEvents
    strCurrent$ = Mid(strList, 1, 1)
    strList$ = Mid(strList, intFnd% + 1, Len(strList))
    intFnd% = InStr(intFnd%, strList$, "|")
    If intFnd% Then
      strReturn$ = strReturn & GetStatusName(strCurrent$) & ", "
    Else
      strReturn$ = strReturn & GetStatusName(strCurrent$)
    End If
  Loop
  GetStatus = strReturn
End Function

Public Function GetStatusName(strChar As String) As String
  Dim strName As String
  Select Case strChar
    Case "W"
      strName = "Wall"
    Case "S"
      strName = "Swoon"
    Case "A"
      strName = "All"
    Case "P"
      strName = "Poison"
    Case "M"
      strName = "Muddle"
    Case "U"
      strName = "Mute"
    Case "T"
      strName = "Stone"
    Case "p"
      strName = "Pig"
    Case "F"
      strName = "Float"
  End Select
  GetStatusName = strName
End Function

'Z = spirit   G = giants
'U = undead   D = dragon
'L = land     W = water
'A = air      N = none

Public Function GetEnemy(strList As String) As String
  Dim intFnd As Integer
  Dim strCurrent As String
  Dim strReturn As String
  intFnd% = InStr(strList, "|")
  Do While intFnd%
    DoEvents
    strCurrent$ = Mid(strList, 1, 1)
    strList$ = Mid(strList, intFnd% + 1, Len(strList))
    intFnd% = InStr(intFnd%, strList$, "|")
    If intFnd% Then
      strReturn$ = strReturn & GetEnemyName(strCurrent$) & ", "
    Else
      strReturn$ = strReturn & GetEnemyName(strCurrent$)
    End If
  Loop
  GetEnemy = strReturn
End Function

Public Function GetEnemyName(strChar As String) As String
  Dim strName As String
  Select Case strChar
    Case "S"
      strName = "Spirit"
    Case "G"
      strName = "Giants"
    Case "U"
      strName = "Undead"
    Case "D"
      strName = "Dragon"
    Case "L"
      strName = "Land"
    Case "W"
      strName = "Water"
    Case "A"
      strName = "Air"
    Case "N"
      strName = "None"
    Case Else
      strName = "None"
  End Select
  GetEnemyName = strName
End Function

'N = none     D = dark
'S = sacred   F = fire
'I = ice      A = air
'X = all
Public Function GetElement(strList As String) As String
  Dim intFnd As Integer
  Dim strCurrent As String
  Dim strReturn As String
  intFnd% = InStr(strList, "|")
  Do While intFnd%
    DoEvents
    strCurrent$ = Mid(strList, 1, 1)
    strList$ = Mid(strList, intFnd% + 1, Len(strList))
    intFnd% = InStr(intFnd%, strList$, "|")
    If intFnd% Then
      strReturn$ = strReturn & GetElementName(strCurrent$) & ", "
    Else
      strReturn$ = strReturn & GetElementName(strCurrent$)
    End If
  Loop
  GetElement = strReturn
End Function

Public Function GetElementName(strChar As String) As String
  Dim strName As String
  Select Case strChar
    Case "D"
      strName = "Dark"
    Case "S"
      strName = "Sacred"
    Case "F"
      strName = "Fire"
    Case "I"
      strName = "Ice"
    Case "A"
      strName = "Air"
    Case "X"
      strName = "All"
    Case "N"
      strName = "None"
    Case Else
      strName = "None"
  End Select
  GetElementName = strName
End Function


