VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NAME    As String  'Name of weapon
Private iTYPE   As Integer 'spear, sword, whip, etc
Private WHO     As String  'Who can use
Private ELEMENT As String  'What element it has
Private ENEMY   As String  'What enemy type it is stong agaist
Private ATTACK  As Integer 'Defense
Private ATTPER  As Integer 'Defense Percent
Private REQ    As Integer  '1 hand, 2 hands, arrows or bow
Private INFO    As String  'Info about item
Private PRICE   As Double  'Price to buy. 0 means can't be dropped, soled, etc.
Private SPECIAL As String  'Anything special to go with it.

Public Sub LoadAll(strName As String, intiType As Integer, strWHO As String, strELEMENT As String, strENEMY As String, intATTACK As Integer, intATTPER As Integer, intREQ As Integer, strINFO As String, dblPRICE As Double, strSPECIAL As String)
  Dim itmX As ListItem
  NAME = strName$
  iTYPE = intiType%
  WHO = strWHO$
  ELEMENT = strELEMENT
  ENEMY = strENEMY
  ATTACK = intATTACK
  ATTPER = intATTPER
  REQ = intREQ%
  INFO = strINFO$
  PRICE = dblPRICE#
  SPECIAL = strSPECIAL$
  'This next command puts the information into the
  'ListView on form2
  Set itmX = Form2.ListView1.ListItems.Add(, , NAME)
  itmX.SubItems(1) = itemType(iTYPE)
  itmX.SubItems(2) = strWHO
  itmX.SubItems(3) = ELEMENT & ENEMY & ATTACK & ATTPER
  itmX.SubItems(4) = INFO
  itmX.SubItems(5) = "$" & PRICE

End Sub

Public Function Get_NAME() As String
  Get_NAME = NAME
End Function

Public Function Get_iTypeValue() As Integer
  Get_iTypeValue = iTYPE
End Function

Public Function Get_iType() As String
  Get_iType = itemType(iTYPE)
End Function

Public Function Get_INFO() As String
  Get_INFO = INFO
End Function

Public Function Get_ELEMENT() As String
  Get_ELEMENT = ELEMENT
End Function

Public Function Get_ENEMY() As String
  Get_ENEMY = ENEMY
End Function

Public Function Get_ATTACK() As Integer
  Get_ATTACK = ATTACK
End Function

Public Function Get_ATTPER() As Integer
  Get_ATTPER = ATTPER
End Function
