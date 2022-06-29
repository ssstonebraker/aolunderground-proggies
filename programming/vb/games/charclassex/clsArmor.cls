VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NAME     As String  'Name of weapon
Private iTYPE    As Integer 'spear, sword, whip, etc
Private WHO      As String  'Who can use
Private ELEMENT  As String  'Elements armor is stong aginst
Private ENEMY    As String  'Enemys strong/weak against
Private DEFENSE  As Integer 'Defense level
Private DEFPER   As Integer 'Defense percent
Private MAGICDEF As Integer 'Magic defense
Private MAGPER   As Integer 'Magic percent
Private REQ      As Integer  'Not important
Private INFO     As String 'Info about item
Private PRICE    As Double 'Price to buy. 0 means can't be dropped, soled, etc.
Private SPECIAL  As String 'Anything special to go with it.

Public Sub LoadAll(strName As String, intiType As Integer, strWHO As String, strELEMENT As String, strENEMY As String, intDEFENSE As Integer, intDEFPER As Integer, intMAGICDEF As Integer, intMAGPER As Integer, intREQ As Integer, strINFO As String, dblPRICE As Double, strSPECIAL As String)
  Dim itmX As ListItem
  NAME = strName$
  iTYPE = intiType%
  WHO = strWHO$
  ELEMENT = strELEMENT
  ENEMY = strENEMY
  DEFENSE = intDEFENSE
  DEFPER = intDEFPER
  MAGICDEF = intMAGICDEF
  MAGPER = intMAGPER
  REQ = intREQ%
  INFO = strINFO$
  PRICE = dblPRICE#
  SPECIAL = strSPECIAL$
  Set itmX = Form2.ListView1.ListItems.Add(, , NAME)
  itmX.SubItems(1) = itemType(iTYPE)
  itmX.SubItems(2) = strWHO
  itmX.SubItems(3) = ELEMENT
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

Public Function Get_DEFENSE() As Integer
  Get_DEFENSE = DEFENSE
End Function

Public Function Get_DEFPER() As Integer
  Get_DEFPER = DEFPER
End Function

Public Function Get_MAGICDEF() As Integer
  Get_MAGICDEF = MAGICDEF
End Function

Public Function Get_MAGPER() As Integer
  Get_MAGPER = MAGPER
End Function

Public Function Get_ELEMENT() As String
  Get_ELEMENT = ELEMENT
End Function

Public Function Get_ENEMY() As String
  Get_ENEMY = ENEMY
End Function
