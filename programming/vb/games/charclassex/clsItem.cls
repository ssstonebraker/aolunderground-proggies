VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NAME    As String 'Name of item
Private iTYPE  As Integer 'Item type as integer: shield, sword, item, etc
Private WHO     As String 'Who can use
Private WHEN   As Integer 'When available...battle, on save spot, etc
Private INFO    As String 'Info about item
Private PRICE   As Double 'Price to buy. 0 means can't be dropped, soled, etc.
Private SPECIAL As String 'Anything special to go with it.

'Stats by item  - what should they do
Private TargetType As Integer 'Target type.  0 single, 1 all, 2 all or single
Private HPadd      As Integer 'Increase HP
Private HPPercent  As Integer 'Increase HP by percent
Private MPadd      As Integer 'Increase MP
Private MPPercent  As Integer 'Increase MP by percent
Private ifHP       As Integer '0 = HP must be 0, 1 = HP must be > 0, 2 = Doesn't matter
Private Cure       As String  'Which status to cure

Public Sub LoadAll(strNAME As String, intiType As Integer, strWHO As String, intTargetType As Integer, intHPadd As Integer, intHPPercent As Integer, intMPAdd As Integer, intMPPercent As Integer, strCure As String, intIfHP As Integer, intWHEN As Integer, strINFO As String, dblPRICE As Double, strSPECIAL As String)
  Dim itmX As ListItem
  NAME = strNAME$
  iTYPE = intiType%
  WHO = strWHO$
  WHEN = intWHEN%
  INFO = strINFO$
  PRICE = dblPRICE#
  SPECIAL = strSPECIAL$
  TargetType = intTargetType
  HPadd = intHPadd
  HPPercent = intHPPercent
  MPadd = intMPAdd
  MPPercent = intMPPercent
  Cure = strCure
  ifHP = intIfHP
  Set itmX = Form2.ListView1.ListItems.Add(, , NAME)
  itmX.SubItems(1) = itemType(iTYPE)
  itmX.SubItems(2) = strWHO
  itmX.SubItems(3) = TargetType & "|" & HPadd & "|" & HPPercent & "|" & MPadd & "|" & MPPercent & "|" & ifHP
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

Public Sub Set_TargetType(intSet As Integer)
  TargetType = intSet
End Sub

Public Sub Set_HPadd(intSet As Integer)
  HPadd = intSet
End Sub

Public Sub Set_MPadd(intSet As Integer)
  MPadd = intSet
End Sub

Public Function Get_HPadd() As Integer
  Get_HPadd = HPadd
End Function

Public Function Get_MPadd() As Integer
 Get_MPadd = MPadd
End Function

Public Function Get_MPPercent() As Integer
 Get_MPPercent = MPPercent
End Function

Public Function Get_HPPercent() As Integer
 Get_HPPercent = HPPercent
End Function

Public Function Get_TargetType() As Integer
  Get_TargetType = TargetType
End Function

Public Function Get_Cure() As String
  Get_Cure = Cure
End Function
