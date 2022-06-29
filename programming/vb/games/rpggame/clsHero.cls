VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class is for describing the hero

'private member variables
Public itsClass As Integer      'the hero's character class (Mage, Warrior, etc.)
Public itsRace As Integer       'the hero's race (Elf, Human, Dwarf etc.)
Public itsName As String        'the hero's name
Public itsHP As Integer         'the hero's hit points
Public itsMP As Integer         'the hero's magic points
Public itsExperience As Integer   'the hero's current experience points
Public itsLevel As Integer        'the hero's current level
Public itsNextLevel As Integer    'the experience needed to reach the next level
Public itsIntelligence As Integer 'the hero's intelligence
Dim itsBaseStats(0 To 4) As Integer 'the hero's base stats
Dim itsCurStats(0 To 4) As Integer  'the hero's current stats
Dim itsGold As Double               'holds the hero's gold
'Dim itsInventory(0 To 20) As New clsItmes      'the hero's inventory
Dim MaxItems As Integer         'holds the maximum inventory items
Dim itsWieght As Integer        'holds the character's current weight
Dim itsMaxWeight As Integer     'holds the character's current maximum wieght
Dim itsHeight As Integer        'holds the character's height


'for accessing the arrays of stats
Public Function setBaseStat(stat As Integer, value As Integer)
    itsBaseStats(stat) = value
End Function

Public Function getBaseStat(stat As Integer) As Integer
    getBaseStat = itsBaseStats(stat)
End Function

Public Function setCurStat(stat As Integer, value As Integer)
    itsCurStats(stat) = value
End Function

Public Function getCurStat(stat As Integer) As Integer
    getCurStat = itsCurStats(stat)
End Function

