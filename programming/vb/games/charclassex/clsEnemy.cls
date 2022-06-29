VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NAME     As String  'Name of enemy
Private eTYPE    As Integer 'Flying, dragon, etc.
Private wELEMENT As String  'Elements weak
Private sELEMENT As String  'Elements strong against
Private DEFENSE  As Integer 'Defense level
Private DEFPER   As Integer 'Defense percent
Private MAGICDEF As Integer 'Magic defense
Private MAGPER   As Integer 'Magic percent
Private ATTACK   As Integer 'Attack power
Private ATTPER   As Integer 'Attack power
Private INFO     As String 'Info about item
Private EXP      As Double 'experience received
Private GOLD     As Double 'gold received
Private MAXHP    As Double 'max hp
Private RHP      As Double 'remaining hp

Public Sub LoadAll(strName As String, inteTYPE As Integer, strWELEMENT As String, strSELEMENT As String, _
          intDEFENSE As Integer, intDEFPER As Integer, intMAGICDEF As Integer, _
          intMAGPER As Integer, intATTACK As Integer, intATTPER As Integer, _
          strINFO As String, dblEXP As Double, dblGOLD As Double, dblHP As Double)
NAME = strName
eTYPE = inteTYPE
wELEMENT = strWELEMENT
sELEMENT = strSELEMENT
DEFENSE = intDEFENSE
DEFPER = intDEFPER
MAGICDEF = intMAGICDEF
MAGPER = intMAGPER
ATTACK = intATTACK
ATTPER = intATTPER
INFO = strINFO
EXP = dblEXP
GOLD = dblGOLD
MAXHP = dblHP
RHP = dblHP
End Sub

Public Function GetName() As String
  GetName = NAME
End Function
