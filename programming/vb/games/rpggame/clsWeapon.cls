VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class represents a weapon
Dim IDNum As Integer
Dim Name As String
Dim Desc As String
Dim MeleeDam As Integer
Dim PiercingDam As Integer
Dim MPPlus As Integer
Dim HPPlus As Integer
Dim INTPlus As Integer
Dim STAPlus As Integer
Dim STRPlus As Integer
Dim ACPlus As Integer
Dim weight As Integer

