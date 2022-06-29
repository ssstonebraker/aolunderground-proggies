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

Private firstName As String    'Their name.
Private CharType As Integer     'Knight, Wizard, etc.
Private charAssociate As Integer 'This is to keep track of the character, regardless of the name
Private level As Integer       'Characters current level
Private currentExp As Double   'current experience points
Private nextLevelExp As Double 'Next level experience points
Private currentHP As Integer   'current hit points
Private totalHP As Integer     'Total hit points
Private currentMP As Integer   'current magic points
Private totalMP As Integer     'Total magic points

Private lHandWeapon As Integer 'weapon in left hand
Private rHandWeapon As Integer 'Weapon in right hand
Private headArmor As Integer   'Helmet
Private bodyArmor As Integer   'Umm... Body armor
Private ringArmor As Integer   'A ring

Private mltStrength As Integer    'Mulitipliers:
Private mltEvade As Integer       'Based only on
Private mltMagicResist As Integer 'current level.  Armor
Private mltSpeed As Integer       'gets added on after.

Private finalStrength As Double
Private finalEvade As Double
Private finalmagicResist As Double
Private finalspeed As Double
'This is the character status including all weapons, armor, etc...
Private status As Double 'current status.  i.e. poisoned, etc.
Private STR As Double 'Strenth
Private evd As Double 'Evade
Private spd As Double 'speed
Private mResist As Double 'Magic resist
Private mPercent As Double 'Magic percent
Private aPower As Double 'Attack resist
Private aPercent As Double 'Attack/Hit percent
Private dResist As Double 'defense resist
Private dPercent As Double 'defense percent
Private eFpercent As Double 'element Fire percent
Private eWpercent As Double 'element Water percent
Private eApercent As Double 'element Ait percent
Private eEpercent As Double 'element Earth percent

'Calculate the status for all of the above
Public Sub CalcStatus()
  Dim currentType As Integer
  Dim currentCode As String
  Call CalcWeapon
  Call CalcArmor
  Call CalcRing
  finalStrength = (level + 1) * 2 + (aPower / 4) + mltStrength
  finalEvade = (mltSpeed / 2) + dResist
  finalmagicResist = mResist
  finalspeed = mltSpeed + (level / 4)
End Sub

Public Function Get_finalSpeed()
  Get_finalSpeed = finalspeed
End Function

Public Function Get_finalMagicResist()
  Get_finalMagicResist = finalmagicResist
End Function

Public Function Get_finalEvade()
  Get_finalEvade = finalEvade
End Function

Public Function Get_finalStrength()
  Get_finalStrength = finalStrength
End Function

Private Sub CalcArmor()
'ex. D5%30M4%5
  Dim defRst As Double, defPer As Double
  Dim majRst As Double, majPer As Double
  
  If headArmor > 0 Then
    Call powerArmor(headArmor, defRst, defPer, majRst, majPer)
  End If
  dResist = dResist + defRst
  dPercent = dPercent + defPer
  mResist = mResist + majRst
  mPercent = mPercent + majPer

  If bodyArmor > 0 Then
    Call powerArmor(bodyArmor, defRst, defPer, majRst, majPer)
  End If
  
  dResist = dResist + defRst
  dPercent = dPercent + defPer
  mResist = mResist + majRst
  mPercent = mPercent + majPer

End Sub
Private Sub CalcRing()
'1+3|2+1|F+10
  Dim fP As Double, wP As Double, eP As Double, aP As Double
  Dim st As Double, ev As Double, mR As Double, sp As Double
 
  If ringArmor > 0 Then
    Call powerRing(ringArmor, fP, wP, eP, aP, st, ev, mR, sp)
  End If
  eFpercent = eFpercent + fP
  eWpercent = eWpercent + wP
  eEpercent = eEpercent + eP
  eApercent = eApercent + aP
  STR = STR + st
  evd = evd + ev
  spd = spd + sp
  mResist = mResist + mR
  
End Sub

Private Sub CalcWeapon()
  Dim atkPwr As Double, atkPer As Double
  Dim D As Integer
  D = 0
  If theItemList(lHandWeapon).Get_type <= 3 And Not theItemList(lHandWeapon).Get_type = 0 Then
    Call powerWeapon(lHandWeapon, atkPwr, atkPer)
    D = D + 1
  End If
  aPercent = aPercent + atkPer
  aPower = aPower + atkPwr
  If theItemList(rHandWeapon).Get_type <= 3 And Not theItemList(rHandWeapon).Get_type = 0 Then
    Call powerWeapon(rHandWeapon, atkPwr, atkPer)
    D = D + 1
  End If
  If D = 2 Then
    aPower = aPower * 0.5
    aPercent = aPercent * 0.9
    atkPer = atkPer * 0.5
    atkPwr = atkPwr * 0.9
  End If
  aPercent = aPercent + atkPer
  aPower = aPower + atkPwr
End Sub

Public Function Get_totalMP() As Integer
  Get_totalMP = totalMP
End Function

Public Function Get_currentMP() As Integer
  Get_currentMP = currentMP
End Function

Public Function Get_totalHP() As Integer
  Get_totalHP = totalHP
End Function

Public Function Get_currentHP() As Integer
  Get_currentHP = currentHP
End Function

Public Function Get_nextLevelExp() As Double
  Get_nextLevelExp = nextLevelExp
End Function

Public Function Get_currentExp() As Double
  Get_currentExp = currentExp
End Function

Public Function Get_level() As Integer
  Get_level = level
End Function

Public Function Get_charAssociate() As Integer
  Get_charAssociate = charAssociate
End Function

Public Function Get_charType() As Integer
  Get_charType = CharType
End Function

Public Function Get_firstName() As String
  Get_firstName = firstName
End Function

Public Function Get_ringArmor() As Integer
  Get_ringArmor = ringArmor
End Function

Public Function Get_bodyArmor() As Integer
  Get_bodyArmor = bodyArmor
End Function

Public Function Get_headArmor() As Integer
  Get_headArmor = headArmor
End Function

Public Function Get_rHandWeapon() As Integer
  Get_rHandWeapon = rHandWeapon
End Function

Public Function Get_lHandWeapon() As Integer
  Get_lHandWeapon = lHandWeapon
End Function

Public Function Get_mltSpeed() As Integer
  Get_mltSpeed = mltSpeed
End Function

Public Function Get_mltMagicResist() As Integer
  Get_mltMagicResist = mltMagicResist
End Function

Public Function Get_mltEvade() As Integer
  Get_mltEvade = mltEvade
End Function

Public Function Get_mltStrength() As Integer
  Get_mltStrength = mltStrength
End Function

Public Sub LoadStatusInfo(mS As Integer, mEv As Integer, mMR As Integer, mSp As Integer)
  mltStrength = mS
  mltEvade = mEv
  mltMagicResist = mMR
  mltSpeed = mSp
End Sub

Public Sub LoadArmorInfo(lHW As Integer, rHW As Integer, hA As Integer _
               , bA As Integer, rA As Integer)
  ringArmor = rA
  bodyArmor = bA
  headArmor = hA
  rHandWeapon = rHW
  lHandWeapon = lHW
End Sub
               
Public Sub LoadCharInfo(fn As String, cT As Integer, cAS As Integer, LV As Integer, cE As Double _
              , nLE As Double, cHP As Integer, tHP As Integer, cMP As Integer _
              , tMP As Integer)
  totalMP = tMP
  currentMP = cMP
  totalHP = tHP
  currentHP = cHP
  nextLevelExp = nLE
  currentExp = cE
  level = LV
  CharType = cT
  firstName = fn
  charAssociate = cAS
End Sub

Public Sub Set_totalMP(nValue As Integer)
  totalMP = nValue
End Sub

Public Sub Set_currentMP(nValue As Integer)
  currentMP = nValue
End Sub

Public Sub Set_totalHP(nValue As Integer)
  totalHP = nValue
End Sub

Public Sub Set_currentHP(nValue As Integer)
  Get_currentHP = currentHP
End Sub

Public Sub Set_nextLevelExp(nValue As Double)
  nextLevelExp = nValue
End Sub

Public Sub Set_currentExp(nValue As Double)
  currentExp = nValue
End Sub

Public Sub Set_level(nValue As Integer)
  level = nValue
End Sub

Public Sub Set_charAssociate(nValue As Integer)
  charAssociate = nValue
End Sub

Public Sub Set_charType(nValue As Integer)
  CharType = nValue
End Sub

Public Sub Set_firstName(nValue As String)
  firstName = nValue
End Sub

Public Sub Set_ringArmor(nValue As Integer)
  ringArmor = nValue
End Sub

Public Sub Set_bodyArmor(nValue As Integer)
 bodyArmor = nValue
End Sub

Public Sub Set_headArmor(nValue As Integer)
  headArmor = nValue
End Sub

Public Sub Set_rHandWeapon(nValue As Integer)
   rHandWeapon = nValue
End Sub

Public Sub Set_lHandWeapon(nValue As Integer)
  lHandWeapon = nValue
End Sub

Public Sub Set_mltSpeed(nValue As Integer)
  mltSpeed = nValue
End Sub

Public Sub Set_mltMagicResist(nValue As Integer)
   mltMagicResist = nValue
End Sub

Public Sub Set_mltEvade(nValue As Integer)
  mltEvade = nValue
End Sub

Public Sub Set_mltStrength(nValue As Integer)
  mltStrength = nValue
End Sub

 

