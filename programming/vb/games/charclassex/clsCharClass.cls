VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'These stats are not derived from any other stat
'The natural abilitys
Private STR  As Integer 'strength
Private AGI  As Integer 'agility
Private VIT  As Integer 'vitality
Private WIS  As Integer 'wisdom
Private WILL As Integer 'will

'These stats are completly determined by the basic
'stats above, level, and what is equiped.
Private ATT As Integer 'attack value
Private Apc As Integer 'hit percent
Private xA  As Integer 'attack multiplier
Private DEF As Integer 'defense value
Private xD  As Integer 'defense multiplier
Private Dpc As Integer 'defend %
Private MD  As Integer 'Magic resistance value
Private xM  As Integer 'magic resistance multiplier
Private Mpc As Integer 'magic evasion
Private DMG As Integer 'Special attack that ignores targets armor or resistance
Private LB  As Integer 'Level bonus
Private SB  As Integer 'Stength bonus
Private VB  As Integer 'Vitality bonus
Private WB  As Integer 'Weapon bonus
Private B   As Integer 'Black magic multiplier
Private W   As Integer 'White magic multiplier
Private WP1 As Integer 'Weapon Power
Private WP2 As Integer 'Weapon Power2 - if ambidextrous
Private rB  As Integer 'Row bonus in battle

'The following are damage attributes
'Defense against
Private dA As Integer 'Air
Private dF As Integer 'Fire
Private dD As Integer 'Darkness
Private dC As Integer 'Sacred
Private dI As Integer 'Ice
'Attack with
Private aA As Integer 'Air
Private aF As Integer 'Fire
Private aD As Integer 'Darkness
Private aC As Integer 'Sacred
Private aI As Integer 'Ice
'Race Multiplier
Private rD As Integer 'Dragon
Private rL As Integer 'Land
Private rW As Integer 'Water
Private rA As Integer 'Air
Private rU As Integer 'Undead
Private rG As Integer 'Giants
Private rZ As Integer 'Spirits

'Other stats
Private LV  As Integer 'Current level
Private MHP As Integer 'Max hit points
Private RHP As Integer 'Remaining hit points
Private MMP As Integer 'Max magic points
Private RMP As Integer 'Remaining magic points
Private EXP As Double 'Total experience
Private NLV As Double 'Points needed until next level

'More important info
Private CC        As Integer 'Character class i.e Knight, WhiteWizard, etc.
Private NAME      As String  'Character name - changable
Private CID       As Integer 'Character ID
Private ACTION    As String  'Action available in battle
Private CONDITION As String  'POISON, CURSE, SWOON, etc.
Private SKILL1    As String  'Holds list of skills. i.e magic
Private SKILL2    As String  'Holds second list of skills.  i.e. if black and white magic user.
Private HND       As Integer 'Handed. 0 = Right, 1 = Left, 2 = Both

'Equipment On Character
Private LHAND     As Integer 'Left hand
Private RHAND     As Integer 'Right hand
Private HELMET    As Integer
Private BODYARMOR As Integer
Private RING      As Integer

Private cA  As Integer 'counter attack
Private cM  As Integer 'counter attack magic

'Character class table------------------
'0 - Knight (Cecil)
'1 - Dragoon (Kain)
'2 - caller (Rydia)
'3 - Sage (Tellah)
'4 - Karate Man (Yang)
'5 - Black Wizard (Porom)
'6 - Bard (Edward)
'7 - Engineer (Cid)
'8 - Ninja (Edge)
'9 - Lunarian (FuSoYa)
'10 - White Wizard (Palom, Rosa)

'Sets the values of the stats
'Note:  these algorithms are not perfect and
'if you game is going to be, you'll need good
'algorithms to determin stats for characters and
'enemys.
Public Sub SetAllValues() 'Non changing values
  LB = LV / 4   'Level bonus: level / 4
  SB = STR / 4  'Strength bonus
  VB = VIT / 2  'Vitality bonus
  WB = GetWeaponBonus 'weapon hit percent
  ATT = GetWeapon()
  xA = (STR / 8) + (AGI / 16) + 1
  Apc = 50 + (LV / 4) + WB
  DEF = GetDefense
  xD = GetDefenseMult  '[ AGI/8 ] + ( [ LV/16 ] if a shield is equipped)
  Dpc = GetDefensePercent
  MD = GetMagicDef
  xM = ((WIS + WILL) / 32) + (AGI / 32)
  Mpc = GetMagicPercent
  B = (WIS / 4) + 1  'Black magic multiplier
  W = (WILL / 4) + 1 'white magic multiplier
  
End Sub

Private Function GetMagicPercent() As Integer
  Dim intBODY As Integer
  Dim intHELMET As Integer
  Dim intRING As Integer
  Dim intShield As Integer
  Dim intTOTAL As Integer
  
  If BODYARMOR > 0 Then
    intBODY = armorList(BODYARMOR - 52).Get_MAGPER
  End If
  If RING > 0 Then
    intRING = armorList(RING - 52).Get_MAGPER
  End If
  If HELMET > 0 Then
    intHELMET = armorList(HELMET - 52).Get_MAGPER
  End If
  If RHAND >= 52 Then
    intShield = armorList(RHAND - 52).Get_MAGPER
  End If
  If LHAND >= 52 Then
    intShield = armorList(LHAND - 52).Get_MAGPER
  End If
    
  intTOTAL = ((WIS + WILL) / 8) + intBODY + intRING + intHELMET + intShield
  GetMagicPercent = intTOTAL

End Function

Private Function GetMagicDef() As Integer
  Dim intBODY As Integer
  Dim intHELMET As Integer
  Dim intRING As Integer
  Dim intShield As Integer
  Dim intTOTAL As Integer
  
  If BODYARMOR > 0 Then
    intBODY = armorList(BODYARMOR - 52).Get_MAGICDEF
  End If
  If RING > 0 Then
    intRING = armorList(RING - 52).Get_MAGICDEF
  End If
  If HELMET > 0 Then
    intHELMET = armorList(HELMET - 52).Get_MAGICDEF
  End If
  If RHAND >= 52 Then
    intShield = armorList(RHAND - 52).Get_MAGICDEF
  End If
  If LHAND >= 52 Then
    intShield = armorList(LHAND - 52).Get_MAGICDEF
  End If
    
  intTOTAL = intBODY + intRING + intHELMET + intShield
  GetMagicDef = intTOTAL

End Function

Private Function GetDefensePercent() As Integer
  Dim intBODY As Integer
  Dim intHELMET As Integer
  Dim intRING As Integer
  Dim intShield As Integer
  Dim intTOTAL As Integer
  
  If BODYARMOR > 0 Then
    intBODY = armorList(BODYARMOR - 52).Get_DEFPER
  End If
  If RING > 0 Then
    intRING = armorList(RING - 52).Get_DEFPER
  End If
  If HELMET > 0 Then
    intHELMET = armorList(HELMET - 52).Get_DEFPER
  End If
  If RHAND >= 52 Then
    intShield = armorList(RHAND - 52).Get_DEFPER
  End If
  If LHAND >= 52 Then
    intShield = armorList(LHAND - 52).Get_DEFPER
  End If
    
  intTOTAL = 30 + intBODY + intRING + intHELMET + intShield
  GetDefensePercent = intTOTAL
  
End Function

Private Function GetDefenseMult()
  Dim intShield As Integer
  Dim intTOTAL As Integer
  
  If RHAND >= 52 Then
    intShield = LV / 16
  End If
  If LHAND >= 52 Then
    intShield = LV / 16
  End If
    
  intTOTAL = (AGI / 8) + intShield
  GetDefenseMult = intTOTAL

End Function

Private Function GetDefense()
  Dim intBODY As Integer
  Dim intHELMET As Integer
  Dim intRING As Integer
  Dim intShield As Integer
  Dim intTOTAL As Integer
  
  If BODYARMOR > 0 Then
    intBODY = armorList(BODYARMOR - 52).Get_DEFENSE
  End If
  If RING > 0 Then
    intRING = armorList(RING - 52).Get_DEFENSE
  End If
  If HELMET > 0 Then
    intHELMET = armorList(HELMET - 52).Get_DEFENSE
  End If
  If RHAND >= 52 Then
    intShield = armorList(RHAND - 52).Get_DEFENSE
  End If
  If LHAND >= 52 Then
    intShield = armorList(LHAND - 52).Get_DEFENSE
  End If
    
  intTOTAL = VB + intBODY + intRING + intHELMET + intShield
  GetDefense = intTOTAL
  
End Function

Private Function GetWeaponBonus() As Integer
  Dim intLHAND As Integer
  Dim intRHAND As Integer
  Dim intAVERAGE As Integer
  Dim intTOTAL As Integer

  If RHAND% >= 12 And RHAND% <= 51 Then
    intRHAND = weaponList(RHAND% - 12).Get_ATTPER
  End If
  If LHAND% >= 12 And LHAND% <= 51 Then
    intLHAND = weaponList(LHAND% - 12).Get_ATTPER
  End If
  intAVERAGE = (intRHAND + intLHAND) / 2
  intTOTAL = 50 + (LV / 4) + intAVERAGE
  GetWeaponBonus = intTOTAL

End Function

Private Function GetWeapon() As Integer
  Dim intLHAND As Integer
  Dim intRHAND As Integer
  Dim intTOTAL As Integer
  
  If RHAND% >= 12 And RHAND% <= 51 Then
    If RHAND% >= 44 And RHAND <= 46 Then
      If HND = 0 Then
        intRHAND = weaponList(RHAND% - 12).Get_ATTACK
      Else
        intRHAND = weaponList(RHAND% - 12).Get_ATTACK * 0.8
      End If
    Else
      intRHAND = weaponList(RHAND% - 12).Get_ATTACK
    End If
  End If
  If LHAND% >= 12 And LHAND% <= 51 Then
    If LHAND% >= 44 And LHAND <= 46 Then
      If HND = 0 Then
        intLHAND = weaponList(LHAND% - 12).Get_ATTACK
      Else
        intLHAND = weaponList(LHAND% - 12).Get_ATTACK * 0.8
      End If
    Else
      intLHAND = weaponList(LHAND% - 12).Get_ATTACK
    End If
  End If
  intTOTAL = LB + SB + intLHAND + intRHAND + STR
   
  GetWeapon = intTOTAL
  
End Function

Public Function ATTACK(eDEF As Integer, exD As Integer, eDpc As Integer, eElement As Integer, eRb) As Integer
  Dim Z As Integer
  Dim n As Double
  Dim Diff As Integer
  Dim dM As Integer 'Damage Multiplier
  Diff = ATT - (1 / 3 * ATT)
  Z = ((Rnd * Diff) + Diff) - (Rnd * eDEF)
  n = Z * (((Rnd * xA + 1) + rB) * Apc) - (((Rnd * exD + 1) + eRb) * eDpc)
  If n > 9999 Then n = 9999
  If n < 0 Then n = 0
  ATTACK = n

End Function

Public Sub LoadCharacter(strName As String, intCC As Integer, intCID As Integer, intLV As Integer, dblEXP As Double, dblNLV As Double, intMHP As Integer, intRHP As Integer, intMMP As Integer, intRMP As Integer, intHND As Integer)
  NAME = LTrim$(strName$)
  CC = intCC%
  CID = intCID%
  LV = intLV%
  EXP = dblEXP#
  NLV = dblNLV#
  MHP = intMHP%
  RHP = intRHP%
  MMP = intMMP%
  RMP = intRMP%
  HND = intHND%
End Sub

Public Sub LoadStat(intSTR As Integer, intAGI As Integer, intVIT As Integer, intWIS As Integer, intWILL As Integer)
  STR = intSTR%
  AGI = intAGI%
  VIT = intVIT%
  WIS = intWIS%
  WILL = intWILL%
End Sub

Public Sub LoadEquipment(intLHAND As Integer, intRHAND As Integer, intHELMET As Integer, intBODYARMOR As Integer, intRING As Integer)
  LHAND = intLHAND%
  RHAND = intRHAND%
  HELMET = intHELMET%
  BODYARMOR = intBODYARMOR%
  RING = intRING%
End Sub

Public Sub LoadAction(strACTION As String, strCONDITION As String)
  ACTION = strACTION$
  CONDITION = strCONDITION$
End Sub

Public Sub LoadSkill1(strSKILL As String)
  SKILL1 = strSKILL$
End Sub

Public Sub LoadSkill2(strSKILL As String)
  SKILL2 = strSKILL$
End Sub

Public Function Get_NAME() As String
  Get_NAME = NAME
End Function

Public Function Get_ACTION() As String
  Get_ACTION = ACTION
End Function

Public Function Get_CONDITION() As String
  Get_CONDITION = CONDITION
End Function

Public Function Get_SKILL1() As String
  Get_SKILL1 = SKILL1
End Function

Public Function Get_SKILL2() As String
  Get_SKILL2 = SKILL2
End Function

Public Function Get_LHAND() As Integer
  Get_LHAND = LHAND
End Function

Public Function Get_RHAND() As Integer
  Get_RHAND = RHAND
End Function

Public Function Get_HELMET() As Integer
  Get_HELMET = HELMET
End Function

Public Function Get_BODYARMOR() As Integer
  Get_BODYARMOR = BODYARMOR
End Function

Public Function Get_RING() As Integer
  Get_RING = RING
End Function

Public Function Get_MHP() As Integer
  Get_MHP = MHP
End Function

Public Function Get_RHP() As Integer
  Get_RHP = RHP
End Function

Public Function Get_MMP() As Integer
  Get_MMP = MMP
End Function

Public Function Get_RMP() As Integer
  Get_RMP = RMP
End Function

Public Function Get_LV() As Integer
  Get_LV = LV
End Function

Public Function Get_EXP() As Double
  Get_EXP = EXP
End Function

Public Function Get_NLV() As Double
  Get_NLV = NLV
End Function

Public Sub Get_BASICstat(intSTR As Integer, intAGI As Integer, intVIT As Integer, intWIS As Integer, intWILL As Integer)
  intSTR = STR%
  intAGI = AGI%
  intVIT = VIT%
  intWIS = WIS%
  intWILL = WILL%
End Sub

Public Sub Get_STATS(intATT As Integer, intxA As Integer, intApc As Integer, intDEF As Integer, intxD As Integer, intDpc As Integer, intMD As Integer, intxM As Integer, intMpc As Integer)
  intATT = ATT
  intxA = xA
  intApc = Apc
  intDEF = DEF
  intxD = xD
  intDpc = Dpc
  intMD = MD
  intxM = xM
  intMpc = Mpc
End Sub
