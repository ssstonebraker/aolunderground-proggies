Attribute VB_Name = "Training"
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Sub Holds Quad-Ball Training Subs and Variables '
'______________________________________________________'

Global InputLoaded As Boolean
Global TopTime As Date
Global OldTime As Date
Global TmpTopTime As String
Global TopTimeName As String
Global TopScore As String
Global TopName As String
Global Exit2Mouse As Boolean

'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' Load The Scores From The Registry '
'___________________________________'
Sub LoadScoreTraining()
 TopScore = GetKeyValue(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopScore")
 If Trim(TopScore) = "" Then
  TopScore = "0"
  RegKeys.UpdateKey HKEY_LOCAL_MACHINE, "software\ArviSehmi\QuadBall\Training", "TopScore", "0"
 End If
 ParentForm.HighestScore.Caption = TopScore
 TopName = GetKeyValue(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopName")
 ParentForm.HighName.Caption = TopName
End Sub

'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' Save The Scores Tp The Registry '
'_________________________________'
Public Sub SaveScoreTraining(Name As String, Score As String)
 Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopScore", Score)
 Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopName", Name)
End Sub

'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' Load The Top Times From The Registry '
'______________________________________'
Public Sub LoadTimeTraining()
 TmpTopTime = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopTime")
 If Trim(TmpTopTime) = "" Then
  TmpTopTime = "00:00:00"
  Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\QuadBall\Training", "TopTime", "00:00:00")
 End If
 ParentForm.HighestTime.Caption = TmpTopTime
 TopTimeName = GetKeyValue(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopTimeName")
 ParentForm.HighTimeName.Caption = TopTimeName
End Sub

'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' Save The Top Time To The Registry '
'___________________________________'
Public Sub SaveTimeTraining(Name As String, tTime As String)
 Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopTime", tTime)
 Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball\training", "TopTimeName", Name)
End Sub

