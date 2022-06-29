Attribute VB_Name = "Module1"
Option Explicit
Type POINTAPI 'Declare types
    x As Long
    Y As Long
End Type
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long 'Declare API
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long

Function RandomNumber(finished As Integer)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
