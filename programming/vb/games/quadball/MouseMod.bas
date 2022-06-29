Attribute VB_Name = "MouseMod"
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Sub Hold Mouse Functions, Api's and Variables '
'____________________________________________________'
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Global ExitMouse As Boolean
Global ExitMainMouse As Boolean
Global ExitInputWindow As Boolean
Global curXY As POINTAPI
Global resp
Public Function Get_Mouse_X() As Long
    resp = GetCursorPos(curXY)
    Get_Mouse_X = curXY.X
End Function
Public Function Get_Mouse_Y() As Long
    resp = GetCursorPos(curXY)
    Get_Mouse_Y = curXY.Y
End Function
Public Sub Set_Mouse_X(X As Long)
    resp = GetCursorPos(curXY)
    resp = SetCursorPos(X, curXY.Y)
End Sub
Public Sub Set_Mouse_Y(Y As Long)
    resp = GetCursorPos(curXY)
    resp = SetCursorPos(curXY.X, Y)
End Sub
Public Sub Set_Mouse_X_Y(X As Long, Y As Long)
    resp = GetCursorPos(curXY)
    resp = SetCursorPos(curXY.X, Y)
    resp = GetCursorPos(curXY)
    resp = SetCursorPos(X, curXY.Y)
End Sub
