Attribute VB_Name = "modGravityMain"
Attribute VB_Description = "Gravity--Sample OLE Server"
Option Explicit

#If Win32 Then
    Declare Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Declare Function timeGetTime Lib "mmsystem.dll" () As Long
#End If

Sub Main()
End Sub

