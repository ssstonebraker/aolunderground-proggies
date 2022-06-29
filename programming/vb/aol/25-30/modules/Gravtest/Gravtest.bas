Attribute VB_Name = "modGravTestMain"
Option Explicit

#If Win32 Then
    Declare Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Declare Function timeGetTime Lib "mmsystem.dll" () As Long
#End If

