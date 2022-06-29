Attribute VB_Name = "LaunchMod"
Public Sub Pause(PauseTime As Double) ' waits
Dim StartTime As Double
StartTime = Timer
Do While Timer < PauseTime + StartTime
 DoEvents
Loop
End Sub
Public Sub ThisDir() ' changes the dir to this dir
 ChDrive App.Path
 ChDir App.Path
End Sub
