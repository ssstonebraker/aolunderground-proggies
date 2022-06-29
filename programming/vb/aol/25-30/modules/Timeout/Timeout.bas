Attribute VB_Name = "Module1"
Sub timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

