Attribute VB_Name = "Module1"
'Iz representing eXcel 2001
Global Start
'This declares 'start' as a global intiger,

Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
