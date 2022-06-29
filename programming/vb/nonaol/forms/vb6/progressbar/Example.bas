Attribute VB_Name = "Example"
' This example.bas is for my Example on how to use a progressbar
' This was made in Visual Basic 6.0 for use with MSCOMCTL.ocx
' You can use pretty much any progressbar.ocx with this, you
' might just have to edit it a little bit.  Pass this to whomever you
' would like, it is freeware (which means it is not to be traded or
' sold).  Please give some credit if you use this in anyway, thanx.
'                                                                                        - Azazel
' Greets to:
' Stubbs, Dim, Beav, Pyschic, And, Mase, Coby, Popz, Spider
' Boogyman, Rizz0, Prone, Burton, Dex0r, humm, Trip
'
'

Public Sub Pause(howlong As Long)
Dim nw As Long
nw = Timer
Do Until Timer - nw >= howlong
DoEvents
Loop
End Sub


Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
' This was taken from MMer40.bas, I think it was written by Mission
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function


