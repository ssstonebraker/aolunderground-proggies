Attribute VB_Name = "Module1"
Sub timeout(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub
Public Sub SndClick()
Playwav ("C:\WINDOWS\MEDIA\Utopia Close.wav")
End Sub
Public Sub FadeLabel3(Label As Label)
Dim I As Integer
For I = 0 To 250 Step 25
Label.ForeColor = RGB(I, I, 255)
timeout (0.1)
Next I

For I = 250 To 0 Step -25
Label.ForeColor = RGB(I, I, 255)
timeout (0.1)
Next I
Label.ForeColor = RGB(0, 0, 0)
End Sub

'Public Sub FadeLabelAnyColor(Label As Label)
'Dim I As Integer
'For I = 0 To 250 Step 25
'Label.ForeColor = RGB(I, I, I)
'timeout (0.1)
'Next I

'For I = 250 To 0 Step -25
'Label.ForeColor = RGB(I, I, I)
'timeout (0.1)
'Next I

'End Sub
