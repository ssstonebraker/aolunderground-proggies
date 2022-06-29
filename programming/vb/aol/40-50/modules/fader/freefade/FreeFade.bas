Attribute VB_Name = "modFormFade"
'---------------------------------------------
' FreeFade Module By S5o
' August, 1999
' S5o@hotmail.com
'---------------------------------------------
' The Fade Functions use the type `FadeColor',
' Be sure to define these values.
'---------------------------------------------
' hFreeFade Fades the form horizontally.
' vFreeFade Fades the form vertically.
'---------------------------------------------

Public Type FadeColor
    Red As Integer
    Blue As Integer
    Green As Integer
End Type

Public Sub hFreeFade(cForm As Form, StartColor As FadeColor, EndColor As FadeColor)
Dim a As Integer
With cForm
    .AutoRedraw = True
    .DrawStyle = 6
    .DrawMode = 13
    .ScaleMode = 3
    .DrawWidth = 2
    .ScaleWidth = 256
    .ScaleHeight = 255
End With

For a = 0 To 255
    If StartColor.Red < EndColor.Red Then StartColor.Red = StartColor.Red + 1
    If StartColor.Red > EndColor.Red Then StartColor.Red = StartColor.Red - 1
    If StartColor.Blue < EndColor.Blue Then StartColor.Blue = StartColor.Blue + 1
    If StartColor.Blue > EndColor.Blue Then StartColor.Blue = StartColor.Blue - 1
    If StartColor.Green < EndColor.Green Then StartColor.Green = StartColor.Green + 1
    If StartColor.Green > EndColor.Green Then StartColor.Green = StartColor.Green - 1
    cForm.Line (a, 0)-(a - 1, cForm.ScaleHeight), RGB(StartColor.Red, StartColor.Green, StartColor.Blue), B
Next a

End Sub

Public Sub vFreeFade(cForm As Form, StartColor As FadeColor, EndColor As FadeColor)
Dim a As Integer
With cForm
    .AutoRedraw = True
    .DrawStyle = 6
    .DrawMode = 13
    .ScaleMode = 3
    .DrawWidth = 2
    .ScaleHeight = 255
    .ScaleWidth = 256
End With

For a = 0 To 255
    If StartColor.Red < EndColor.Red Then StartColor.Red = StartColor.Red + 1
    If StartColor.Red > EndColor.Red Then StartColor.Red = StartColor.Red - 1
    If StartColor.Blue < EndColor.Blue Then StartColor.Blue = StartColor.Blue + 1
    If StartColor.Blue > EndColor.Blue Then StartColor.Blue = StartColor.Blue - 1
    If StartColor.Green < EndColor.Green Then StartColor.Green = StartColor.Green + 1
    If StartColor.Green > EndColor.Green Then StartColor.Green = StartColor.Green - 1
    cForm.Line (0, a)-(cForm.ScaleWidth, a - 1), RGB(StartColor.Red, StartColor.Green, StartColor.Blue), B
Next a

End Sub

