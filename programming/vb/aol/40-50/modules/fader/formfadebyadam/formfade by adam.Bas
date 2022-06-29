Attribute VB_Name = "formfade¹"
' this is a really simple bas to use. i spent about 15 minutes making it
' because someone requested it and i felt like makin it. it has some
' real pretty fades too,hehehe. email me at - aoadam@hotmail.com
'              -adam
'  examples-
'
'
' want to fade your form GreenRed? then in the form_paint put-
' FadeFormGreenRed Me
' that's simple eh?
' ¸.·´¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯`·.¸
' ¦   be sure to check out my full prog enervate y2k with over  '¦
' `·.¸     300 flawless options. www.blo0donline.com/adam  ¸.·´
'     ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Dim fade As Long
Dim Instr As Long
Dim intLooP As Long
Sub FadeFormRedGreen(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 64, 56), B
    Next intLooP
End Sub
Sub FadeFormOrangeGreen(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 125, 34), B
    Next intLooP
End Sub

Sub FadeFormPinkPurple(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 25, 84), B
    Next intLooP
End Sub
Sub FadeFormGreenRed(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 45, 34), B
    Next intLooP
End Sub
Sub FadeFormRedBlack(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 0, 14), B
    Next intLooP
End Sub
Sub FadeFormYellowLightGreen(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 200, 133), B
    Next intLooP
End Sub
Sub FadeFormWhiteLightBlue(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 234, 233), B
    Next intLooP
End Sub
Sub FadeFormPinkBlue(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 55, 133), B
    Next intLooP
End Sub
Sub FadeFormPurpleBlue(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 100, 304), B
    Next intLooP
End Sub
Sub FadeFormPinkLightBlue(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 180, 241), B
    Next intLooP
End Sub
Sub FadeFormPeachGreen(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 199, 199), B
    Next intLooP
End Sub

Sub FadeFormYellowGreen(vForm As Form)
    On Error Resume Next
    Dim intLooP As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLooP = 0 To 255
        vForm.Line (0, intLooP)-(Screen.Width, intLooP - 1), RGB(255 - intLooP, 155, 33), B
    Next intLooP
End Sub
