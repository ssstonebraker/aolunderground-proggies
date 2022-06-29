Attribute VB_Name = "click32"

Public Const vbBabyBlue = 16560384
Public Const vbOrange = 33023
Public Const vbNavyBlue = 9896450
Public Const vbDarkGreen = 49152
Public Const vbDarkRed = 192
Public Const vbSilver = 14737632



Public Sub FadeBy2(picbox As Object, firstcolor As Long, secondcolor As Long)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim i&, J&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim firstcolorRed%, firstcolorGreen%, firstcolorBlue%
Dim secondcolorRed%, secondcolorGreen%, secondcolorBlue%
Dim ColorDifRed, ColorDifGreen, ColorDifBlue
SaveScale = picbox.ScaleMode: SaveStyle = picbox.DrawStyle
SaveRedraw = picbox.AutoRedraw: picbox.ScaleMode = 3
firstcolorRed = firstcolor And 255
  firstcolorGreen = (firstcolor And 65280) / 256
    firstcolorBlue = (firstcolor And 16711680) / 65536
secondcolorRed = secondcolor And 255
  secondcolorGreen = (secondcolor And 65280) / 256
    secondcolorBlue = (secondcolor And 16711680) / 65536
      aRed = firstcolorRed
      aGreen = firstcolorGreen
      aBlue = firstcolorBlue
      pixels = picbox.ScaleWidth
    If pixels <= 0 Then Exit Sub
        ColorDifRed = (secondcolorRed - firstcolorRed)
        ColorDifGreen = (secondcolorGreen - firstcolorGreen)
        ColorDifBlue = (secondcolorBlue - firstcolorBlue)
          RedDelta = ColorDifRed / pixels
          GreenDelta = ColorDifGreen / pixels
          BlueDelta = ColorDifBlue / pixels
        picbox.DrawStyle = 5
        picbox.AutoRedraw = True
For Y = 0 To pixels + 1
        aRed = aRed + RedDelta
            If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
            If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
            If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
            If ThisColor > -1 Then
                picbox.Line (Y - 2, -2)-(Y - 2, picbox.Height + 2), ThisColor, BF
            End If
    Next Y
picbox.ScaleMode = SaveScale
picbox.DrawStyle = SaveStyle
picbox.AutoRedraw = SaveRedraw
End Sub





