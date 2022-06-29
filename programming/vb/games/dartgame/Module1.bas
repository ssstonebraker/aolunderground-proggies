Attribute VB_Name = "Module1"
Option Explicit
Global Fleche As Byte
Const srcand = &H8800C6
Const srccopy = &HCC0020
Const SRCERASE = &H440328
Const srcinvert = &H660046

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As _
    Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal _
    hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

                        





Public Sub BitBltTarget_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim success As Integer
    Const picwidth = 20
    Const picheight = 20
    BitBlt Form1.bitblttarget.hDC, x, Y, picwidth, picheight, Form1.BitBltpic(Fleche).hDC, 0, picheight, srcand
    BitBlt Form1.bitblttarget.hDC, x, Y, picwidth, picheight, Form1.BitBltpic(Fleche).hDC, 0, 0, srcinvert


End Sub



Function ASin(x As Double)
Dim pi
    pi = 4 * Atn(1)
    ASin = Atn(x) * 180 / pi
End Function

