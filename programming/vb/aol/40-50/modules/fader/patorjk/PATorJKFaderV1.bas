Attribute VB_Name = "PATorJKFaderV1"
Option Explicit
'PAT or JK Fader Bas 1.0 for VB5
'By PAT or JK (e-mail:patorjk@aol.com)
'Webpage: http://www.patorjk.com/
'
'note: to be able to use this bas you need to know how to use the RGB function
'and how to use arrays

Sub Side2SideFade(What As Object, Colors, numcolor As Integer)
'This sub will fade an object (like a picture or a form)
On Error Resume Next
If numcolor < 2 Then Exit Sub
Dim i%, i2%, DNum(1 To 3) As Single, TheA As Integer
Dim X As Single, x2 As Double, y%
ReDim TheColors(1 To numcolor, 1 To 3) As Double
TheA = 100
x2 = What.ScaleWidth / (TheA)
y% = What.ScaleHeight
What.Cls
For i = 1 To numcolor
  For i2 = 1 To 3
    TheColors(i, i2) = Colors(i, i2)
  Next
Next

What.BackColor = RGB(TheColors(numcolor, 1), TheColors(numcolor, 2), TheColors(numcolor, 3))

For i = 1 To (numcolor - 1)
  DNum(1) = (TheColors(i + 1, 1) - TheColors(i, 1)) / (TheA / (numcolor - 1))
  DNum(2) = (TheColors(i + 1, 2) - TheColors(i, 2)) / (TheA / (numcolor - 1))
  DNum(3) = (TheColors(i + 1, 3) - TheColors(i, 3)) / (TheA / (numcolor - 1))

  For i2 = 1 To (TheA / (numcolor - 1))
    What.Line (X, 0)-(X + x2, y%), RGB(TheColors(i, 1), TheColors(i, 2), TheColors(i, 3)), BF
    TheColors(i, 1) = TheColors(i, 1) + DNum(1)
    TheColors(i, 2) = TheColors(i, 2) + DNum(2)
    TheColors(i, 3) = TheColors(i, 3) + DNum(3)
    X = X + x2
  Next i2
Next i

' Example on how 2 use: (put this in a button)

'ReDim Colors(1 To 5, 3) As Integer
'Colors(1, 1) = 255 'amount of red in color1
'Colors(1, 2) = 0   'amount of green in color1
'Colors(1, 3) = 0   'amount of blue in color1
'Colors(2, 1) = 0   'amount of red in color2
'Colors(2, 2) = 0   'amount of green in color2
'Colors(2, 3) = 255 'amount of blue in color2
'Colors(3, 1) = 0   '
'Colors(3, 2) = 255 '
'Colors(3, 3) = 0   '
'Colors(4, 1) = 0   '
'Colors(4, 2) = 0   '
'Colors(4, 3) = 255 '
'Colors(5, 1) = 255 '
'Colors(5, 2) = 0   '
'Colors(5, 3) = 0   '
'Call Side2SideFade(Picture1, Colors, 5)

' explain:
'Side2SideFade(Object to fade, Array holding colors, Number of colors to fade)
'the object can be anything that has a line method (like a form or picturebox)
End Sub

Sub Top2BottomFade(What As Object, Colors, numcolor As Integer)
'This sub will fade an object (such as a picture or a form)
On Error Resume Next
If numcolor < 2 Then Exit Sub
Dim i%, i2%, DNum(1 To 3) As Single, TheA As Integer
Dim X As Single, x2 As Double, y%
ReDim TheColors(1 To numcolor, 1 To 3) As Double
TheA = 100
x2 = What.ScaleHeight / (TheA)
y% = What.ScaleWidth
What.Cls
For i = 1 To numcolor
  For i2 = 1 To 3
    TheColors(i, i2) = Colors(i, i2)
  Next
Next

What.BackColor = RGB(TheColors(numcolor, 1), TheColors(numcolor, 2), TheColors(numcolor, 3))

For i = 1 To (numcolor - 1)
  DNum(1) = (TheColors(i + 1, 1) - TheColors(i, 1)) / (TheA / (numcolor - 1))
  DNum(2) = (TheColors(i + 1, 2) - TheColors(i, 2)) / (TheA / (numcolor - 1))
  DNum(3) = (TheColors(i + 1, 3) - TheColors(i, 3)) / (TheA / (numcolor - 1))

  For i2 = 1 To (TheA / (numcolor - 1))
    What.Line (0, X)-(y%, X + x2), RGB(TheColors(i, 1), TheColors(i, 2), TheColors(i, 3)), BF
    TheColors(i, 1) = TheColors(i, 1) + DNum(1)
    TheColors(i, 2) = TheColors(i, 2) + DNum(2)
    TheColors(i, 3) = TheColors(i, 3) + DNum(3)
    X = X + x2
  Next i2
Next i

' Example on how 2 use: (put this in a button)

'ReDim Colors(1 To 12, 3) As Integer
'Colors(1, 1) = 255 'amount of red in color1
'Colors(1, 2) = 0   'amount of green in color1
'Colors(1, 3) = 0   'amount of blue in color1
'Colors(2, 1) = 0   'amount of red in color2
'Colors(2, 2) = 0   'amount of green in color2
'Colors(2, 3) = 255 'amount of blue in color2
'Colors(3, 1) = 0   '
'Colors(3, 2) = 255 '
'Colors(3, 3) = 0   '
'Colors(4, 1) = 0   '
'Colors(4, 2) = 0   '
'Colors(4, 3) = 255 '
'Colors(5, 1) = 255 '
'Colors(5, 2) = 0   '
'Colors(5, 3) = 0   '
'Call Top2BottomFade(Picture1, Colors, 5)

' explain:
'Top2BottomFade(Object to fade, Array holding colors, Number of colors to fade)
'the object that is faded can be anything that has a line method (ie forms, pictureboxes)
End Sub

Public Function HtmlFade1(Text1 As String, Colors, numcolor As Integer, Wave As Boolean)
' Use this function for making fades in webpages
' Make Wave equal to true if you want the text
' to also be faded.
If numcolor < 2 Then Exit Function
If numcolor > Len(Text1) Then Exit Function
Dim i%, i2%, DNum(1 To 3) As Double, fadedtext$, whereat%
Dim Waving(1 To 4) As String, Wn%, XMP%
ReDim TheColors(1 To numcolor, 1 To 3) As Double
XMP = Len(Text1)
Waving(1) = "<sub>"
Waving(2) = "</sub>"
Waving(3) = "<sup>"
Waving(4) = "</sup>"
Wn% = 1

For i = 1 To numcolor
  For i2 = 1 To 3
  TheColors(i, i2) = Colors(i, i2)
  Next
Next
whereat% = 1
For i = 1 To (numcolor - 1)
  DNum(1) = (TheColors(i + 1, 1) - TheColors(i, 1)) / (Len(Text1) / (numcolor - 1))
  DNum(2) = (TheColors(i + 1, 2) - TheColors(i, 2)) / (Len(Text1) / (numcolor - 1))
  DNum(3) = (TheColors(i + 1, 3) - TheColors(i, 3)) / (Len(Text1) / (numcolor - 1))

  For i2 = 1 To (Len(Text1) \ (numcolor - 1))
    If Asc(Mid(Text1, whereat%, 1)) = 13 Then
      fadedtext$ = fadedtext$ + "<br>"
    ElseIf Asc(Mid(Text1, whereat%, 1)) = 10 Then
    Else
      If Wave = True Then
       fadedtext$ = fadedtext$ + "<font color=" & Chr(34) & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + """>" + Waving(Wn%) + Mid$(Text1, whereat%, 1) + "</font>"
      Else
       fadedtext$ = fadedtext$ + "<font color=" & Chr(34) & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + """>" + Mid$(Text1, whereat%, 1) + "</font>"
      End If
      Wn% = Wn% + 1
      If Wn% > 4 Then Wn% = 1
    End If
      TheColors(i, 1) = TheColors(i, 1) + DNum(1)
      TheColors(i, 2) = TheColors(i, 2) + DNum(2)
      TheColors(i, 3) = TheColors(i, 3) + DNum(3)
      whereat% = whereat% + 1
  Next i2
Next i

Do While whereat% <= XMP
  If Asc(Mid(Text1, whereat%, 1)) = 13 Then
  fadedtext$ = fadedtext$ + "<br>"
  ElseIf Asc(Mid(Text1, whereat%, 1)) = 10 Then
  Else
    If Wave = True Then
      Wn% = Wn% + 1
      If Wn% > 4 Then Wn% = 1
      fadedtext$ = fadedtext$ + "<font color=" & Chr(34) & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + """>" + Waving(Wn%) + Mid$(Text1, whereat%, 1) + "</font>"
    Else
      fadedtext$ = fadedtext$ + "<font color=" & Chr(34) & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + """>" + Mid$(Text1, whereat%, 1) + "</font>"
    End If
  End If
  whereat% = whereat% + 1
Loop
'End If
HtmlFade1 = fadedtext$
'Example on how to use: (put in button)

'ReDim Colors(1 To 2, 3) As Integer
'Colors(1, 1) = 255 'amount of red in color1
'Colors(1, 2) = 0   'amount of green in color1
'Colors(1, 3) = 0   'amount of blue in color1
'Colors(2, 1) = 0   'amount of red in color2
'Colors(2, 2) = 0   'amount of green in color2
'Colors(2, 3) = 255 'amount of blue in color2
'Text1.text = HtmlFade1(Text1, Colors, 2, True)

'explain:
'HtmlFade1(Text to fade, array holding colors, number of colors to fade, will text wave?)
End Function

Public Function HtmlFade2(Text1 As String, Colors, numcolor As Integer, Wave As Boolean)
' Use this function for making fades for aol chat
' Make Wave equal 2 true if you want the text
' 2 also be faded.
If numcolor < 2 Then Exit Function
If numcolor > Len(Text1) Then Exit Function
Dim i%, i2%, DNum(1 To 3) As Double, fadedtext$, whereat%
Dim Waving(1 To 4) As String, Wn%, XMP%
ReDim TheColors(1 To numcolor, 1 To 3) As Double
XMP = Len(Text1)
Waving(1) = "<sub>"
Waving(2) = "</sub>"
Waving(3) = "<sup>"
Waving(4) = "</sup>"
Wn% = 1

For i = 1 To numcolor
  For i2 = 1 To 3
  TheColors(i, i2) = Colors(i, i2)
  Next
Next
whereat% = 1
For i = 1 To (numcolor - 1)
  DNum(1) = (TheColors(i + 1, 1) - TheColors(i, 1)) / (Len(Text1) / (numcolor - 1))
  DNum(2) = (TheColors(i + 1, 2) - TheColors(i, 2)) / (Len(Text1) / (numcolor - 1))
  DNum(3) = (TheColors(i + 1, 3) - TheColors(i, 3)) / (Len(Text1) / (numcolor - 1))

  For i2 = 1 To (Len(Text1) \ (numcolor - 1))
    If Asc(Mid(Text1, whereat%, 1)) = 13 Then
      fadedtext$ = fadedtext$ + "<br>"
    ElseIf Asc(Mid(Text1, whereat%, 1)) = 10 Then
    Else
      If Wave = True Then
       fadedtext$ = fadedtext$ + "<font color=" & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + ">" + Waving(Wn%) + Mid$(Text1, whereat%, 1)
      Else
       fadedtext$ = fadedtext$ + "<font color=" & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + ">" + Mid$(Text1, whereat%, 1)
      End If
      Wn% = Wn% + 1
      If Wn% > 4 Then Wn% = 1
    End If
      TheColors(i, 1) = TheColors(i, 1) + DNum(1)
      TheColors(i, 2) = TheColors(i, 2) + DNum(2)
      TheColors(i, 3) = TheColors(i, 3) + DNum(3)
      whereat% = whereat% + 1
  Next i2
Next i

Do While whereat% <= XMP
  If Asc(Mid(Text1, whereat%, 1)) = 13 Then
  fadedtext$ = fadedtext$ + "<br>"
  ElseIf Asc(Mid(Text1, whereat%, 1)) = 10 Then
  Else
    If Wave = True Then
      Wn% = Wn% + 1
      If Wn% > 4 Then Wn% = 1
      fadedtext$ = fadedtext$ + "<font color=" & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + ">" + Waving(Wn%) + Mid$(Text1, whereat%, 1)
    Else
      fadedtext$ = fadedtext$ + "<font color=" & "#" + Cov(Hex(TheColors(i, 1))) + Cov(Hex(TheColors(i, 2))) + Cov(Hex(TheColors(i, 3))) + ">" + Mid$(Text1, whereat%, 1)
    End If
  End If
  whereat% = whereat% + 1
Loop
HtmlFade2 = fadedtext$
'Example on how to use: (put in button)

'ReDim Colors(1 To 2, 3) As Integer
'Colors(1, 1) = 255 'amount of red in color1
'Colors(1, 2) = 0   'amount of green in color1
'Colors(1, 3) = 0   'amount of blue in color1
'Colors(2, 1) = 0   'amount of red in color2
'Colors(2, 2) = 0   'amount of green in color2
'Colors(2, 3) = 255 'amount of blue in color2
'Text1.text = HtmlFade2(Text1, Colors, 2, True)

'explain:
'HtmlFade2(Text to fade, array holding colors, number of colors to fade, will text wave?)
End Function

Public Function Cov(text As String) As String
Do While Len(text) < 2
  text = "0" + text
Loop
Cov = text
End Function

Public Sub HtmlPreview(Text1 As String, Picture As PictureBox, Colors, numcolor As Integer, Wave As Boolean)
'This sub lets you preview what a fade will look
'like
If numcolor < 2 Then Exit Sub
If numcolor > Len(Text1) Then Exit Sub
Dim i%, i2%, DNum(1 To 3) As Double, fadedtext$, whereat%
Dim Wn%, XMP%
ReDim TheColors(1 To numcolor, 1 To 3) As Double
XMP = Len(Text1)
Wn% = 1

For i = 1 To numcolor
  For i2 = 1 To 3
  TheColors(i, i2) = Colors(i, i2)
  Next
Next
whereat% = 1
Picture.Cls
Picture.CurrentX = 0
Picture.CurrentY = 0

For i = 1 To (numcolor - 1)
  DNum(1) = (TheColors(i + 1, 1) - TheColors(i, 1)) / (Len(Text1) / (numcolor - 1))
  DNum(2) = (TheColors(i + 1, 2) - TheColors(i, 2)) / (Len(Text1) / (numcolor - 1))
  DNum(3) = (TheColors(i + 1, 3) - TheColors(i, 3)) / (Len(Text1) / (numcolor - 1))

  For i2 = 1 To (Len(Text1) \ (numcolor - 1))
    If Asc(Mid(Text1, whereat%, 1)) = 13 Then
      Picture.Print
    ElseIf Asc(Mid(Text1, whereat%, 1)) = 10 Then
    Else
      If Wave = True Then
        Wn% = Wn% + 1
        If Wn% > 4 Then Wn% = 1
        Select Case Wn%
        Case 1: Picture.CurrentY = Picture.CurrentY - 15
        Case 2: Picture.CurrentY = Picture.CurrentY + 15
        Case 3: Picture.CurrentY = Picture.CurrentY + 15
        Case 4: Picture.CurrentY = Picture.CurrentY - 15
        End Select
        Picture.ForeColor = RGB(TheColors(i, 1), TheColors(i, 2), TheColors(i, 3))
        Picture.Print Mid$(Text1, whereat%, 1);
      Else
        Picture.ForeColor = RGB(TheColors(i, 1), TheColors(i, 2), TheColors(i, 3))
        Picture.Print Mid$(Text1, whereat%, 1);
      End If
      Wn% = Wn% + 1
      If Wn% > 4 Then Wn% = 1
    End If
      TheColors(i, 1) = TheColors(i, 1) + DNum(1)
      TheColors(i, 2) = TheColors(i, 2) + DNum(2)
      TheColors(i, 3) = TheColors(i, 3) + DNum(3)
      whereat% = whereat% + 1
  Next i2
Next i

Do While whereat% <= XMP
  If Asc(Mid(Text1, whereat%, 1)) = 13 Then
  Picture.Print
  ElseIf Asc(Mid(Text1, whereat%, 1)) = 10 Then
  Else
    If Wave = True Then
      Wn% = Wn% + 1
      If Wn% > 4 Then Wn% = 1
        Select Case Wn%
        Case 1: Picture.CurrentY = Picture.CurrentY - 15
        Case 2: Picture.CurrentY = Picture.CurrentY + 15
        Case 3: Picture.CurrentY = Picture.CurrentY + 15
        Case 4: Picture.CurrentY = Picture.CurrentY - 15
        End Select
        Picture.ForeColor = RGB(TheColors(i, 1), TheColors(i, 2), TheColors(i, 3))
        Picture.Print Mid$(Text1, whereat%, 1);
    Else
        Picture.ForeColor = RGB(TheColors(i, 1), TheColors(i, 2), TheColors(i, 3))
        Picture.Print Mid$(Text1, whereat%, 1);
    End If
  End If
  whereat% = whereat% + 1
Loop
'Example on how to use: (put in button)

'ReDim Colors(1 To 11, 3) As Integer
'Colors(1, 1) = 255 'amount of red in color1
'Colors(1, 2) = 0   'amount of green in color1
'Colors(1, 3) = 0   'amount of blue in color1
'Colors(2, 1) = 0   'amount of red in color2
'Colors(2, 2) = 0   'amount of green in color2
'Colors(2, 3) = 255 'amount of blue in color2
'Call HtmlPreview(Text1, Picture1, Colors, 2, True)

'explain:
'HtmlPreview(Text to fade, Picturebox to preview fade in, Array holding colors, number of colors to fade, text waving?)
End Sub
