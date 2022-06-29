Attribute VB_Name = "dimefade"
'dime fade
'original code written by hob with code warrior
'modified and converted in vb5 by dime
'thanks hob


'ohh yeah it does work,you need to use the "Arial" font

'example : ChatSend fade("red", "black", "blah blah blah blah")
'needs a bas with a chatsend sub
'and it only works with certain color
'very buggy ver
'=)


Function fade(color1 As String, color2 As String, Text As String)
Dim aa, cc, dd, ee, c, f, g, l As String
Dim d, e, h, i, j, k As Integer
c = Text + "             "
h = Len(c)
d = (h / 6)
If d < 1 Then d = 1
e = 1
f = ""
i = h / d
j = 1
Do
ee = "00"
cc = "00"
dd = "00"
aa = color1
If aa = "red" Then
    If j = 1 Then
        cc = "CC"
    ElseIf j = 2 Then
        cc = "AA"
    ElseIf j = 3 Then
        cc = "88"
    ElseIf j = 4 Then
        cc = "66"
    ElseIf j = 5 Then
        cc = "33"
    ElseIf j = 6 Then
        cc = "00"
    End If
ElseIf aa = "blue" Then
    If j = 1 Then
        ee = "CC"
    ElseIf j = 2 Then
        ee = "AA"
    ElseIf j = 3 Then
        ee = "88"
    ElseIf j = 4 Then
        ee = "66"
    ElseIf j = 5 Then
        ee = "33"
    ElseIf j = 6 Then
        ee = "00"
    End If
ElseIf aa = "green" Then
    If j = 1 Then
        dd = "CC"
    ElseIf j = 2 Then
        dd = "AA"
    ElseIf j = 3 Then
        dd = "88"
    ElseIf j = 4 Then
        dd = "66"
    ElseIf j = 5 Then
        dd = "33"
    ElseIf j = 6 Then
        dd = "00"
    End If
ElseIf aa = "gold" Then
    If j = 1 Then
        cc = "CC"
        dd = "CC"
    ElseIf j = 2 Then
        cc = "AA"
        dd = "AA"
    ElseIf j = 3 Then
        cc = "88"
        dd = "88"
    ElseIf j = 4 Then
        cc = "66"
        dd = "66"
    ElseIf j = 5 Then
        cc = "33"
        dd = "33"
    ElseIf j = 6 Then
        cc = "00"
        dd = "00"
    End If
ElseIf aa = "purple" Then
    If j = 1 Then
        cc = "CC"
        ee = "CC"
    ElseIf j = 2 Then
        cc = "AA"
        ee = "AA"
    ElseIf j = 3 Then
        cc = "88"
        ee = "88"
    ElseIf j = 4 Then
        cc = "66"
        ee = "66"
    ElseIf j = 5 Then
        cc = "33"
        ee = "33"
    ElseIf j = 6 Then
        cc = "00"
        ee = "00"
    End If
ElseIf aa = "green" Then
    If j = 1 Then
        dd = "CC"
    ElseIf j = 2 Then
        dd = "AA"
    ElseIf j = 3 Then
        dd = "88"
    ElseIf j = 4 Then
        dd = "66"
    ElseIf j = 5 Then
        dd = "33"
    ElseIf j = 6 Then
        dd = "00"
    End If
End If
aa = color2
If aa = "red" Then
    If j = 1 Then
        cc = "00"
    ElseIf j = 2 Then
        cc = "33"
    ElseIf j = 3 Then
        cc = "66"
    ElseIf j = 4 Then
        cc = "88"
    ElseIf j = 5 Then
        cc = "AA"
    ElseIf j = 6 Then
        cc = "CC"
    End If
ElseIf aa = "blue" Then
    If j = 1 Then
        ee = "00"
    ElseIf j = 2 Then
        ee = "33"
    ElseIf j = 3 Then
        ee = "66"
    ElseIf j = 4 Then
        ee = "88"
    ElseIf j = 5 Then
        ee = "AA"
    ElseIf j = 6 Then
        ee = "CC"
    End If
ElseIf aa = "green" Then
    If j = 1 Then
        dd = "00"
    ElseIf j = 2 Then
        dd = "33"
    ElseIf j = 3 Then
        dd = "66"
    ElseIf j = 4 Then
        dd = "88"
    ElseIf j = 5 Then
        dd = "AA"
    ElseIf j = 6 Then
        dd = "CC"
    End If
ElseIf aa = "gold" Then
    If j = 1 Then
        cc = "00"
        dd = "00"
    ElseIf j = 2 Then
        cc = "33"
        dd = "33"
    ElseIf j = 3 Then
        cc = "66"
        dd = "66"
    ElseIf j = 4 Then
        cc = "88"
        dd = "88"
    ElseIf j = 5 Then
        cc = "AA"
        dd = "AA"
    ElseIf j = 6 Then
        cc = "CC"
        dd = "CC"
    End If
ElseIf aa = "purple" Then
    If j = 1 Then
        cc = "00"
        ee = "00"
    ElseIf j = 2 Then
        cc = "33"
        ee = "33"
    ElseIf j = 3 Then
        cc = "66"
        ee = "66"
    ElseIf j = 4 Then
        cc = "88"
        ee = "88"
    ElseIf j = 5 Then
        cc = "AA"
        ee = "AA"
    ElseIf j = 6 Then
        cc = "CC"
        ee = "CC"
    End If
ElseIf aa = "green" Then
    If j = 1 Then
        dd = "00"
    ElseIf j = 2 Then
        dd = "33"
    ElseIf j = 3 Then
        dd = "66"
    ElseIf j = 4 Then
        dd = "88"
    ElseIf j = 5 Then
        dd = "AA"
    ElseIf j = 6 Then
        dd = "CC"
    End If
End If
    g = cc & dd & ee
    k = (e + d - 1)
    If j = 6 Then k = h
    
    l = Left(c, k)
   
   f = f + ("<font color=""#" & g & """>") + Mid(l, e)
    e = (e + d)
    j = j + 1
Loop Until j = 6
fade = f
End Function

