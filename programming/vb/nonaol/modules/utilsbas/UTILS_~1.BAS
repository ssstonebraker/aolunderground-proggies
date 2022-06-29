Attribute VB_Name = "Utils¹"
' Bas By Naïve
' Version: 1.37
'
' This Bas Is Full Of Stuff To Help Yah
' Out, Utility Stuff...
'
' Credit has been given where due
' Otherwise you better give me credit if ya
' bite my codes
'
' Peace Out,
' -= Naïve =-



Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Sub Bas_Contact()
' If you want to get ahold of me, my email is:
' Naive99@hotmail.com

' Later ,
' Naïve

End Sub

Sub Bas_Greetz()
'Greetz To:
'
' JaGGeD, TaRBaS, Dragon, SouL, K³, Mystical
' BMXBee, Nusance, Kraze, Step, SeeN, RaouL, OtisP,
' Druso, WolfMagicX

'
' Thanks To Anyone Who Posts or Uses My Files!
'

End Sub

Sub Form_HideAllControls()

Dim Crl As Control
For Each Crl In Form1.Controls
    Crl.Visible = False
Next Crl
End Sub



Sub Form_ShowAllControls()

Dim Crl As Control
For Each Crl In Form1.Controls
    Crl.Visible = True
Next Crl
End Sub



Sub Help_CheckBoxesInListBoxes()
' If you want your listbox to have a
' check box, set the style property to 1

End Sub

Sub Keys_FlashLocks()
'this is dumb

Dim x As Integer
x = 20
Do While x > 0
SendKeys "{CAPSLOCK}"
SendKeys "{SCROLLLOCK}"
SendKeys "{NUMLOCK}"
x = x - 1
Loop
End Sub

Sub List_Save(Dir, LST As ListBox)
Dim savelist As Long
On Error Resume Next
Open Dir For Output As #1
For savelist& = 0 To LST.ListCount - 1
Print #1, LST.List(savelist&)
Next savelist&
Close #1
End Sub
Function List_IsInList(SearchFer As String, LST As ListBox) As Boolean
For x = 1 To LST.ListCount - 1
    If LCase(LST.List(0)) = LCase(SearchFer) Then
    IsInList = True
    Exit Function
    If LCase(LST.List(x)) = LCase(SearchFer) Then
    IsInList = True
    Exit Function
    End If
    End If
Next x
IsInList = False
End Function


Sub List_Load(Dir, LST As ListBox)
Dim strin As String
On Error Resume Next
Open Dir For Input As #1
While Not EOF(1)
Input #1, strin$
DoEvents
LST.AddItem strin$
Wend
Close #1
Exit Sub
End Sub

'
'
Function Filter(text$, charfind$, charchange$)
If InStr(text$, charfind$) = 0 Then
ReplaceText = text$
Exit Function
End If

For ReplaceThis = 1 To Len(text$)
TheChar$ = Mid(text$, ReplaceThis, 1)
thechars$ = thechars$ & TheChar$

If TheChar$ = charfind$ Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange$
End If
Next ReplaceThis

Filter = thechars$

End Function




r
Sub Form_FadeToBlack(Frm As Form, Optional ShouldExit As Boolean)
Frm.BackColor = 16777215
TimeOutX (0.02)
Frm.BackColor = 14737632
TimeOutX (0.02)
Frm.BackColor = 12632256
TimeOutX (0.02)
Frm.BackColor = 8421504
TimeOutX (0.02)
Frm.BackColor = 4210752
TimeOutX (0.02)
Frm.BackColor = 0
If ShouldExit = True Then Unload Frm
End Sub

Sub Form_Flash(Frm As Form, Optional NumFlashes As Integer)
On Error Resume Next
Dim x As Integer
Dim Y As String
Y = Frm.BackColor
If Y = 0 Then Y = &H0&

x = NumFlashes
If x = 0 Then x = 100

Do While x > 0
Frm.BackColor = &HFFFFFF
TimeOutX (0.01)
Frm.BackColor = &H0&
TimeOutX (0.01)
x = x - 1
Loop
Frm.BackColor = Y
End Sub

Sub Form_GloryHoleFade(Frm As Form)
' This sub is an adaptaion from another
' sub. There was no name, so I can't
' give kredit

Frm.AutoRedraw = True
Frm.Cls
Dim Cx, Cy, i
    Frm.ScaleMode = 3
    Cx = Frm.ScaleWidth \ 2
    Cy = Frm.ScaleHeight \ 2

Frm.DrawWidth = 2
For i = 0 To 255
Frm.Circle (Cx, Cy), i, RGB(255 - i, 0, 0)  'Red to Black
Next i


End Sub

Sub Form_ExitTopLeft(Frm As Form)

Dim OldH As Integer
On Error Resume Next
Do While Frm.Top + Frm.Height > 0
With Frm
    .Top = .Top - 50
    .Left = .Left - 50
End With
DoEvents
Loop
Unload Frm
End Sub
Sub Form_ExitTopRight(Frm As Form)

Dim OldH As Integer
On Error Resume Next
Do While Frm.Top + Frm.Height > 0
With Frm
    .Top = .Top - 50
    .Left = .Left + 50
End With
DoEvents
Loop
Unload Frm
End Sub

Sub Form_ExitBottomLeft(Frm As Form)


Dim OldH As Integer
On Error Resume Next
Do While Frm.Top < Screen.Height
With Frm
    .Top = .Top + 50
    .Left = .Left - 50
End With
DoEvents
Loop

Unload Frm
End Sub

Sub Form_ExitBottomRight(Frm As Form)

Dim OldH As Integer
On Error Resume Next
Do While Frm.Top < Screen.Height
With Frm
    .Top = .Top + 50
    .Left = .Left + 50
End With
DoEvents
Loop

Unload Frm
End Sub


Sub Form_ExitShrinkTopLeft(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Height = .Height - 50
    .Width = .Width - 50
    End With
TimeOutX (0.01)
Loop
Unload Frm
End Sub
Sub Form_ExitShrinkTopRight(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Height = .Height - 50
    .Width = .Width - 50
    .Left = .Left + 50
End With

Loop
Unload Frm
End Sub

Sub Form_ExitShrinkBottomRight(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Top = .Top + 50
    .Height = .Height - 50
    .Width = .Width - 50
    .Left = .Left + 50
End With
TimeOutX (0.01)
Loop
Unload Frm
End Sub
Sub Form_ExitShrinkBottomLeft(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Top = .Top + 50
    .Height = .Height - 50
    .Width = .Width - 50
    
End With
TimeOutX (0.01)
Loop
Unload Frm
End Sub




Sub Form_Implode(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Top = .Top + 40
    .Height = .Height - 80
    .Width = .Width - 80
    .Left = .Left + 40
End With
TimeOutX (0.001)

Loop
Unload Frm
End Sub

Sub Form_Explode(Frm As Form)
Form_Center Frm
Dim x, Y As Integer
x = Frm.Height
Y = Frm.Width

With Frm
    .Height = 50
    .Width = 50
End With
Dim OldH As Integer
On Error Resume Next
Do
    OldH = Frm.Height
    If Frm.Width < Y Then
        Frm.Width = Frm.Width + 80
    End If
Frm.Top = Frm.Top - 40
    If Frm.Height < x Then
        Frm.Height = Frm.Height + 80
    End If
Frm.Left = Frm.Left - 40
Form_Center Frm
TimeOutX (0.001)

Loop Until x <= Frm.Height Or Y <= Frm.Width
Frm.Height = x
Frm.Width = Y

End Sub




Sub Form_KrazeFade(Frm)
' This sub is an adaptaion from another
' sub. There was no name, so I can't
' give kredit if you use it or optimize it
' give us credit (Naïve, Crystal Dawn)

Frm.AutoRedraw = True
Frm.Cls
Dim Cx, Cy, f, F1, F2, i
Frm.AutoRedraw = True
Frm.Cls
    Frm.ScaleMode = 3
    Cx = Frm.ScaleWidth / 2
    Cy = Frm.ScaleHeight / 2
DrawWidth = 2
For i = 255 To 0 Step -2
f = i / 255
F1 = 1 - f: F2 = 1 + f
Frm.ForeColor = RGB(i, 0, i)
Frm.Line (Cx * F1, Cy * F1)-(Cx * F2, Cy * F2), , BF
Next i
    Frm.ScaleMode = 3
    Cx = Frm.ScaleWidth / 2
    Cy = Frm.ScaleHeight / 2
    Frm.DrawWidth = 2

For i = 0 To 255
Frm.ForeColor = RGB(i, 0, i)
f = i / 255
F1 = 1 - f: F2 = 1 + f
        
Frm.Line (Cx * F1, Cy)-(Cx, Cy * F1)
Frm.Line -(Cx * F2, Cy)
Frm.Line -(Cx, Cy * F2)
Frm.Line -(Cx * F1, Cy)
Next i
End Sub


Sub Form_SuperWarp(Frm As Form, Optional NumTrips As Integer)
Dim oldbk As String
oldbk = Frm.BackColor

If NumTrips = 0 Then NumTrips = 5
Do While NumTrips > 0
Frm.BackColor = 16711680
TimeOutX (0.1)
Frm.BackColor = 12582912
TimeOutX (0.1)
Frm.BackColor = 8388608
TimeOutX (0.1)
Frm.BackColor = 8388736
TimeOutX (0.1)
Frm.BackColor = 12583104
TimeOutX (0.1)
Frm.BackColor = 12583104
TimeOutX (0.1)
Frm.BackColor = 12583104
TimeOutX (0.1)
Frm.BackColor = 16761087
TimeOutX (0.1)
Frm.BackColor = 16777215
TimeOutX (0.1)
Frm.BackColor = 14737632
TimeOutX (0.1)
Frm.BackColor = 12632256
TimeOutX (0.1)
Frm.BackColor = 8421504
TimeOutX (0.1)
Frm.BackColor = 4210752
TimeOutX (0.1)
Frm.BackColor = 128
TimeOutX (0.1)
Frm.BackColor = 192
TimeOutX (0.1)
Frm.BackColor = 255
TimeOutX (0.1)
Frm.BackColor = 8421631
TimeOutX (0.1)
Frm.BackColor = 12632319
TimeOutX (0.1)
Frm.BackColor = 12640511
TimeOutX (0.1)
Frm.BackColor = 8438015
TimeOutX (0.1)
Frm.BackColor = 3313
TimeOutX (0.1)
Frm.BackColor = 16576
NumTrips = NumTrips - 1
Loop
Frm.BackColor = oldbk
End Sub

Function Integer_Increment(IntX As String)
If TestIfInteger(IntX) = True Then
IntX = IntX + 1
Integer_Increment = IntX
Else
Integer_Increment = IntX
End If
End Function

Sub List_KillDupes2(LST As ListBox)
' This is the old style
' if th other one gives ya probs
' use this, but its a lot slower
On Error GoTo Kraze
OldInd = LST.ListIndex
For x = 0 To LST.ListCount - 1
For Y = 0 To LST.ListCount - 1
If Y < LST.ListCount Then
LST.ListIndex = Y
End If
LuCiD = LST.text
If x < LST.ListCount Then
LST.ListIndex = x
Else
Exit Sub
End If
If Y = x Then GoTo Kraze
If Y < LST.ListCount Then
If LuCiD = LST.text Then LST.RemoveItem (Y)
End If
Kraze:
Next Y
Next x
If OldInd > LST.ListCount Then
LST.ListIndex = OldInd
End If
End Sub
Sub List_KillDupes(LST As ListBox)
On Error Resume Next
For x = 1 To LST.ListCount + 2
    LuCiD = LST.List(x)
    For Y = 1 To LST.ListCount + 2
    Kraze = LST.List(Y)
    If Kraze = LuCiD Then
    If x > Y Then
    LST.RemoveItem Y
    ElseIf x < Y Then
    LST.RemoveItem Y
    End If
    End If
    Next Y
    Next x
LST.ListIndex = LST.ListCount - 1
If LST.ListCount > 0 Then
Do While LST.List(0) = LST.List(LST.ListCount - 1)
LST.RemoveItem LST.ListCount - 1
Loop
End If
End Sub


Function Form_GetActiveControl() As String
' Put this in incase ya don't know

' returns true if the control that called it is active
' else it returns the nam of the active control
Form_GetActiveControl = Screen.ActiveControl
End Function
Function Form_GetActiveForm() As String
' Put this in incase ya don't know
Form_GetActiveForm = Screen.ActiveForm
End Function

Sub Form_MoveForm(Frm As Form)
ReleaseCapture
ret& = SendMessage(Frm.hwnd, &HA1, 2, 0)
End Sub



Function List_ListToSring(LSD As ListBox)
Dim XXX As String
Dim NL As String
NL = Chr(13) + Chr(10)
XXX = ""
For x = 1 To LSD.ListCount
LSD.ListIndex = 0
XXX = XXX & LSD.text & NL
Next x
List_ListToString = XXX
End Function

Function List_ListToSringNumbered(LSD As ListBox)
Dim XXX As String
Dim NL As String
NL = Chr(13) + Chr(10)
XXX = ""
For x = 1 To LSD.ListCount
LSD.ListIndex = 0
XXX = XXX & LSD.ListIndex + 1 & ") " & LSD.text & NL
Next x
List_ListToStringNumbered = XXX
End Function


Sub StatusBar_SetText(StatBar, PanelNumber, newcaption)
    StatBar.Panels(PanelNumber).text = newcaption
End Sub


Sub Fun_99Bottles()
Dim x As Integer
x = 99
Do While x > 0
MsgBox x & " bottles of beer on the wall, " & x & " bottles of beer. Take on down, pass it around, " & x - 1 & " bottles of beer on the wall."
x = x - 1
Loop
End Sub

Sub Fun_Kids()
' never ending. kinda fun
Do
Call MsgBox("It's " & Time & " do you know where your  children are?", vbQuestion, "Warning:")
TimeOutX (0.1)
Loop
End Sub


Function String_Increment(str)
On Error Resume Next
If TestIfInteger(str) = True Then
String_Increment = str + 1
Exit Function
End If
For x = 1 To Len(str)
    Y = Mid$(str, x, 1)
    If TestIfInteger(Y) = True Then
        X2 = InStr(x, str, Y)
            'If TestIfInteger(Mid$(Str, X + 1, 1)) = True Then
            '        W = 0
            '    Do
            '        W = W + 1
            '    Loop While TestIfInteger(Mid$(Str, X + W, 1)) = True
            '        Kraze = Mid$(Str, X + W)
            '            Pop = W - X
            '            MsgBox Kraze
            '            Str = Left$(Str, X - 1) & Kraze & Right$(Str, Len(Str) - Pop)
            '        GoTo Poppy
            'End If
                Woo = Y + 1
        str = Left$(str, X2 - 1) & Woo & Right$(str, Len(str) - X2)
        String_Increment = str
    Else: String_Increment = str
    End If
Poppy:
Next x
End Function
Function String_SpaceWithChar(str As String, TheChar)
' This inserts th char between every charachter
Dim x, Y As Integer
Dim TempString As String
Y = 1
For x = 1 To Len(str)
    TempString = TempString & Mid$(str, Y, 1) & TheChar
    Y = Y + 1
    
Next x
String_SpaceWithChar = TempString
End Function

Function Text_Reverse(str As String)
On Error GoTo Error
Dim words As Integer
For words = Len(text) To 1 Step -1
rt = rt & Mid(text, words, 1)
Next words
Text_Reverse = rt
Exit Function
Error:
Err = 1
End Function

Function TextBox_Filter(txt As TextBox, StringToFind, StringToReplace)
txt.txt = Filter(txt, StringToFind, StringToReplace)
End Function


Sub List_Filter(LstX As ListBox, StringToFind As String)
' This will remove any item without the
' String you pass to it

On Error Resume Next
LstX.ListIndex = 0
For x = 1 To LstX.ListCount
If InStr(1, StringToFind, ListX.text) Then LstX.RemoveItem LstX.ListIndex
LstX.ListIndex = LstX.ListIndex + 1
Next x
End Sub
Sub List_RemoveByFilter(LstX As ListBox, StringToFind As String)
' This will remove any item with the
' String you pass to it

On Error Resume Next
LstX.ListIndex = 0
For x = 1 To LstX.ListCount
If InStr(1, StringToFind, ListX.text) = 0 Then LstX.RemoveItem LstX.ListIndex
LstX.ListIndex = LstX.ListIndex + 1
Next x
End Sub




Sub Form_UnloadAll()
Dim Frm As Form
For Each Frm In Forms
Unload Frm
Next Frm
End Sub

'Coded By DiP
End Sub

Sub Form_HideAll()
Dim Frm As Form
For Each Frm In Forms
Frm.Visible = False
Next Frm

End Sub

Sub Form_ShowAll()
Dim Frm As Form
For Each Frm In Forms
Frm.Visible = True
Next Frm

End Sub


Function HTML_Bold(str As String)
HTML_Bold = "<B>" & str & "</B>"
End Function
Function HTML_Italic(str As String)
HTML_Italic = "<I>" & str & "</I>"
End Function
Function HTML_Underline(str As String)
HTML_Underline = "<U>" & str & "</U>"
End Function

Sub List_AsciiChars(LSD As ListBox)
For x = 33 To 223
LSD.AddItem Chr(x)
Next x
End Sub

Function TestIfInteger(Chk) As Boolean
' Tells you if a string is an integer
' I know this is like the vb function
' IsNumeric, but i wanted to find
' how it was done

Dim Y As Integer
On Error GoTo Nope
Y = Chk / 2
TestIfInteger = True
Exit Function
Nope:
TestIfInteger = False

End Function
Function HTML_Blue(str As String)
HTML_Blue = "<FONT COLOR=""#0000FF"">" & str
End Function

Function HTML_Red(str As String)
HTML_Red = "<FONT COLOR=""#FF0000"">" & str
End Function
Function HTML_Green(str As String)
HTML_Green = "<FONT COLOR=""#008000"">" & str
End Function


Function HTML_Black(str As String)
HTML_Black = "<FONT COLOR=""#FFFFFF"">" & str
End Function

Function HTML_Yellow(str As String)
HTML_Yellow = "<FONT COLOR=""&H0000FFFF&"">" & str
End Function


Function HTML_Purple(str As String)
HTML_Purple = "<FONT COLOR=""&H00C000C0&"">" & str
End Function



Function HTML_White(str As String)
HTML_White = "<FONT COLOR=""&H00FFFFFF&"">" & str
End Function




Sub TextBox_Bold(txt As TextBox)
If txt.Font.Bold = True Then txt.Font.Bold = False
If txt.Font.Bold = False Then txt.Font.Bold = True
End Sub


Sub TextBox_Italic(txt As TextBox)
If txt.Font.Italic = True Then txt.Font.Italic = False
If txt.Font.Italic = False Then txt.Font.Italic = True
End Sub

Sub TextBox_Underline(txt As TextBox)
If txt.Font.Underline = True Then txt.Font.Underline = False
If txt.Font.Underline = False Then txt.Font.Underline = True
End Sub





Sub Form_ExitDown(ThaForm As Form)
Do
ThaForm.Top = Trim(str(Int(ThaForm.Top) + 300))
DoEvents
Loop Until ThaForm.Top > 7200
End Sub



Sub Form_ExitLeft(ThaForm As Form)
Do
ThaForm.Left = Trim(str(Int(ThaForm.Left) - 300))
DoEvents
Loop Until ThaForm.Left < -6300
End Sub



Sub Form_ExitRight(ThaForm As Form)
Do
ThaForm.Left = Trim(str(Int(ThaForm.Left) + 300))
DoEvents
Loop Until ThaForm.Left > 9600
End Sub



Sub Form_ExitUp(ThaForm As Form)
Do
ThaForm.Top = Trim(str(Int(ThaForm.Top) - 300))
DoEvents
Loop Until ThaForm.Top < -4500
End Sub


Sub Label_FlyInOut(lab As Label)
' This makes the Label's Text look
' Like it's flying in and out

lab.Visible = True
lab.FontSize = 1
TimeOutX (0.15)
lab.FontSize = 2
TimeOutX (0.15)
lab.FontSize = 4
TimeOutX (0.15)
lab.FontSize = 6
TimeOutX (0.15)
lab.FontSize = 8
TimeOutX (0.15)
lab.FontSize = 10
TimeOutX (0.15)
lab.FontSize = 12
TimeOutX (0.15)
lab.FontSize = 14
TimeOutX (0.15)
lab.FontSize = 16
TimeOutX (0.15)
lab.FontSize = 18
TimeOutX (0.15)
lab.FontSize = 20
TimeOutX (0.15)
lab.FontSize = 22
TimeOutX (0.15)
lab.FontSize = 24
TimeOutX (0.15)
lab.FontSize = 26
TimeOutX (0.15)
lab.FontSize = 24
TimeOutX (0.15)
lab.FontSize = 22
TimeOutX (0.15)
lab.FontSize = 20
TimeOutX (0.15)
lab.FontSize = 18
TimeOutX (0.15)
lab.FontSize = 16
TimeOutX (0.15)
lab.FontSize = 14
TimeOutX (0.15)
lab.FontSize = 12
TimeOutX (0.15)
lab.FontSize = 10
TimeOutX (0.15)
lab.FontSize = 8
TimeOutX (0.15)
lab.FontSize = 6
TimeOutX (0.15)
lab.FontSize = 4
TimeOutX (0.15)
lab.FontSize = 2
TimeOutX (0.15)
lab.FontSize = 1
lab.Visible = False

End Sub
Sub Label_RedWarp(lbl As Label)
' Fades Label From Red To Black And back
' Call it in a loop to make it go on
' Forever


lbl.ForeColor = 4210752
TimeOutX (0.2)
lbl.ForeColor = 128
TimeOutX (0.2)
lbl.ForeColor = 192
TimeOutX (0.2)
lbl.ForeColor = 255
TimeOutX (0.2)
lbl.ForeColor = 8421631
TimeOutX (0.2)
lbl.ForeColor = 12632319
TimeOutX (0.2)
lbl.ForeColor = 8421631
TimeOutX (0.2)
lbl.ForeColor = 255
TimeOutX (0.2)
lbl.ForeColor = 192
TimeOutX (0.2)
lbl.ForeColor = 128
TimeOutX (0.2)
lbl.ForeColor = 4210752
TimeOutX (0.2)

End Sub

Sub Form_RedWarp(lbl As Form, Optional NumWarps As Integer)
' Fades Label From Red To Black And back
' Call it in a loop to make it go on
' Forever
If NumWarps = 0 Then NumWarps = 5
Do While NumWarps > 0
lbl.BackColor = 4210752
TimeOutX (0.1)
lbl.BackColor = 128
TimeOutX (0.1)
lbl.BackColor = 192
TimeOutX (0.1)
lbl.BackColor = 255
TimeOutX (0.1)
lbl.BackColor = 8421631
TimeOutX (0.1)
lbl.BackColor = 12632319
TimeOutX (0.1)
lbl.BackColor = 8421631
TimeOutX (0.1)
lbl.BackColor = 255
TimeOutX (0.1)
lbl.BackColor = 192
TimeOutX (0.1)
lbl.BackColor = 128
TimeOutX (0.1)
lbl.BackColor = 4210752
TimeOutX (0.1)
NumWarps = NumWarps - 1
Loop

End Sub



Sub TimeOutX(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub


Sub Form_StayOnTop(TheForm As Form)
ret& = SetWindowPos(TheForm.hwnd, -1, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Sub List_AddAsciiStrings(LSD As ListBox)
' Use This To Make An Ascii Shop
' be sure to put Naïve (me) on credits

LSD.AddItem "[`·¸]"
LSD.AddItem "[¸·´]"
LSD.AddItem "(•("
LSD.AddItem ")•)"
LSD.AddItem "(•)"
LSD.AddItem ")•("
LSD.AddItem "`·´`·´"
LSD.AddItem "¸.-~·*'˜¨¯"
LSD.AddItem "¨˜'*·~-.¸"
LSD.AddItem "¸·`¯¨˜'*·~-.¸"
LSD.AddItem "¨ '¹i|¡,¡|i¹'¨"
LSD.AddItem "¨‘¨ˆ‘°ª·--·³°ˆ‘ˆ "
LSD.AddItem "¨ˆ‘°·-.,,.-·³°¨"
LSD.AddItem "·]¦[·"
LSD.AddItem "•,¸,.·´¯`·.,¸,•"
LSD.AddItem "· .¸·´\"
LSD.AddItem "/`·¸. ·"
LSD.AddItem "··¤÷×("
LSD.AddItem ")×÷¤··"
LSD.AddItem "«··÷×)·("
LSD.AddItem ")·(×÷··»"
LSD.AddItem "(¯`·¸ ["
LSD.AddItem "]¸·´¯)"
LSD.AddItem "«·÷·¦["
LSD.AddItem "]¦·÷·»"
LSD.AddItem "•···×"
LSD.AddItem "×···•"
LSD.AddItem "(`¨˜°º²³³™º˜¨"
LSD.AddItem "˜°º™³³²º°˜¨´)"
LSD.AddItem "«÷±·´)"
LSD.AddItem "(`·±÷»"
LSD.AddItem "˜¨°ºðØ"
LSD.AddItem "Øðº°¨˜"
LSD.AddItem "¸.´)(`·•||["
LSD.AddItem "]||•·´)(` .¸"
LSD.AddItem "[`·¸][`·÷<l["
LSD.AddItem "]l>÷·´][¸·´]"
LSD.AddItem "«·÷´'°ºðØ"
LSD.AddItem "Øðº°'`÷·»"
LSD.AddItem "ðº°¨˜°ºð"
LSD.AddItem "ðº°¨˜°ºð"
LSD.AddItem "«·÷´)´)"
LSD.AddItem "(`(`÷·»"
LSD.AddItem "(`·¸_¸.·´)"
LSD.AddItem "(`·¸_¸.·´)"
LSD.AddItem "(¸.·´)"
LSD.AddItem "(¸.·´)"
LSD.AddItem "‹›í¦ì‹›"
LSD.AddItem "•÷±/)"
LSD.AddItem "(\±÷•"
LSD.AddItem "(¯`v"
LSD.AddItem "v´¯)"
LSD.AddItem "•v^÷·•"
LSD.AddItem "•·÷^v•"
LSD.AddItem "(`·~^v÷"
LSD.AddItem "÷v^~·´)"
LSD.AddItem "*·._.•÷"
LSD.AddItem "÷•._.·*"
LSD.AddItem "(¯`·._ (¯`·._"
LSD.AddItem "_.·´¯)_.·´¯)"
LSD.AddItem "(¯`·._"
LSD.AddItem "_.·´¯)"
LSD.AddItem "«­.·´¯`•"
LSD.AddItem "•´¯`·.-»"
LSD.AddItem "‹v÷"
LSD.AddItem "÷v›"
LSD.AddItem "¸.´)(`·["
LSD.AddItem "]·´)(` .¸"
LSD.AddItem "·÷±‡±"
LSD.AddItem "±‡±÷·"
LSD.AddItem "´)·(`"
LSD.AddItem "´)·(`"
LSD.AddItem "¨˜°ºð"
LSD.AddItem "ðº°˜¨"
LSD.AddItem "···¤›"
LSD.AddItem "‹¤···"
LSD.AddItem "«··•[]¤)"
LSD.AddItem "(¤[]•··»"
LSD.AddItem "···÷±¦)"
LSD.AddItem "(¦±÷···"
LSD.AddItem "•···¤»"
LSD.AddItem "«¤···•"
LSD.AddItem "–··•)"
LSD.AddItem "(•··–"
LSD.AddItem "¤···¤(•)"
LSD.AddItem "(•)¤···¤"
LSD.AddItem "···÷±|¤"
LSD.AddItem "¤|±÷···"
LSD.AddItem "()··¤"
LSD.AddItem "¤··()"
LSD.AddItem "€¨'·.,•"
LSD.AddItem "•,.·'¨€"
LSD.AddItem "¤··»"
LSD.AddItem "«··¤"
LSD.AddItem "«¤•···/"
LSD.AddItem "\···•¤»"
End Sub

Sub INI_Write(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub


Function INI_Read(Appname, KeyName As String, FileName As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    INI_Read = Left(sRet, GetPrivateProfileString(Appname, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function


Sub List_CopyListToList(ListA As ListBox, ListB As ListBox)
On Error Resume Next
ListA.ListIndex = 0
ListB.Clear
For x = 0 To ListA.ListCount - 1
    ListB.AddItem ListA
    
    ListA.ListIndex = ListA.ListIndex + 1
Next x
End Sub


Sub List_Fonts(LSD As ListBox)
For x = 1 To Screen.FontCount
    LSD.AddItem Screen.Fonts(x)
Next x
End Sub

Sub List_RemoveMultiple(LSD As ListBox)
On Error Resume Next
For x = 1 To LSD.SelCount
LSD.RemoveItem LSD.ListIndex
Next x
End Sub


Sub List_MakeIndexList(ListFrom As ListBox, ListTo As ListBox)
On Error Resume Next

ListTo.Clear
ListFrom.ListIndex = 0
For x = 0 To ListFrom.ListCount
    ListTo.AddItem ListFrom.ListIndex
    If ListFrom.ListIndex < ListFrom.ListCount + 1 Then
        ListFrom.ListIndex = ListFrom.ListIndex + 1
    End If
Next x
ListTo.RemoveItem ListTo.ListCount - 1
If ListFrom.ListCount <> ListTo.ListCount Then
NL = Chr(13) & Chr(10)
X9 = MsgBox("There are errors in this mail list. It is recommended that you restart the server setup, and relist your mail." & NL & NL & NL & "Restart Server Setup?", vbYesNo, "Server Setup")
If X9 = vbYes Then
Unload Form2
Form2.Show
End If
End If
End Sub
Sub List_SetIndexToString(acid As String, poop As ListBox)
On Error Resume Next
poop.ListIndex = 0
For x = 1 To poop.ListCount
    If poop.text = acid Then
    poop.ListIndex = x - 1
    Exit Sub
    End If
Next x
End Sub


Function TrimSpaces(str As String)
If InStr(1, str, " ") = 0 Then GoTo GoD
Dim vFound As Boolean
Dim x As Integer
Dim Y As String

Do
    x = InStr(1, str, " ")
    Y = Mid$(str, x, 1)
        TrimSpaces = Left$(str, x - 1) & Right$(str, Len(str) - x)
        str = TrimSpaces
Loop Until InStr(1, str, " ") = 0
GoD:
TrimSpaces = str
End Function
Function TrimString(str As String, totrim As String)
If InStr(1, str, totrim) = 0 Then GoTo GoD
Dim vFound As Boolean
Dim x As Integer
Dim Y As String

Do
    x = InStr(1, str, totrim)
    Y = Mid$(str, x, 1)
        TrimString = Left$(str, x - 1) & Right$(str, Len(str) - x)
        str = TrimString
Loop Until InStr(1, str, totrim) = 0
GoD:
TrimString = str
End Function


Sub Help_SendKeys()
' A lotta peeps hate sendkys.
' some times they're good though
'
' Special Keys:
'BACKSPACE:       {BACKSPACE}, {BS}, or {BKSP}
'BREAK:           {BREAK}
'CAPS LOCK:       {CAPSLOCK}
'DEL or DELETE:   {DELETE} or {DEL}
'DOWN ARROW:      {DOWN}
'END:             {END}
'ENTER:           {ENTER}or ~
'ESC:             {ESC}
'HELP:            {HELP}
'HOME:            {HOME}
'INSERT:          {INSERT} or {INS}
'LEFT ARROW:      {LEFT}
'NUM LOCK:        {NUMLOCK}
'PAGE DOWN:       {PGDN}
'PAGE UP:         {PGUP}
'PRINT SCREEN:    {PRTSC}
'RIGHT ARROW:     {RIGHT}
'SCROLL LOCK:     {SCROLLLOCK}
'TAB:             {TAB}
'UP ARROW:        {UP}
'F1:              {F1}
'F2:              {F2}
'F3:              {F3}
'F4:              {F4}
'F5:              {F5}
'F6:              {F6}
'F7:              {F7}
'F8:              {F8}
'F9:              {F9}
'F10:             {F10}
'F11:             {F11}
'F12:             {F12}

'More Special Keys:
'SHIFT:   +
'CTRL:    ^
'ALT:     %
'
' Example: SendKeys"%(m)"
' that does alt + m
' the rest is ez
End Sub
Function Web_DipsVBWorld()
Web_DipsVBWorld = "http://come.to/Dipsvbworld"
End Function


Function Web_KnK()
Web_KnK = "http://www.nwozone.com/knk4o"
End Function



Sub Form_Center(Frm As Form)
A% = (Screen.Width - Frm.Width) / 2
b% = (Screen.Height - Frm.Height) / 2
Frm.Move A%, b%
End Sub


Function Web_Hider()
Web_Hider = "http://www.hider.com"
End Function


