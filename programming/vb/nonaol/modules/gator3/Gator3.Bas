Attribute VB_Name = "Gator3"
'This Bas was made by Gator (Andrew UFO@aol.com)
'Special Thanks to monk-e-god, DiP, and aDRaMoLEk
'This is version 3.1 of my bas.

'(¯¯¯¯¯¯¯¯¯¯`•GatorSoft•´¯¯¯¯¯¯¯¯¯¯¯)'
'Codes:

'BottomCenter              BottomCenter Me
'BottomLeftCorner:        BottomLeftCorner Me
'BottomRightCorner:      BottomRightCorner Me
'Bring To Top                success = BringWindowToTop(Me.hwnd)
'Center_Form:               Center_Form Me
'CenterOne4th               CenterOne4th Me
'CenterOnTop                CenterOnTop Me
'CenterThree4th             CenterThree4th Me
'ColorBars                     **See Function**
'DisableCtrlAltDel           DisableCtrlAltDel bDisabled:=True (To Enable it: DisableCtrlAltDel bDisabled:=False)
'GetRGB                       GetRGB
'HTMLtoRGB                 HTMLtoRGB
'LeftCenter                    LeftCenter Me
'LeftOne4th                   LeftOne4th Me
'LeftThree4th                 LeftThree4th Me
'LoadList                       Call LoadList(List1,"C:\WINDOWS\filename.ext")
'LoadText                      Call LoadText(Text1,"C:\WINDOWS\filename.ext")
'PlayMidFile                  PlayMidFile("Location")
'ReplaceText                 **See Function**
'RightCenter                  RightCenter Me
'RightOne4th                 RightOne4th Me
'RightThree4th               RightThree4th Me
'SaveText                      Call SaveText(Text1,"C:\WINDOWS\filename.ext")
'SetWinOnTop:              SetWinOnTop = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'Timeout                        Call Timeout(#)
'TopLeftCorner               TopLeftCorner Me
'TopRightCorner             TopRightCorner Me
'(¯¯¯¯¯¯¯¯¯¯`•GatorSoft•´¯¯¯¯¯¯¯¯¯¯¯)'

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function sndPlaySound& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Dim success As Long

Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

Type COLORRGB
Red As Long
Green As Long
Blue As Long
End Type

Sub LoadList(lst As ListBox, File As String)
'Visit DiP's VB World At:
'http://come.to/dipsvbworld
On Error GoTo Error
Open File For Input As #1
Do Until EOF(1)
Input #1, a$
lst.AddItem a$
Loop
Close 1
Exit Sub

Error:
X = MsgBox(Error)
'Example: Call LoadList(List1,"C:\WINDOWS\filename.ext")
'Coded by DiP
End Sub

                             

Sub SaveList(lst As ListBox, File As String)
'Visit DiP's VB World At: http://come.to/dipsvbworld
On Error GoTo Error
Open File For Output As #1
For i = 0 To lst.ListCount - 1
a$ = lst.List(i)
Print #1, a$
Next
Close 1
Exit Sub
Error:
X = MsgBox(Error)

End Sub
Sub LoadText(lst As TextBox, File As String)
 'Example: Call LoadText(Text1,"C:\WINDOWS\filename.ext")
 On Error GoTo Error
 Dim mystr As String
 Open File For Input As #1
 Do While Not EOF(1)
 Line Input #1, a$
 texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
 Loop
 lst = texto$
 Close #1
 Exit Sub
Error:
 X = MsgBox("File Not Found", vbOKOnly, "Error!!")
 End Sub


Sub RightThree4th(Gator As Form)
Gator.Top = (Screen.Height * 0.85) / 1.1 - Gator.Height / 1.1
Gator.Left = Screen.Width - Gator.Width
End Sub

Sub LeftThree4th(Gator As Form)
Gator.Top = (Screen.Height * 0.85) / 1.1 - Gator.Height / 1.1
Gator.Left = 0
End Sub
Sub Timeout(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub
Sub DisableCtrlAltDel(bDisabled As Boolean)
'Disables Control Alt Delete Breaking as well as Ctrl-Escape
Dim beef As Long
beef = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

Function ColorBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
ColorBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of Label1 when you scroll the bars
'Label1.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)
End Function

Function GetRGB(ByVal CVal As Long) As COLORRGB
GetRGB.Blue = Int(CVal / 65536)
GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Function HTMLtoRGB(TheHTML$)
'converts HTML like 0000FF to a RGB value like &HFF0000
If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)

RedX$ = Left(TheHTML$, 2)
GreenX$ = Mid(TheHTML$, 3, 2)
BlueX$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + BlueX$ + GreenX$ + RedX$ + "&"
HTMLtoRGB = Val(rgbhex$)
'this code possible by monk-e-god
End Function
Sub PlayMidFile(ByVal ID$)
'PlayMidFile("Loaction")
'example:
'PlayMidFile ("c:\windows\gator.mid")
  ret = mciSendString("play" & ID$, 0&, 0, 0)
End Sub

Sub Center_Form(Gator As Form)
'To Center The Form place in Form_Load()
'Center_Form Me
Gator.Top = (Screen.Height * 0.85) / 2 - Gator.Height / 2
Gator.Left = Screen.Width / 2 - Gator.Width / 2
End Sub

Sub CenterOne4th(Gator As Form)
Gator.Top = (Screen.Height * 0.85) / 4 - Gator.Height / 4
Gator.Left = Screen.Width / 2 - Gator.Width / 2
End Sub

Sub CenterThree4th(Gator As Form)
Gator.Top = (Screen.Height * 0.85) / 1.1 - Gator.Height / 1.1
Gator.Left = Screen.Width / 2 - Gator.Width / 2
End Sub
Sub RightOne4th(Gator As Form)
Gator.Top = (Screen.Height * 0.85) / 4 - Gator.Height / 4
Gator.Left = Screen.Width - Gator.Width
End Sub

Sub LeftOne4th(Gator As Form)
Gator.Top = (Screen.Height * 0.85) / 4 - Gator.Height / 4
Gator.Left = 0
End Sub
Sub CenterOnTop(Gator As Form)
'To Center The Form at the top of the screen put in Form_Load():
'CenterOnTop Me
Gator.Top = 0
Gator.Left = Screen.Width / 2 - Gator.Width / 2
End Sub

Function ReplaceText(txt As String, find As String, repl As String) As String
'To replace text do this:
'temp$ = text1.text
'temp$ = replacetext(temp$, "a", "â")
'text1.text = temp$
findl = Len(find)
For X = 1 To Len(txt)
a$ = Mid(txt, X, findl)
If a$ = find Then
    Temp$ = Temp$ & repl
     X = X + findl - 1
Else
    Temp$ = Temp$ & Mid(txt, X, 1)
End If
Next X
ReplaceText = Temp$
End Function

Sub TopLeftCorner(Gator As Form)
Gator.Top = 0
Gator.Left = 0
End Sub
Sub TopRightCorner(Gator As Form)
Gator.Top = 0
Gator.Left = Screen.Width - Gator.Width
'(¯¯¯¯¯¯¯¯¯¯`•Gator•´¯¯¯¯¯¯¯¯¯¯)
'  (¯¯¯¯¯¯¯¯¯`•Gator•´¯¯¯¯¯¯¯¯¯)
'    (¯¯¯¯¯¯¯¯`•Gator•´¯¯¯¯¯¯¯¯)
'      (¯¯¯¯¯¯¯`•Gator•´¯¯¯¯¯¯¯)
'        (¯¯¯¯¯¯`•Gator•´¯¯¯¯¯¯)
'          (¯¯¯¯¯`•Gator•´¯¯¯¯¯)
'            (¯¯¯¯`•Gator•´¯¯¯¯)
'              (¯¯¯`•Gator•´¯¯¯)
'                (¯¯`•Gator•´¯¯)
'                  (¯`•Gator•´¯)
'                    (`•Gator•´)
'can you tell i was bored?? heh
End Sub

Sub SaveText(lst As TextBox, File As String)
 On Error GoTo Error
 Dim mystr As String
 Open File For Output As #1
 Print #1, lst
 Close 1
 Exit Sub
Error:
 X = MsgBox("Error!!", vbOKOnly, "Error!!")
 End Sub
Sub BottomRightCorner(Gator As Form)
Gator.Top = Screen.Height - Gator.Height
Gator.Left = Screen.Width - Gator.Width
End Sub

Sub BottomLeftCorner(Gator As Form)
Gator.Top = Screen.Height - Gator.Height
Gator.Left = 0
End Sub

Sub BottomCenter(Gator As Form)
Gator.Top = Screen.Height - Gator.Height
Gator.Left = Screen.Width / 2 - Gator.Width / 2
End Sub

Sub LeftCenter(Gator As Form)
Gator.Top = Screen.Height / 2 - Gator.Height / 2
Gator.Left = 0
End Sub

Sub RightCenter(Gator As Form)
Gator.Top = Screen.Height / 2 - Gator.Height / 2
Gator.Left = Screen.Width - Gator.Width
End Sub
