Attribute VB_Name = "Premier32"
'Premier Thirty Two Bit Module By Galen Grover
'Representing Trumedia 2000
'Http://Trumedia.Gyrate.org
'The defense filled all the gaps the runner had nowhere to go.

Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const EM_UNDO = &HC7

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_SETHORIZONTALEXTENT = &H194

Public Const SC_SCREENSAVE = &HF140

Public Const SPI_SCREENSAVERRUNNING = 97

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const MAX_PATH = 260

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Function Colors_ExtractGreen_Picturebox(pic As PictureBox)
    
'Extracts green value from a picturebox
'Example: greenvalue& = Colors_ExtractGreen_Picturebox(pic)

Dim blue As Long
Dim green As Long

blue& = pic.BackColor \ 65536
green& = (pic.BackColor - blue& * 65536) \ 256
    
If green& < 0 Then Colors_ExtractGreen_Picturebox = 0: Exit Function
    
    
Colors_ExtractGreen_Picturebox = green&

End Function
Public Function Colors_ExtractBlue_Picturebox(pic As PictureBox)
    
'Extracts blue value from a picturebox
'Example: bluevalue& = Colors_ExtractBlue_Picturebox(pic)

Dim blue As Long
     
blue& = pic.BackColor \ 65536
    
If blue& < 0 Then Colors_ExtractBlue_Picturebox = 0: Exit Function
 
    
Colors_ExtractBlue_Picturebox = blue&

End Function
Function Colors_ExtractRed_Picturebox(pic As PictureBox)
    
'Extracts red value from a picturebox
'Example: redvalue& = Colors_ExtractRed_Picturebox(pic)

Dim blue As Long
Dim green As Long
Dim red As Long

blue& = pic.BackColor \ 65536
green& = (pic.BackColor - blue& * 65536) \ 256
red& = pic.BackColor - blue& * 65536 - green& * 256

If red& < 0 Then Colors_ExtractRed_Picturebox = 0: Exit Function

Colors_ExtractRed_Picturebox = red&

End Function
Public Function Colors_ExtractGreen_Screen()
    
'Extracts green value from a picturebox
'Pic.backcolor = rgb(Colors_ExtractRed_Screen, Colors_ExtractGreen_Screen, Colors_ExtractBlue_Screen)

Dim blue As Long
Dim green As Long
Dim deskhdc As Long
Dim Pxy As POINTAPI
Dim Colors As Long

deskhdc& = GetDC(0)
GetCursorPos Pxy
Colors& = GetPixel(deskhdc&, Pxy.X, Pxy.Y)

blue& = Colors& \ 65536
green& = (Colors& - blue& * 65536) \ 256
    
Colors_ExtractGreen_Screen = green&

End Function
Public Function Colors_ExtractBlue_Screen()
    
'Extracts green value from a picturebox
'Pic.backcolor = rgb(Colors_ExtractRed_Screen, Colors_ExtractGreen_Screen, Colors_ExtractBlue_Screen)

Dim blue As Long
Dim deskhdc As Long
Dim Pxy As POINTAPI
Dim Colors As Long

deskhdc& = GetDC(0)
GetCursorPos Pxy
Colors& = GetPixel(deskhdc&, Pxy.X, Pxy.Y)

blue& = Colors& \ 65536
    
Colors_ExtractBlue_Screen = blue&

End Function
Public Function Colors_ExtractRed_Screen()
    
'Extracts red value from a picturebox
'Pic.backcolor = rgb(Colors_ExtractRed_Screen, Colors_ExtractGreen_Screen, Colors_ExtractBlue_Screen)

Dim red As Long
Dim green As Long
Dim blue As Long
Dim deskhdc As Long
Dim Pxy As POINTAPI
Dim gc As Long
Dim Colors As Long

deskhdc& = GetDC(0)
GetCursorPos Pxy
Colors& = GetPixel(deskhdc&, Pxy.X, Pxy.Y)
     
blue& = Colors& \ 65536
green& = (Colors& - blue& * 65536) \ 256
red& = Colors& - blue& * 65536 - green& * 256

Colors_ExtractRed_Screen = red&

End Function
Public Function Control_Compare(ctl1 As Control, ctl2 As Control, matchcase As Boolean)

'This will return the number of alike items two controls contain

Dim i As Integer
Dim i2 As Integer
Dim comp As String
Dim dacount As Integer

For i% = 0 To ctl1.ListCount - 1
    
    comp$ = ctl1.List(i%)
    
    For i2% = 0 To ctl2.ListCount - 1
            
        If matchcase = True Then
            
            If comp$ = ctl2.List(i2%) Then
            
                dacount% = dacount% + 1: Exit For
                
            End If
             
        Else
            
            If LCase$(comp$) = LCase$(ctl2.List(i2%)) Then
            
                dacount% = dacount% + 1: Exit For
                
            End If
            
        End If
            
    Next i2%
    
Next i%
                
Control_Compare = dacount%

End Function
Public Sub Colors_FadePicturebox2Colors(pic2fade As PictureBox, pic1 As PictureBox, pic2 As PictureBox)

'Used as a fade preview
'Fades a picturebox [pic2fade] from one color [pic1] to a second color [pic2]
'Example: Colors_FadePicturebox2Colors preview, color1, color2

Dim reds As Single, greens As Single, blues As Single
Dim rede As Integer, greene As Integer, bluee As Integer
Dim redc As Single, greenc As Single, bluec As Single
Dim i As Integer

reds! = Colors_ExtractRed_Picturebox(pic1)
greens! = Colors_ExtractGreen_Picturebox(pic1)
blues! = Colors_ExtractBlue_Picturebox(pic1)

rede% = Colors_ExtractRed_Picturebox(pic2)
greene% = Colors_ExtractGreen_Picturebox(pic2)
bluee% = Colors_ExtractBlue_Picturebox(pic2)

redc! = (reds! - rede%) / pic2fade.width
greenc! = (greens! - greene%) / pic2fade.width
bluec! = (blues! - bluee%) / pic2fade.width

For i% = 1 To pic2fade.width

    pic2fade.Line (i%, 0)-(i%, pic2fade.height), RGB(reds!, greens!, blues!)
    reds! = reds! - redc!
    greens! = greens! - greenc!
    blues! = blues! - bluec!
    
Next i%

End Sub
Public Sub Colors_FadePicturebox3Colors(pic2fade As PictureBox, pic1 As PictureBox, pic2 As PictureBox, pic3 As PictureBox)

'Used as a fade preview
'Fades a picturebox [pic2fade] from one color [pic1] to a second color [pic2] to a third color [pic3]
'Example: Colors_FadePicturebox2Colors preview, color1, color2, color3

Dim reds As Single, greens As Single, blues As Single
Dim rede As Single, greene As Single, bluee As Single
Dim rede2 As Single, greene2 As Single, bluee2 As Single
Dim redc As Single, greenc As Single, bluec As Single
Dim redc2 As Single, greenc2 As Single, bluec2 As Single
Dim i As Integer, i2 As Integer

reds! = Colors_ExtractRed_Picturebox(pic1)
greens! = Colors_ExtractGreen_Picturebox(pic1)
blues! = Colors_ExtractBlue_Picturebox(pic1)

rede! = Colors_ExtractRed_Picturebox(pic2)
greene! = Colors_ExtractGreen_Picturebox(pic2)
bluee! = Colors_ExtractBlue_Picturebox(pic2)

rede2! = Colors_ExtractRed_Picturebox(pic3)
greene2! = Colors_ExtractGreen_Picturebox(pic3)
bluee2! = Colors_ExtractBlue_Picturebox(pic3)

redc! = (reds! - rede!) / (pic2fade.width / 2)
greenc! = (greens! - greene!) / (pic2fade.width / 2)
bluec! = (blues! - bluee!) / (pic2fade.width / 2)

redc2! = (rede! - rede2!) / (pic2fade.width / 2)
greenc2! = (greene! - greene2!) / (pic2fade.width / 2)
bluec2! = (bluee! - bluee2!) / (pic2fade.width / 2)

For i% = 1 To pic2fade.width / 2

    pic2fade.Line (i%, 0)-(i%, pic2fade.height), RGB(reds!, greens!, blues!)
    reds! = reds! - redc!
    greens! = greens! - greenc!
    blues! = blues! - bluec!
    
Next i%

For i2% = pic2fade.width / 2 To pic2fade.width

    pic2fade.Line (i2%, 0)-(i2%, pic2fade.height), RGB(rede!, greene!, bluee!)
    rede! = rede! - redc2!
    greene! = greene! - greenc2!
    bluee! = bluee! - bluec2!
    
Next i2%

End Sub
Public Function Colors_FadeString2Colors(string2fade As String, pic1 As PictureBox, pic2 As PictureBox)

'Fades a string [string2fade] from one color [pic1] to a second color [pic2]
'Example:
'        String_Save Colors_FadeString2Colors(text, color1, color2), "c:\windows\desktop\preview.html"
'        Net_WebPage "c:\windows\desktop\preview.html"

'This example will open up the users default web browser and display the string [text]
'faded from pic1.backcolor to pic2.backcolor


Dim red As Single, green As Single, blue As Single
Dim redf As Single, greenf As Single, bluef As Single
Dim redc As Single, greenc As Single, bluec As Single
Dim lngth As Integer
Dim i As Integer
Dim fadedstring As String

red! = Colors_ExtractRed_Picturebox(pic1)
green! = Colors_ExtractGreen_Picturebox(pic1)
blue! = Colors_ExtractBlue_Picturebox(pic1)

redf! = Colors_ExtractRed_Picturebox(pic2)
greenf! = Colors_ExtractGreen_Picturebox(pic2)
bluef! = Colors_ExtractBlue_Picturebox(pic2)

redc! = (red! - redf!) / Len(string2fade$)
greenc! = (green! - greenf!) / Len(string2fade$)
bluec! = (blue! - bluef!) / Len(string2fade$)

For i% = 1 To Len(string2fade$)
    
    fadedstring$ = fadedstring$ & "<font color=" & Colors_FixHex(hex(red!)) & Colors_FixHex(hex(green!)) & Colors_FixHex(hex(blue!)) & ">" & Mid(string2fade, i%, 1)
    red! = red! - redc!
    green! = green! - greenc!
    blue! = blue! - bluec!

Next i%

Colors_FadeString2Colors = fadedstring$

End Function
Public Function Colors_FadeString3Colors(string2fade As String, pic1 As PictureBox, pic2 As PictureBox, pic3 As PictureBox)

'Fades a string [string2fade] from one color [pic1] to a second color [pic2] to a third color [pic3]
'Example:
'        String_Save Colors_FadeString#Colors(text, color1, color2, color3), "c:\windows\desktop\preview.html"
'        Net_WebPage "c:\windows\desktop\preview.html"

'This example will open up the users default web browser and display the string [text]
'faded from pic1.backcolor to pic2.backcolor to pic3.backcolor

Dim red As Single, green As Single, blue As Single
Dim redf As Single, greenf As Single, bluef As Single
Dim redf2 As Single, greenf2 As Single, bluef2 As Single
Dim redc As Single, greenc As Single, bluec As Single
Dim redc2 As Single, greenc2 As Single, bluec2 As Single
Dim lngth As Integer
Dim i As Integer
Dim i2 As Integer
Dim fadedstring As String

red! = Colors_ExtractRed_Picturebox(pic1)
green! = Colors_ExtractGreen_Picturebox(pic1)
blue! = Colors_ExtractBlue_Picturebox(pic1)

redf! = Colors_ExtractRed_Picturebox(pic2)
greenf! = Colors_ExtractGreen_Picturebox(pic2)
bluef! = Colors_ExtractBlue_Picturebox(pic2)

redf2! = Colors_ExtractRed_Picturebox(pic3)
greenf2! = Colors_ExtractGreen_Picturebox(pic3)
bluef2! = Colors_ExtractBlue_Picturebox(pic3)

If InStr(Len(string2fade$) / 2, ".") = True Then

    lngth% = Len(string2fade$) - 1
    
Else

    lngth% = Len(string2fade$)
    
End If

redc! = (red! - redf!) / (lngth% / 2)
greenc! = (green! - greenf!) / (lngth% / 2)
bluec! = (blue! - bluef!) / (lngth% / 2)

redc2! = (redf! - redf2!) / (lngth% / 2)
greenc2! = (greenf! - greenf2!) / (lngth% / 2)
bluec2! = (bluef! - bluef2!) / (lngth% / 2)

For i% = 1 To lngth% / 2
    
    fadedstring$ = fadedstring$ & "<font color=" & Colors_FixHex(hex(red!)) & Colors_FixHex(hex(green!)) & Colors_FixHex(hex(blue!)) & ">" & Mid(string2fade$, i%, 1)
    red! = red! - redc!
    green! = green! - greenc!
    blue! = blue! - bluec!

Next i%

For i2% = lngth% / 2 To lngth%
    
    fadedstring$ = fadedstring$ & "<font color=" & Colors_FixHex(hex(redf!)) & Colors_FixHex(hex(greenf!)) & Colors_FixHex(hex(bluef!)) & ">" & Mid(string2fade$, i2%, 1)
    redf! = redf! - redc2!
    greenf! = greenf! - greenc2!
    bluef! = bluef! - bluec2!

Next i2%

If lngth% < Len(string2fade$) Then fadedstring$ = fadedstring$ & "<font color=" & Colors_FixHex(hex(redf!)) & Colors_FixHex(hex(greenf!)) & Colors_FixHex(hex(bluef!)) & ">" & Right(string2fade$, 1)

Colors_FadeString3Colors = fadedstring$

End Function
Public Function Colors_FixHex(hex)

'Fixes hex color codes by adding a 0 in front values containing only 1 character
'Example:
'        hex(0) & hex(255) & hex(0) = 0ff0 which is not a valid html color
'        However fixhex(hex(0)) & fixhex(hex(255)) & fixhex(hex(0)) = 00ff00

If Len(hex) = 1 Then

    Colors_FixHex = "0" & hex
    
Else

    Colors_FixHex = hex
    
End If

End Function
Public Function Colors_HtmlToRgb(html As Variant)

'Simply converts an html color to an rgb value
'Example: hscroll1.value = HtmlToRgb "ff"
'This would set the value of hscroll1 to 255

If html = "00" Then Colors_HtmlToRgb = 0
If html = "01" Then Colors_HtmlToRgb = 1
If html = "02" Then Colors_HtmlToRgb = 2
If html = "03" Then Colors_HtmlToRgb = 3
If html = "04" Then Colors_HtmlToRgb = 4
If html = "05" Then Colors_HtmlToRgb = 5
If html = "06" Then Colors_HtmlToRgb = 6
If html = "07" Then Colors_HtmlToRgb = 7
If html = "08" Then Colors_HtmlToRgb = 8
If html = "09" Then Colors_HtmlToRgb = 9
If html = "0a" Then Colors_HtmlToRgb = 10
If html = "0b" Then Colors_HtmlToRgb = 11
If html = "0c" Then Colors_HtmlToRgb = 12
If html = "0d" Then Colors_HtmlToRgb = 13
If html = "0e" Then Colors_HtmlToRgb = 14
If html = "0f" Then Colors_HtmlToRgb = 15
If html = "10" Then Colors_HtmlToRgb = 16
If html = "11" Then Colors_HtmlToRgb = 17
If html = "12" Then Colors_HtmlToRgb = 18
If html = "13" Then Colors_HtmlToRgb = 19
If html = "14" Then Colors_HtmlToRgb = 20
If html = "15" Then Colors_HtmlToRgb = 21
If html = "16" Then Colors_HtmlToRgb = 22
If html = "17" Then Colors_HtmlToRgb = 23
If html = "18" Then Colors_HtmlToRgb = 24
If html = "19" Then Colors_HtmlToRgb = 25
If html = "1a" Then Colors_HtmlToRgb = 26
If html = "1b" Then Colors_HtmlToRgb = 27
If html = "1c" Then Colors_HtmlToRgb = 28
If html = "1d" Then Colors_HtmlToRgb = 29
If html = "1e" Then Colors_HtmlToRgb = 30
If html = "1f" Then Colors_HtmlToRgb = 31
If html = "20" Then Colors_HtmlToRgb = 32
If html = "21" Then Colors_HtmlToRgb = 33
If html = "22" Then Colors_HtmlToRgb = 34
If html = "23" Then Colors_HtmlToRgb = 35
If html = "24" Then Colors_HtmlToRgb = 36
If html = "25" Then Colors_HtmlToRgb = 37
If html = "26" Then Colors_HtmlToRgb = 38
If html = "27" Then Colors_HtmlToRgb = 39
If html = "28" Then Colors_HtmlToRgb = 40
If html = "29" Then Colors_HtmlToRgb = 41
If html = "2a" Then Colors_HtmlToRgb = 42
If html = "2b" Then Colors_HtmlToRgb = 43
If html = "2c" Then Colors_HtmlToRgb = 44
If html = "2d" Then Colors_HtmlToRgb = 45
If html = "2e" Then Colors_HtmlToRgb = 46
If html = "2f" Then Colors_HtmlToRgb = 47
If html = "30" Then Colors_HtmlToRgb = 48
If html = "31" Then Colors_HtmlToRgb = 49
If html = "32" Then Colors_HtmlToRgb = 50
If html = "33" Then Colors_HtmlToRgb = 51
If html = "34" Then Colors_HtmlToRgb = 52
If html = "35" Then Colors_HtmlToRgb = 53
If html = "36" Then Colors_HtmlToRgb = 54
If html = "37" Then Colors_HtmlToRgb = 55
If html = "38" Then Colors_HtmlToRgb = 56
If html = "39" Then Colors_HtmlToRgb = 57
If html = "3a" Then Colors_HtmlToRgb = 58
If html = "3b" Then Colors_HtmlToRgb = 59
If html = "3c" Then Colors_HtmlToRgb = 60
If html = "3d" Then Colors_HtmlToRgb = 61
If html = "3e" Then Colors_HtmlToRgb = 62
If html = "3f" Then Colors_HtmlToRgb = 63
If html = "40" Then Colors_HtmlToRgb = 64
If html = "41" Then Colors_HtmlToRgb = 65
If html = "42" Then Colors_HtmlToRgb = 66
If html = "43" Then Colors_HtmlToRgb = 67
If html = "44" Then Colors_HtmlToRgb = 68
If html = "45" Then Colors_HtmlToRgb = 69
If html = "46" Then Colors_HtmlToRgb = 70
If html = "47" Then Colors_HtmlToRgb = 71
If html = "48" Then Colors_HtmlToRgb = 72
If html = "49" Then Colors_HtmlToRgb = 73
If html = "4a" Then Colors_HtmlToRgb = 74
If html = "4b" Then Colors_HtmlToRgb = 75
If html = "4c" Then Colors_HtmlToRgb = 76
If html = "4d" Then Colors_HtmlToRgb = 77
If html = "4e" Then Colors_HtmlToRgb = 78
If html = "4f" Then Colors_HtmlToRgb = 79
If html = "50" Then Colors_HtmlToRgb = 80
If html = "51" Then Colors_HtmlToRgb = 81
If html = "52" Then Colors_HtmlToRgb = 82
If html = "53" Then Colors_HtmlToRgb = 83
If html = "54" Then Colors_HtmlToRgb = 84
If html = "55" Then Colors_HtmlToRgb = 85
If html = "56" Then Colors_HtmlToRgb = 86
If html = "57" Then Colors_HtmlToRgb = 87
If html = "58" Then Colors_HtmlToRgb = 88
If html = "59" Then Colors_HtmlToRgb = 89
If html = "5a" Then Colors_HtmlToRgb = 90
If html = "5b" Then Colors_HtmlToRgb = 91
If html = "5c" Then Colors_HtmlToRgb = 92
If html = "5d" Then Colors_HtmlToRgb = 93
If html = "5e" Then Colors_HtmlToRgb = 94
If html = "5f" Then Colors_HtmlToRgb = 95
If html = "60" Then Colors_HtmlToRgb = 96
If html = "61" Then Colors_HtmlToRgb = 97
If html = "62" Then Colors_HtmlToRgb = 98
If html = "63" Then Colors_HtmlToRgb = 99
If html = "64" Then Colors_HtmlToRgb = 100
If html = "65" Then Colors_HtmlToRgb = 101
If html = "66" Then Colors_HtmlToRgb = 102
If html = "67" Then Colors_HtmlToRgb = 103
If html = "68" Then Colors_HtmlToRgb = 104
If html = "69" Then Colors_HtmlToRgb = 105
If html = "6a" Then Colors_HtmlToRgb = 106
If html = "6b" Then Colors_HtmlToRgb = 107
If html = "6c" Then Colors_HtmlToRgb = 108
If html = "6d" Then Colors_HtmlToRgb = 109
If html = "6e" Then Colors_HtmlToRgb = 110
If html = "6f" Then Colors_HtmlToRgb = 111
If html = "70" Then Colors_HtmlToRgb = 112
If html = "71" Then Colors_HtmlToRgb = 113
If html = "72" Then Colors_HtmlToRgb = 114
If html = "73" Then Colors_HtmlToRgb = 115
If html = "74" Then Colors_HtmlToRgb = 116
If html = "75" Then Colors_HtmlToRgb = 117
If html = "76" Then Colors_HtmlToRgb = 118
If html = "77" Then Colors_HtmlToRgb = 119
If html = "78" Then Colors_HtmlToRgb = 120
If html = "79" Then Colors_HtmlToRgb = 121
If html = "7a" Then Colors_HtmlToRgb = 122
If html = "7b" Then Colors_HtmlToRgb = 123
If html = "7c" Then Colors_HtmlToRgb = 124
If html = "7d" Then Colors_HtmlToRgb = 125
If html = "7e" Then Colors_HtmlToRgb = 126
If html = "7f" Then Colors_HtmlToRgb = 127
If html = "80" Then Colors_HtmlToRgb = 128
If html = "81" Then Colors_HtmlToRgb = 129
If html = "82" Then Colors_HtmlToRgb = 130
If html = "83" Then Colors_HtmlToRgb = 131
If html = "84" Then Colors_HtmlToRgb = 132
If html = "85" Then Colors_HtmlToRgb = 133
If html = "86" Then Colors_HtmlToRgb = 134
If html = "87" Then Colors_HtmlToRgb = 135
If html = "88" Then Colors_HtmlToRgb = 136
If html = "89" Then Colors_HtmlToRgb = 137
If html = "8a" Then Colors_HtmlToRgb = 138
If html = "8b" Then Colors_HtmlToRgb = 139
If html = "8c" Then Colors_HtmlToRgb = 140
If html = "8d" Then Colors_HtmlToRgb = 141
If html = "8e" Then Colors_HtmlToRgb = 142
If html = "8f" Then Colors_HtmlToRgb = 143
If html = "90" Then Colors_HtmlToRgb = 144
If html = "91" Then Colors_HtmlToRgb = 145
If html = "92" Then Colors_HtmlToRgb = 146
If html = "93" Then Colors_HtmlToRgb = 147
If html = "94" Then Colors_HtmlToRgb = 148
If html = "95" Then Colors_HtmlToRgb = 149
If html = "96" Then Colors_HtmlToRgb = 150
If html = "97" Then Colors_HtmlToRgb = 151
If html = "98" Then Colors_HtmlToRgb = 152
If html = "99" Then Colors_HtmlToRgb = 153
If html = "9a" Then Colors_HtmlToRgb = 154
If html = "9b" Then Colors_HtmlToRgb = 155
If html = "9c" Then Colors_HtmlToRgb = 156
If html = "9d" Then Colors_HtmlToRgb = 157
If html = "9e" Then Colors_HtmlToRgb = 158
If html = "9f" Then Colors_HtmlToRgb = 159
If html = "a0" Then Colors_HtmlToRgb = 160
If html = "a1" Then Colors_HtmlToRgb = 161
If html = "a2" Then Colors_HtmlToRgb = 162
If html = "a3" Then Colors_HtmlToRgb = 163
If html = "a4" Then Colors_HtmlToRgb = 164
If html = "a5" Then Colors_HtmlToRgb = 165
If html = "a6" Then Colors_HtmlToRgb = 166
If html = "a7" Then Colors_HtmlToRgb = 167
If html = "a8" Then Colors_HtmlToRgb = 168
If html = "a9" Then Colors_HtmlToRgb = 169
If html = "aa" Then Colors_HtmlToRgb = 170
If html = "ab" Then Colors_HtmlToRgb = 171
If html = "ac" Then Colors_HtmlToRgb = 172
If html = "ad" Then Colors_HtmlToRgb = 173
If html = "ae" Then Colors_HtmlToRgb = 174
If html = "af" Then Colors_HtmlToRgb = 175
If html = "b0" Then Colors_HtmlToRgb = 176
If html = "b1" Then Colors_HtmlToRgb = 177
If html = "b2" Then Colors_HtmlToRgb = 178
If html = "b3" Then Colors_HtmlToRgb = 179
If html = "b4" Then Colors_HtmlToRgb = 180
If html = "b5" Then Colors_HtmlToRgb = 181
If html = "b6" Then Colors_HtmlToRgb = 182
If html = "b7" Then Colors_HtmlToRgb = 183
If html = "b8" Then Colors_HtmlToRgb = 184
If html = "b9" Then Colors_HtmlToRgb = 185
If html = "ba" Then Colors_HtmlToRgb = 186
If html = "bb" Then Colors_HtmlToRgb = 187
If html = "bc" Then Colors_HtmlToRgb = 188
If html = "bd" Then Colors_HtmlToRgb = 189
If html = "be" Then Colors_HtmlToRgb = 190
If html = "bf" Then Colors_HtmlToRgb = 191
If html = "c0" Then Colors_HtmlToRgb = 192
If html = "c1" Then Colors_HtmlToRgb = 193
If html = "c2" Then Colors_HtmlToRgb = 194
If html = "c3" Then Colors_HtmlToRgb = 195
If html = "c4" Then Colors_HtmlToRgb = 196
If html = "c5" Then Colors_HtmlToRgb = 197
If html = "c6" Then Colors_HtmlToRgb = 198
If html = "c7" Then Colors_HtmlToRgb = 199
If html = "c8" Then Colors_HtmlToRgb = 200
If html = "c9" Then Colors_HtmlToRgb = 201
If html = "ca" Then Colors_HtmlToRgb = 202
If html = "cb" Then Colors_HtmlToRgb = 203
If html = "cc" Then Colors_HtmlToRgb = 204
If html = "cd" Then Colors_HtmlToRgb = 205
If html = "ce" Then Colors_HtmlToRgb = 206
If html = "cf" Then Colors_HtmlToRgb = 207
If html = "d0" Then Colors_HtmlToRgb = 208
If html = "d1" Then Colors_HtmlToRgb = 209
If html = "d2" Then Colors_HtmlToRgb = 210
If html = "d3" Then Colors_HtmlToRgb = 211
If html = "d4" Then Colors_HtmlToRgb = 212
If html = "d5" Then Colors_HtmlToRgb = 213
If html = "d6" Then Colors_HtmlToRgb = 214
If html = "d7" Then Colors_HtmlToRgb = 215
If html = "d8" Then Colors_HtmlToRgb = 216
If html = "d9" Then Colors_HtmlToRgb = 217
If html = "da" Then Colors_HtmlToRgb = 218
If html = "db" Then Colors_HtmlToRgb = 219
If html = "dc" Then Colors_HtmlToRgb = 220
If html = "dd" Then Colors_HtmlToRgb = 221
If html = "de" Then Colors_HtmlToRgb = 222
If html = "df" Then Colors_HtmlToRgb = 223
If html = "e0" Then Colors_HtmlToRgb = 224
If html = "e1" Then Colors_HtmlToRgb = 225
If html = "e2" Then Colors_HtmlToRgb = 226
If html = "e3" Then Colors_HtmlToRgb = 227
If html = "e4" Then Colors_HtmlToRgb = 228
If html = "e5" Then Colors_HtmlToRgb = 229
If html = "e6" Then Colors_HtmlToRgb = 230
If html = "e7" Then Colors_HtmlToRgb = 231
If html = "e8" Then Colors_HtmlToRgb = 232
If html = "e9" Then Colors_HtmlToRgb = 233
If html = "ea" Then Colors_HtmlToRgb = 234
If html = "eb" Then Colors_HtmlToRgb = 235
If html = "ec" Then Colors_HtmlToRgb = 236
If html = "ed" Then Colors_HtmlToRgb = 237
If html = "ee" Then Colors_HtmlToRgb = 238
If html = "ef" Then Colors_HtmlToRgb = 239
If html = "f0" Then Colors_HtmlToRgb = 240
If html = "f1" Then Colors_HtmlToRgb = 241
If html = "f2" Then Colors_HtmlToRgb = 242
If html = "f3" Then Colors_HtmlToRgb = 243
If html = "f4" Then Colors_HtmlToRgb = 244
If html = "f5" Then Colors_HtmlToRgb = 245
If html = "f6" Then Colors_HtmlToRgb = 246
If html = "f7" Then Colors_HtmlToRgb = 247
If html = "f8" Then Colors_HtmlToRgb = 248
If html = "f9" Then Colors_HtmlToRgb = 249
If html = "fa" Then Colors_HtmlToRgb = 250
If html = "fb" Then Colors_HtmlToRgb = 251
If html = "fc" Then Colors_HtmlToRgb = 252
If html = "fd" Then Colors_HtmlToRgb = 253
If html = "fe" Then Colors_HtmlToRgb = 254
If html = "ff" Then Colors_HtmlToRgb = 255

End Function
Public Sub Colors_RandomPictureboxColor(pic As PictureBox)

'This sub will generate a random backcolor for a picturebox
'Example: Colors_RandomPictureboxColor(pic1)

pic.BackColor = RGB(Number_Random(255), Number_Random(255), Number_Random(255))

End Sub
Public Sub Control_Sort(Ctl As Control)

'You can't change the sort property of a listbox during runtime
'So I added this sub
'The whole basis behind this sub is that > can be used to alphabetize
'Example: Control_Sort(list1)

Dim i As Integer
Dim X As Integer
Dim temp As String

For i% = 0 To Ctl.ListCount - 2

    For X% = i% + 1 To Ctl.ListCount - 1
    
        If Ctl.List(i%) > Ctl.List(X%) Then
        
            temp$ = Ctl.List(i%)
            Ctl.List(i%) = Ctl.List(X%)
            Ctl.List(X%) = temp
            
        End If
        
    Next X%
    
Next i%

End Sub
Public Sub Control_SizeToForm(Ctl As Control)

'This will resize a control to fit the form
'Put this in the resize procedure of the form
'Example: Control_SizeToForm text1

Ctl.width = ScaleWidth
Ctl.height = ScaleHeight

End Sub
Public Sub Control_Save(Ctl As Control, fullpath As String)

'Saves a control
'Example: Control_Save playlist, "c:\windows\desktop\playlist.txt"

Dim i As Integer
Dim freenumber

File_CheckReadOnly fullpath$, 1

If File_Validity(fullpath$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile

Open fullpath$ For Output As #freenumber 'Open file

    For i% = 0 To Ctl.ListCount 'Go through the whole list
    
        Print #freenumber, Ctl.List(i%) 'Save each line of the list
        
    Next i%
    
Close #freenumber

File_SetNormal fullpath$
    
End Sub
Public Sub Control_Load(Ctl As Control, fullpath As String)

'Loads a control
'Example: Control_Load list1, "c:\windows\desktop\pw.txt"

Dim freenumber
Dim ctlitem As String

If File_Validity(fullpath$, 3) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile

Open fullpath$ For Input As #freenumber

    While Not EOF(freenumber) 'Stop when end of file is reached

        Input #freenumber, ctlitem$ 'Input all lines from text
        DoEvents
        Ctl.AddItem ctlitem$ 'Add item to control
    
    Wend
    
Close #freenumber

If Ctl.List(Ctl.ListCount - 1) = "" Then Ctl.ListIndex = Ctl.ListCount - 1: Ctl.RemoveItem Ctl.ListIndex 'If last item is "" then remove it

End Sub
Public Sub Control_RemoveDuplicates(Ctl As Control, casesensitive As Boolean)

Dim i As Integer
Dim X As Integer
Dim dupe As String


End Sub
Public Sub Control_Transfer(fromctl As Control, toctl As Control, cleartoctl As Boolean, lowcase As Boolean, clearfromctl As Boolean)

'This will copy a control to another control
'List and comboboxes and mixed
'Example: Control_Transfer list1, combo1, true, true, true

Dim i As Integer

If fromctl.ListCount = 0 Then Exit Sub 'If origin control is empty exit sub
If cleartoctl = True Then toctl.Clear 'Clear target control if wanted

For i% = 0 To fromctl.ListCount - 1

    If lowcase = True Then 'Check case option
    
        toctl.AddItem LCase$(fromctl.List(i%)) 'Add lowercase item to target control
        
    Else
    
        toctl.AddItem fromctl.List(i%) 'Add item to target control
        
    End If
    
Next i%

outoffor:
If clearfromctl = True Then fromctl.Clear: Exit Sub 'Clear origin control if wanted

End Sub
Public Sub Control_3d(frm As Form, Ctl As Control)

'This will make all your control have a nice 3d effect
'Make sure your control property is set to flat and fixed single
'Example Control_3d me, list1

frm.ScaleMode = 3
frm.CurrentX = Ctl.left - 1
frm.CurrentY = Ctl.top + Ctl.height
frm.Line -Step(0, -(Ctl.height + 1)), RGB(92, 92, 92)
frm.Line -Step(Ctl.width + 1, 0), RGB(92, 92, 92)
frm.Line -Step(0, Ctl.height + 1), RGB(255, 255, 255)
frm.Line -Step(-(Ctl.width + 1), 0), RGB(255, 255, 255)

End Sub
Public Sub Control_AddSuffix(Ctl As ListBox, Suffix As String)

'Adds a string to the end of each item in a control
'Example: List_AddSuffix snlist, "@aol.com"

Dim X As Integer
Dim i As Integer
Dim ctlcount As Integer

ctlcount% = Ctl.ListCount

For X% = 0 To Ctl.ListCount - 1

    Ctl.AddItem Ctl.List(X%) & Suffix$
    
Next X%

For i% = 1 To ctlcount%

    Ctl.ListIndex = 0
    Ctl.RemoveItem Ctl.ListIndex
    
Next i%

End Sub
Public Sub Control_AddPrefix(Ctl As ListBox, prefix As String)

'Adds a prefix to each item in a control
'Example:
'        dim i as integer
'        for i% = 0 to list1.listcount-1
'        Control_AddPrefix list1, i% & ". "
'        next i%

Dim X As Integer
Dim i As Integer
Dim ctlcount As Integer

ctlcount% = Ctl.ListCount

For X% = 0 To Ctl.ListCount - 1

    Ctl.AddItem prefix & Ctl.List(X%)
    
Next X%

For i% = 1 To ctlcount%

    Ctl.ListIndex = 0
    Ctl.RemoveItem Ctl.ListIndex
    
Next i%

End Sub
Sub Control_AddFiles(Ctl As Control, Path As String, Optional ExtFilter As String)

'Add all the files of the given type in the dir to a control
'Example : List_AddFiles List1, "C:\windows\", "*.dll"

Dim FindFile As String

If ExtFilter = "" Then ExtFilter = "*.*"
If Right(Path, 1) <> "\" Then Path = Path & "\"

FindFile$ = Dir(Path$ & ExtFilter$)

Do Until FindFile$ = ""

    Ctl.AddItem FindFile$
    FindFile$ = Dir
    
Loop

End Sub
Public Sub Control_EqualsSearch(Ctl As Control, SearchFor As String, casesensitive As Boolean)

'Will highlight the first instance of a string in a listbox
'Will set the text of a combo to the first instance of a string
'Example: Control_EqualSearch list1, "eminem", true

Dim i As Integer

For i% = 0 To Ctl.ListCount - 1

    If casesensitive = True Then
    
        If Ctl.List(i%) = SearchFor$ Then Ctl.ListIndex = i%: Exit Sub
        
    Else
    
        If LCase$(Ctl.List(i%)) = LCase$(SearchFor$) Then Ctl.ListIndex = i%: Exit Sub
        
    End If
    
Next i%

MsgBox "Search string not found.", 64, "Search..."

End Sub
Public Sub Control_InstrSearch(Ctl As Control, SearchFor As String, casesensitive As Boolean)

'Will highlight the first item in a listbox containing the string
'Will set the text of a combo to the first item containing the string
'Example: Control_EqualSearch list1, "emi", true
Dim i As Integer

For i% = 0 To Ctl.ListCount - 1

    If casesensitive = True Then
    
        If InStr(Ctl.List(i%), SearchFor$) <> 0 Then Ctl.ListIndex = i%: Exit Sub
        
    Else
    
        If InStr(LCase$(Ctl.List(i%)), LCase$(SearchFor$)) <> 0 Then Ctl.ListIndex = i%: Exit Sub
        
    End If
    
Next i%

MsgBox "Search string not found.", 64, "Search..."

End Sub
Public Sub Control_InstrSearchToList(searchctl As Control, SearchFor As String, ctl2 As Control, matchcase As Boolean)

'Will search a list and add each item containing search string to a second list
'Example: Control_InstrSearchToList list1, "emin", list2

Dim i As Integer

If matchcase = True And InStr(Control_ToNumberedString(searchctl), SearchFor$) = False Then MsgBox "Search string not found.", 64, "Search..."
If matchcase = False And InStr(LCase$(Control_ToNumberedString(searchctl)), LCase$(SearchFor$)) = False Then MsgBox "Search string not found.", 64, "Search..."

For i% = 0 To Ctl.ListCount - 1

    If matchcase = False Then
    
        If InStr(LCase$(searchctl.List(i%)), LCase$(SearchFor$)) <> 0 Then ctl2.AddItem searchctl.List(i%)
    
    Else
    
        If InStr(searchctl.List(i%), SearchFor$) <> 0 Then ctl2.AddItem searchctl.List(i%)

    End If
    
Next i%

End Sub
Public Sub Control_DeleteDuplicates(Ctl As Control)

'Deletes duplicate instances in a listbox

Dim i As Integer
Dim X As Integer
Dim dp As String

For i% = 0 To Ctl.ListCount - 1
    dp$ = Ctl.List(i%)
    For X% = 0 To Ctl.ListCount - 1
        If LCase$(Ctl.List(X%)) Like LCase$(dp$) Then Ctl.ListIndex = X%: Ctl.RemoveItem Ctl.ListIndex
    Next X%
Next i%
    

End Sub
Public Function Control_InstrMatches(Ctl As Control, SearchFor As String, casesensitive As Boolean)

'Will return number of items in listbox contain the search string
'Msgbox "There are " & Control_InstrMatches(list1, "emin", false) " songs containing " & searchfor$ "." ,64, "info..."

Dim i As Integer
Dim count As Integer

For i% = 0 To Ctl.ListCount - 1

    If casesensitive = True Then
    
        If InStr(Ctl.List(i%), SearchFor$) <> 0 Then count% = count% + 1
        
    Else
    
        If InStr(LCase$(Ctl.List(i%)), LCase$(SearchFor$)) <> 0 Then count% = count% + 1
        
    End If
    
Next i%

Control_InstrMatches = count%

End Function
Public Function Control_EqualMatches(Ctl As Control, SearchFor As String, casesensitive As Boolean)

'Will return number of items in listbox are equal to the search string
'Msgbox "There are " & Control_EqualMatches(list1, "emin", false) " songs containing " & searchfor$ "." ,64, "info..."

Dim i As Integer
Dim count As Integer

For i% = 0 To Ctl.ListCount - 1

    If casesensitive = True Then
    
        If Ctl.List(i%) = SearchFor$ Then count% = count% + 1
        
    Else
    
        If LCase$(Ctl.List(i%)) = LCase$(SearchFor$) Then count% = count% + 1
        
    End If
    
Next i%

Control_EqualMatches = count%

End Function
Public Function Control_ToNumberedString(Ctl As Control)

'This will take the items in a control and put them into a numbered string where
'each item is numbered according to its place in the list
'Songlist$ = Control_ToNumberedString(list1)

Dim i As Integer
Dim numstr As String

For i% = 0 To Ctl.ListCount - 1

    numstr$ = numstr$ & i% + 1 & ".] " & Ctl.List(i%) & vbCrLf
    
Next i%

Control_ToNumberedString = numstr$

End Function
Public Sub Control_AddFonts(ctrl As Control)

'Adds system fonts to a control
'Control_AddFonts fontcombo

Dim i As Integer

For i% = 1 To Screen.FontCount 'Go through each font

    ctrl.AddItem Screen.Fonts(i%) 'Add each font to the control
    
Next i%

End Sub
Public Sub Control_AddAsciis(ctrl As Control)

'Adds ascii's to a control
'Example: Control_AddAsciis list1

Dim i As Integer

For i% = 33 To 255 'Go through all the asciis

    ctrl.AddItem Chr$(i%) 'Add each ascii to the control
    
Next i%

End Sub
Public Function CP_CtrlAltDel(enabled As Boolean)

'Enables and disables ctl alt del function
'Examples:
'         CP_CtrlAltDel true - will enable ctl alt del
'         CP_CtrlAltDel false - will disable ctl alt del

Dim lReturn  As Long
Dim lBool As Long

If enabled = False Then lReturn = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, lBool, vbNull)
If enabled = True Then lReturn = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, lBool, vbNull)

End Function
Function CP_DriveEmpty(driveletter As String) As Boolean

'Returns true if drive is empty, false if it is not empty
'Example: If CP_DriveEmpty("a") = True Then exit sub

Dim dirct As String

On Error Resume Next 'Continue because we need the error

dirct = Dir$(driveletter$ & ":\*.*")

If Err.Number = 52 Then '52 is empty drive error

    CP_DriveEmpty = Tru
    
Else

    CP_DriveEmpty = False
    
End If

End Function
Public Sub CP_ShowDesktop()

'Shows the desktop
'Example: CP_ShowDesktop

On Error Resume Next

CP_Run "C:\WINDOWS\SYSTEM\Show Desktop.scf"

End Sub
Public Sub CP_NoBeep()

'This will stop the textboxes from beeping
'Copy and paste the code into the keypress procedure of a textbox

Dim KeyCode

If KeyCode = 13 Then KeyCode = 0

End Sub
Public Sub CP_SaveDialog(frm As Form, fileType As String, fileext As String, defaultext As String, Caption As String, lbl As Label)

'Fileext [file extension] format: "*.html*.txt"
'The label is what holds the path to the file
'Caption is the caption of the open dialog
'Filetype is the types of files the extensions are: "Text Files"
'Defaultext is the extension added to the file if the user omits an extension
'Rest is self explanatory
'Example: CP_SaveDialog me, "All Formats", "*.htm;*.html;*.txt", "Save As...", label1

Dim ofn As OPENFILENAME
Dim A

ofn.lStructSize = Len(ofn)
ofn.hwndOwner = frm.hwnd
ofn.hInstance = App.hInstance
ofn.lpstrFilter = fileType$ + Chr$(0) + fileext$ + Chr$(0)
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = CurDir
ofn.lpstrTitle = Caption$
ofn.FLAGS = 0

A = GetOpenFileName(ofn)

If (A) Then lbl.Caption = Trim$(ofn.lpstrFile)
If InStr(lbl.Caption, ".") = 0 Then lbl.Caption = lbl.Caption & defaultext$

End Sub
Public Function CP_TempPath()

'Returns computer's temp path
'Example: Textbox_Save text1, cp_temppath & "\temp.txt"

Dim strfldr As String
Dim lngrslt As Long

strfldr$ = String(MAX_PATH, 0)
lngrslt& = GetTempPath(MAX_PATH, strfldr)

If lngrslt& <> 0 Then

  CP_TempPath = left(strfldr$, InStr(strfldr$, Chr(0)) - 1)
  
Else

  CP_TempPath = ""
  
End If

End Function
Public Sub CP_StandBy(frm As Form)

'Turns your computer on standby mode
'CP_StandBy me

Call SendMessage(frm.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)

End Sub
Public Sub CP_Run(file As String)

'Opens anything you want
'Example: Run "c:\" will open up the c folder

Dim hwnd

Call ShellExecute(hwnd, "Open", file$, "", App.Path, 1)

End Sub
Public Function CP_SystemPath()

'Returns windows system directory
'Example: File_Cut "c:\windows\desktop\threed32.ocx", CP_SystemPath & "\Threed32.ocx"

Dim strfldr As String
Dim lngrslt As Long

strfldr = String(MAX_PATH, 0)
lngrslt = GetSystemDirectory(strfldr, MAX_PATH)

If lngrslt <> 0 Then

  CP_SystemPath = left(strfldr, InStr(strfldr, Chr(0)) - 1) & "\"
  
Else

  CP_SystemPath = ""
  
End If

End Function
Public Sub Directory_Delete(drctry As String)

'Deletes Directory
'Example: Directory_Delete "C:\Windows\"
'Dont ever do that

If InStr(drctry, ":\") = 0 Then Exit Sub

RmDir drctry$

End Sub
Public Function Directory_Exists(drctry As String) As Boolean



End Function
Public Sub Directory_Create(drctry As String)

'Creates a directory
'Example: Directory_Create "C:\Windows\Desktop\My Program\"

On Error GoTo endit:

If InStr(drctry$, ":\") = 0 Then Exit Sub

MkDir drctry$

endit: Exit Sub

End Sub
Public Function File_CheckReadOnly(fullpath As String, SetNormal)

'This function will see if a file is readonly and if so will set the file to normal if desired
'Example:
'        CP_SaveDialog Me, "All Formats", "*.htm;*.txt", "Save As...", Label1
'        File_CheckReadOnly Label1.caption, True
'        TextBox_Save Text1, Label1.caption
        
Select Case SetNormal
    
    Case 1
    
        If File_GetAttributes(fullpath, 2) = "1" Or _
        File_GetAttributes(fullpath, 2) = "3" Or _
        File_GetAttributes(fullpath, 2) = "5" Or _
        File_GetAttributes(fullpath, 2) = "33" Or _
        File_GetAttributes(fullpath, 2) = "35" Or _
        File_GetAttributes(fullpath, 2) = "37" Then
        
            File_CheckReadOnly = True
            
        Else
        
            File_CheckReadOnly = False
            
        End If
        
        File_SetNormal (fullpath$)
        
    Case 2
        
        If File_GetAttributes(fullpath, 2) = "1" Or _
        File_GetAttributes(fullpath, 2) = "3" Or _
        File_GetAttributes(fullpath, 2) = "5" Or _
        File_GetAttributes(fullpath, 2) = "33" Or _
        File_GetAttributes(fullpath, 2) = "35" Or _
        File_GetAttributes(fullpath, 2) = "37" Then
        
            File_CheckReadOnly = True
        
        Else
        
            File_CheckReadOnly = False
        
        End If
    
End Select

End Function
Public Sub File_Copy(originfilepath As String, newfilepath As String)

'Copies file
'Using the built in sub is just as easy but this will check for file validity first
'Example: File_Copy "C:\My shit\butt sex.jpg", "C:\Homework\History\butt sex.jpg"

If File_Validity(originfilepath$, 1) = False Or File_Validity(newfilepath$, 3) = False Then Exit Sub

FileCopy originfilepath$, newfilepath$

End Sub
Public Sub File_Cut(originfilepath As String, newfilepath As String)

'Copies file and deletes original
'Using the built in sub is just as easy but this will check for file validity first
'Example: File_Copy "C:\My shit\butt sex.jpg", "C:\Homework\History\butt sex.jpg"

If File_Validity(originfilepath$, 1) = False Or File_Validity(newfilepath$, 3) = False Then Exit Sub

FileCopy originfilepath$, newfilepath$
Kill originfilepath$

End Sub
Public Function File_GetAttributes(filefullpath As String, Form)

'Gets the attributes of a file
'Case one will return the name of the attribute
'Case two will return the integer of the attribute
'0 = Normal | 1 = ReadOnly | 2 = Hidden | 4 = System | 32 = Archive

Dim daattr As Integer

Select Case Form

    Case 1 'Return string
    
        If File_Validity(filefullpath$, 3) = False Then Exit Function
        
        daattr% = GetAttr(filefullpath$) 'Get integer
        
        If daattr% = 0 Then File_GetAttributes = "Normal"
        If daattr% = 1 Then File_GetAttributes = "ReadOnly"
        If daattr% = 2 Then File_GetAttributes = "Hidden"
        If daattr% = 3 Then File_GetAttributes = "ReadOnly and System"
        If daattr% = 4 Then File_GetAttributes = "System"
        If daattr% = 5 Then File_GetAttributes = "ReadOnly, Hidden and System"

        If daattr% = 32 Then File_GetAttributes = "Archive"
        If daattr% = 33 Then File_GetAttributes = "Archive and ReadOnly"
        If daattr% = 34 Then File_GetAttributes = "Archive and Hidden"
        If daattr% = 35 Then File_GetAttributes = "Archive and ReadOnly and Hidden"
        If daattr% = 36 Then File_GetAttributes = "Archive and System"
        If daattr% = 37 Then File_GetAttributes = "Archive, ReadOnly, Hidden and System"

    Case 2 'Return integer
        
        If File_Validity(filefullpath$, 3) = False Then Exit Function
        
        File_GetAttributes = GetAttr(filefullpath$)

End Select

End Function
Public Sub File_SetNormal(filefullpath$)

'Sets file attribute to normal

If File_Validity(filefullpath$, 3) = False Then Exit Sub

SetAttr filefullpath$, vbNormal

End Sub
Public Sub File_SetReadOnly(filefullpath$)

'Sets file attribute to read only

If File_Validity(filefullpath$, 3) = False Then Exit Sub

SetAttr filefullpath$, vbReadOnly

End Sub
Public Sub File_SetHidden(filefullpath$)

'Sets file attribute to hidden

If File_Validity(filefullpath$, 3) = False Then Exit Sub

SetAttr filefullpath$, vbHidden

End Sub
Public Sub File_SetArchive(filefullpath$)

'Sets file attribute to archive

If File_Validity(filefullpath$, 3) = False Then Exit Sub

SetAttr filefullpath$, vbArchive

End Sub
Public Function File_GetDirectory(dafile As String)

'This functions returns the directory of a file given the full path

Dim i As Integer
Dim start As Integer

For i% = Len(dafile$) To 1 Step -1

If Mid(dafile$, i%, 1) = "\" Then File_GetDirectory = left(dafile$, i%): Exit Function

Next i%

End Function
Public Sub File_TextAppend(dastring As String, fullpath As String)

'Adds to a text file already saved
'Example: File_TextAppend vbcrlf & "Freedom [RATM]", "C:\my shit\song list.txt"

Dim freenumber
Dim datext As String

File_CheckReadOnly fullpath$, 1

If File_Validity(fullpath$, 1) = False Then Exit Sub 'Check for validity of file

freenumber = FreeFile
datext$ = dastring$

Open fullpath$ For Append As #freenumber 'Open the path to edit it

    Print #freenumber, datext$ 'Add the string to the existing file

Close #freenumber 'Close file

End Sub
Public Function File_Validity(filefullpath As String, thecase) As Boolean

'This will check all aspects of file validity
'Case 1: File existance
'Case 2: File name validity
'Case 3: Both cases
'Example:
'        dim Myfile as string
'        Myfile$ = "C:\windows\desktop\pwlist.txt"
'        If File_Validity(Myfile$, 3) = true then
'            Text_Load Myfile, pwtext.Text
'        Else
'            Msgbox "File does not exist" ,64,"error..."
'        End If

Select Case thecase

    Case 1 'Check existance of file
    
        On Error GoTo done: 'On error goto label "done:"
        FileLen (filefullpath$) 'Gets size of file, error if file doesn't exist, hence on error
        
done:

        File_Validity = False: Exit Function 'File doesn't exist therefore fnction is false
        File_Validity = True
        
    Case 2 'Search filename for necessary and illegal characters
    
        If InStr(filefullpath, ":\") = 0 Then File_Validity = False: Exit Function
        If InStr(Right(filefullpath, 5), ".") = 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "?") <> 0 Then File_Validity = False:  Exit Function
        If InStr(filefullpath, "*") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "<") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, ">") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, Chr(34)) <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "|") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "/") <> 0 Then File_Validity = False: Exit Function
        File_Validity = True

    Case 3 'Perform both cases
    
        On Error GoTo done2:
        
        FileLen (filefullpath$)
        
        'Search filename for necessary and illegal characters
        If InStr(filefullpath, ":\") = 0 Then File_Validity = False: Exit Function
        If InStr(Right(filefullpath, 5), ".") = 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "?") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "*") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "<") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, ">") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, Chr(34)) <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "|") <> 0 Then File_Validity = False: Exit Function
        If InStr(filefullpath, "/") <> 0 Then File_Validity = False: Exit Function
        
        File_Validity = True
        

End Select

File_Validity = True: Exit Function

done2:      File_Validity = False: Exit Function 'From case 3... file doesn't exist

End Function
Public Sub File_Delete(filefullpath As String)

'Deletes file
'Careful with this one
'Example: File_Delete "C:\Windows\my porn\sister in shower.jpg"

If File_Validity(filefullpath$, 1) = False Then Exit Sub 'Check file for existance

Kill filefullpath$ 'Delete file

End Sub
Public Sub File_Rename(filefullpath As String, newfilefullpath As String)

'Renames file
'Using the built in sub is just as easy but this will check for file validity first
'Example: File_Rename "C:\My shit\horse sex.jpg", "C:\My shit\term paper.txt"

If File_Validity(filefullpath$, 1) = False Or File_Validity(newfilefullpath$, 3) = False Then Exit Sub

Name filefullpath$ As newfilefullpath$

End Sub
Public Function File_GetFile(dafile As String)

'This functions returns the directory of a file given the full path

Dim i As Integer
Dim start As Integer

For i% = Len(dafile$) To 1 Step -1
    
    If Mid(dafile$, i%, 1) = "\" Then File_GetFile = Right(dafile$, Len(dafile$) - i%): Exit Function

Next i%

End Function
Public Sub Form_Move(frm As Form)

'Will move a form
'Put this in the mousedown procedure

Call ReleaseCapture
Call SendMessage(frm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
    
End Sub
Public Sub Form_SizePosition(frm As Form, left As Integer, top As Integer, Optional height As Integer, Optional width As Integer)

'Set the position of the form
'Example: Form_Position me, 0, 0

frm.left = left%
frm.top = top%
frm.height = height%
frm.width = width%

End Sub
Public Sub Form_SetTop(frm As Form)

'Form will always be topmost window
'Example: Form_SetTop me

Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)

End Sub
Public Sub Form_SetNotTop(frm As Form)

'Form will no longer be on top
'Example: Form_SetNotTop me

Call SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)

End Sub
Public Sub Form_Center(frm As Form)

'Centers a form
'Usually used in form_load

frm.left = Screen.width / 2 - frm.width / 2
frm.top = Screen.height / 2 - frm.height / 2

End Sub
Public Function Ini_Read(section As String, key As String, pathofini As String) As String

'Gets the info from a ini for a specific item
'Example:
'        If Ini_Read("Options", "intro art", "c:\program files\trumedia designs\trumedia html editor.ini") = false then main.load


Dim buf As String

buf = String(750, Chr(0))
key$ = LCase$(key$)
Ini_Read$ = left(buf$, GetPrivateProfileString(section$, ByVal key$, "", buf$, Len(buf$), pathtoini$))

End Function
Public Sub Ini_Write(section As String, key As String, KeyValue As String, Path As String)

'Writes to an ini file
'Example: Ini_Write "Options", "intro art", "false", "c:\program files\trumedia designs\trumedia html editor.ini"

File_CheckReadOnly Path$, 1
Call WritePrivateProfileString(section$, UCase$(key$), KeyValue$, Path$)

End Sub
Public Sub Label_Spell(dalabel As Label, datext As String, Speed As String)

'This will spell a string to a labels caption
'Speeds:    [time in between in letter placed]
'       1[1.00 seconds]
'       2[0.90 seconds]
'       3[0.80 seconds]
'       4[0.70 seconds]
'       5[0.60 seconds]
'       6[0.50 seconds]
'       7[0.40 seconds]
'       8[0.30 seconds]
'       9[0.20 seconds]
'      10[0.10 seconds]

Dim i As Integer
Dim speeda As Single
Dim spell As String

If Number_Valid(Speed$) = False Then Exit Sub

If Speed = 1 Then speeda! = 1 'define speeds
If Speed = 2 Then speeda! = 0.9
If Speed = 3 Then speeda! = 0.8
If Speed = 4 Then speeda! = 0.7
If Speed = 5 Then speeda! = 0.6
If Speed = 6 Then speeda! = 0.5
If Speed = 7 Then speeda! = 0.4
If Speed = 8 Then speeda! = 0.3
If Speed = 9 Then speeda! = 0.2
If Speed = 10 Then speeda! = 0.1

For i% = 1 To Len(datext$)

    spell$ = Mid(datext$, i%, 1) 'Set variable to letter
    dalabel.Caption = dalabel.Caption & spell$ 'Add letter to textbox
    Program_Pause speeda! 'Timeout to  control thespeed

Next i%

End Sub
Public Sub List_RemoveSelected(lst As ListBox)

'Removes all selected items from a listbox

Dim i As Integer

For i% = lst.ListCount - 1 To 0 Step -1

    If lst.Selected(i%) Then lst.RemoveItem i% 'If list item is selected remove it
    
Next i

End Sub
Public Sub List_HScrollBar(lst As ListBox)

'Gives a listbox horizontal a horizontal scrollbar

Dim DoIt As Long
Dim wID As Integer

wID% = lst.width + 1 'new width in pixels
DoIt& = SendMessage(lst.hwnd, LB_SETHORIZONTALEXTENT, wID%, ByVal 0&)

End Sub
Public Sub List_CopySelectedToControl(lst As ListBox, Ctl As Control)

'Removes all selected items from a listbox

Dim i As Integer

For i% = lst.ListCount - 1 To 0 Step -1
    If lst.Selected(i%) Then Ctl.AddItem lst.List(i%) 'If list item is selected add it to control
    
Next i%

End Sub
Public Sub List_RemoveSelectedToControl(lst As ListBox, Ctl As Control)

'Removes all selected items from a listbox and adds them to a control

Dim i As Integer

For i% = lst.ListCount - 1 To 0 Step -1

    If lst.Selected(i%) Then Ctl.AddItem lst.List(i%): lst.RemoveItem i% 'If list item is selected add it to control and remove it

Next i%

End Sub
Public Sub Menu_RunByNumber(ProgramClassName As String, TopMenu As Long, SubMenu As Long)

Dim classname As Long
Dim menu1 As Long
Dim menu2 As Long
Dim menuid As Long

'Runs any programs menu by a set of numbers
'0 is the first number in the top menu and submenu
'Example:
'        Menu_RunByNumber "Aol Frame25", 0, 0
'This will execute new in the file menu of aol

classname& = FindWindow(ProgramClassName$, vbNullString)
menu1& = GetMenu(classname&)
menu2& = GetSubMenu(menu1&, TopMenu&)
menuid& = GetMenuItemID(menu2&, SubMenu&)

Call SendMessageLong(classname&, WM_COMMAND, menuid&, 0&)

End Sub
Public Sub Net_Email(emailadress As String)

'Sends email to given address
'Example: Net_Email "klowde@netzero.net"

Dim hwnd As Long

ShellExecute hwnd, "open", "mailto:" + emailadress$, vbNullString, vbNullString, 5

End Sub
Sub Net_Webpage(Address As String)

'Opens a web page
'Example: Net_Webpage "http://trumedia.gyrate.org"

Dim hwnd As Long

Call ShellExecute(hwnd, "Open", Address$, "", App.Path, 1)

End Sub
Public Function Number_Valid(dastring As Variant) As Boolean

'Finds if a string is a number
'Example: if Number_Valid(text1) = false then exit sub

Number_Valid = IsNumeric(dastring)

End Function
Public Function Number_EvenOrOdd(dastring As String)

'Returns if number is even or odd or decimal
'Example:
'        if Number_EvenorOdd(hscroll1.value) = even then goto continue
'        exit sub
'        continue:

Dim operate As Single

If Number_Valid(dastring$) = False Then Exit Function 'Exit if not a number

If InStr(dastring$, ".") <> 0 Then Number_EvenOrOdd = "decimal": Exit Function 'Set function to decimal and exit

operate! = dastring$ / 2 'Divide number by 2

If InStr(operate!, ".") <> 0 Then 'If . is found number was odd
    
    Number_EvenOrOdd = "odd" 'So set function to odd
    
Else

    Number_EvenOrOdd = "even" 'Else number was even
    
End If

End Function
Public Function Number_RandomCustom(low As Integer, high As Integer)

'Generates random number between and including high and low
'Example:
'        Dim guess As Integer
'        Dim correct As Integer
'        Dim try
'        correct% = Number_RandomCustom(1, 100)
'        start:
'        On Error Resume Next
'        guess% = InputBox("Enter a number from 1 to 100", "Guessing game")
'        If guess% = 0 Then Exit Sub
'        If Number_Valid(guess%) = False Then MsgBox "Enter a number", vbCritical, "error...": GoTo start:
'        If guess% = correct% Then MsgBox "Correct answer", 64, "Good Job!!!": Exit Sub
'        try = MsgBox("Wrong Answer, Try again?", vbYesNo, "Try Again?")
'        If try = vbYes Then GoTo start
'        Exit Sub
        


Dim high2 As Integer
Dim darnd As Integer

high2% = high% - low% + 1 'Fix random so high and low ends are variables
Randomize 'Initialize random number generator
darnd% = Int((Rnd * high2%) + low%) 'Generate random number

Number_RandomCustom = darnd%

End Function
Public Function Number_Random(high As Integer)

'Generates random number from 0 to high using custom version
'Example: picture1.backcolor = Number_Random(255)

Number_Random = Number_RandomCustom(high%, "0")

End Function
Public Sub Program_Pause(duration As Variant)

'Pauses the program for a given duration
'Example:
'        dim i as integer
'        for i% = 1 to len(props$)
'        text1.seltext = mid(props$, i%, 1)
'        next i%

Dim starttime

starttime = Timer 'Set variable to timer

Do While Timer - starttime < duration
    
    DoEvents

Loop

End Sub
Public Sub Program_AboutBox(frm As Form, Caption As String, Optional copyright As Variant)

'Professional looking about box for your program
'Example: Program_AboutBox me, "About", "Trumedia Designs"

If VarType(copyright) = vbString Then
    
    Call ShellAbout(frm.hwnd, Caption$, copyright$, frm.Icon)

Else
    
    Call ShellAbout(frm.hwnd, Caption$, "", frm.Icon)

End If

End Sub
Public Function String_Cryption(dastring As String, encrypt_decrypt)

'This is a weak encryption but still an encryption
'Example:
'        text1.text = String_Cryption(text1.text, 1)...this will encrypt the text in text1
'        text1.text = String_Cryption(text1.text, 2)...this will decrypt the text in text1

Dim i As Integer
Dim char As String
Dim ascii As Integer
Dim newchar As String
Dim NewString As String

Select Case encrypt_decrypt
    
    Case 1 'Case 1 is encrypt
    
        For i% = 1 To Len(dastring$)
            
            char$ = Mid(dastring$, i%, 1) 'Get the char
            ascii% = Asc(char$) + 10 'Change char to asc and add 64 to encrypt
            newchar$ = Chr$(ascii%) 'Change new asc to new character
            NewString$ = NewString$ & newchar$ 'Add new character to encrypted string
        
        Next i
        
    String_Cryption = NewString$

    Case 2 'Case 2 is decrypt
        
        For i% = 1 To Len(dastring$)
            
            char$ = Mid(dastring$, i%, 1) 'Get the char
            ascii% = Asc(char$) - 10 'Change char to asc and subtract the 64 used to encrypt
            newchar$ = Chr$(ascii%) 'Change new asc back to new character
            NewString$ = NewString$ & newchar$ 'Add new character to decrypted string
        
        Next i
        
    String_Cryption = NewString$

End Select

End Function
Public Function String_RandomLetter(capitalize As Boolean)

'Generates random letter
'Example:
'        dim letter as string
'        letter$ = String_RandomLetter(false)

Dim letter As Integer

If capitalize = False Then String_RandomLetter = LCase$(Chr$(Number_RandomCustom(65, 90)))
If capitalize = True Then String_RandomLetter = Chr$(Number_RandomCustom(65, 90))

End Function
Public Function String_Scramble(dastring As String)

'This randomly scrambles a string
'For scramble bots, games whatever
'Example:
'        dim scramble as string
'        scramble$ = String_Scramble(text1.text)

Dim length As Integer
Dim part As String
Dim point As Integer
Dim checkrandom As String
Dim scrstr As String
Dim times As Integer

length% = Len(dastring$)
checkrandom$ = ","
times% = 0

Do

startagain:
    
    point% = Number_Random(length%) + 1
    If point% > length% Then GoTo startagain:
    If InStr(checkrandom$, "," & point% & ",") = 0 Then GoTo skip:

    GoTo startagain:

skip:
    
    checkrandom = checkrandom$ & point% & ","
    scrstr$ = scrstr$ & Mid(dastring$, point%, 1)
    times% = times% + 1

Loop Until times% = length%


String_Scramble = scrstr$

End Function
Public Function String_FirstLine(dastring As String)

'Returns first line of a string
'Example:
'dim lastline as string
'firstline$ = String_FirstLine (text1.text)

Dim rspot As Integer

If InStr(dastring$, Chr$(10)) = 0 Then String_FirstLine = dastring$: Exit Function
rspot% = InStr(dastring$, Chr$(10)) 'Find return
String_FirstLine = left(dastring$, rspot% - 2) 'Get everything to the left of the return not includiong return

End Function
Public Function String_LastLine(dastring As String)

'Returns last line of a string
'Example:
'dim lastline as string
'lastline$ = String_LastLine (text1.text)

Dim i As Integer
Dim rspot As Integer

For i% = Len(dastring$) To 1 Step -1

    If Mid(dastring$, i%, 1) = Chr$(10) Then String_LastLine = Right(dastring$, Len(dastring$) - i%): Exit Function
    
Next i%

End Function
Public Function String_Double(dastring As String)

'Doubles each letter in a string
'Example: msgbox String_Double("premier") would return "pprreemmiieerr"
'I know it's basically useless

Dim i As Integer
Dim dblchr As String
Dim dblstrng As String

For i% = 1 To Len(dastring$) 'Used to get start position of mid function

    dblchr$ = Mid(dastring$, i%, 1) 'Gets character at value of i
    dblstrng$ = dblstrng$ & dblchr$ & dblchr$ 'Doubles character and adds it to end of string
    
Next i%

String_Double = dblstrng$

End Function
Public Function String_Load(fullpath As String)

'Load a text file into a string instead of a textbox
'Example:
'dim songlist as string
'songlist$ = String_Load("C:\My shit\songlist.txt")

Dim datext As String
Dim freenumber

If File_Validity(fullpath$, 3) = False Then Exit Function

freenumber = FreeFile

Open fullpath$ For Input As #freenumber

    String_Load = Input(LOF(freenumber), #freenumber)

Close #freenumber

End Function
Public Function String_Cryption_v2(dastring As String, encrypt_decrypt)

'This is a weak encryption but still an encryption
'Same as original just encrypts into different characters
'Example:
'text1.text = String_Cryption2(text1.text, 1)...this will encrypt the text in text1
'text1.text = String_Cryption2(text1.text, 2)...this will decrypt the text in text1

Dim i As Integer
Dim char As String
Dim ascii As Integer
Dim newchar As String
Dim NewString As String
Dim dachar As String
Dim dastring2 As String
Dim eochar As String

Select Case encrypt_decrypt
    
    Case 1 'Case 1 is encrypt
        
        For i% = 1 To Len(dastring$)
            
            dachar$ = Mid(dastring$, i%, 1)
            dastring2$ = dastring2$ & dachar$ & Number_Random("9")
        
        Next i%
        
        For i% = 1 To Len(dastring2$)
            
            char$ = Mid(dastring2$, i%, 1) 'Get the char
            ascii% = Asc(char$) + 64 'Change char to asc and add 64 to encrypt
            newchar$ = Chr$(ascii%) 'Change new asc to new character
            NewString$ = NewString$ & newchar$ 'Add new character to encrypted string
        
        Next i
        
    String_Cryption_v2 = NewString$

    Case 2 'Case 2 is decrypt
        
        For i% = 1 To Len(dastring$) Step 2
            
            eochar$ = eochar$ & Mid(dastring$, i%, 1)
        
        Next i%
        
        For i% = 1 To Len(eochar$)
            
            char$ = Mid(eochar$, i%, 1) 'Get the char
            ascii% = Asc(char$) - 64 'Change char to asc and subtract the 64 used to encrypt
            newchar$ = Chr$(ascii%) 'Change new asc back to new character
            NewString$ = NewString$ & newchar$ 'Add new character to decrypted string
        
        Next i
        
    String_Cryption_v2 = NewString$

End Select

End Function
Public Function String_Reverse(dastring As String)

'Simply reverses a string
'Example: msgbox String_Reverse("premier") would return "reimerp"
'Basically useless

Dim i As Integer
Dim reverse As String

For i% = Len(dastring$) To 1 Step -1 'Start at the end
    
    reverse$ = reverse$ & Mid(dastring$, i%, 1) 'Get last charcter and add it to variable

Next i%

String_Reverse = reverse$

End Function
Public Sub String_Save(dastring As String, fullpath As String)

'Saves a string instead of saving a texbox
'Example: String_Save List_ToNumberedString(songlistbox)

Dim freenumber

File_CheckReadOnly fullpath$, 1
If File_Validity(fullpath$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile

Open fullpath$ For Output As #freenumber

    Print #freenumber, dastring$ 'Print text to file

Close #freenumber

End Sub
Public Function String_Replace(dastring As String, replacewhat As String, replacewith As String, matchcase As Boolean, messagebox As Boolean)

'This will replace every instance of a string within a string with another string
'Sound confusing?
'Then keep learning
'Example String_Replace aolscreename$, "@aol.com", ""
'This will replace all @aol.com with "" or nothing
'So if aolscreename$ equaled "kwest one@aol.com, premier zero@aol.com"
'It would return kwest one, premier zero

Dim spot As Integer
Dim theleft As String
Dim theright As String
Dim danewstring As String
Dim times As Integer

If matchcase = True And InStr(dastring$, replacewhat$) = 0 Then MsgBox "Search text not found.", 64, "Done...": String_Replace = "": Exit Function
If matchcase = False And InStr(LCase$(dastring$), LCase$(replacewhat$)) = 0 Then MsgBox "Search text not found.", 64, "Done...": String_Replace = "": Exit Function

Do
    If matchcase = False Then
        
        spot% = InStr(LCase$(dastring$), LCase$(replacewhat$))
    
    Else
        
        spot% = InStr(dastring$, replacewhat$)
    
    End If
    
    theleft$ = left(dastring$, spot% - 1)
    theright$ = Right(dastring$, Len(dastring$) - (Len(theleft$) + Len(replacewhat$)))
    dastring$ = theleft$ & replacewith$ & theright$
    times% = times% + 1

Loop Until InStr(dastring$, replacewhat$) = 0

String_Replace = dastring$

If messagebox = True Then MsgBox "Search complete, " & times% & " replacements made.", vbInformation, "Search Complete..."

End Function
Public Function String_TrimNull(dastring As String)

'Removes all null characters chr$(32) or space
'dim newstring as string
'Example: newstring$ = String_TrimNull (mystring$)

String_TrimNull = String_Replace(dastring$, " ", "", False, False)

End Function
Public Function String_TrimChar(dastring As String, chartotrim As String)

'Removes all chosen characters from a string
'dim newstring as string
'Example: newstring$ = String_TrimChar (mystring$, ".")

String_TrimChar = String_Replace(dastring$, chartotrim$, "", False, False)

End Function
Public Sub TextBox_Copy(datextbox As TextBox)

'Copies selected text to clipboard
'Example: TextBox_Copy text1

Clipboard.SetText datextbox.SelText

End Sub

Public Sub TextBox_Cut(datextbox As TextBox)

'Copies selected text to clipboard
'Sets selected text to ""
'Example: TextBox_Cut text1

Clipboard.SetText datextbox.SelText
datextbox.SelText = ""

End Sub
Public Sub TextBox_Paste(datextbox As TextBox)

'Sets selected text in text to clipboards text
'TextBox_Paste text1

datextbox.SelText = Clipboard.GetText
datextbox.SetFocus

End Sub
Public Sub TextBox_SelectAll(datextbox As TextBox)

'Selects all contents of a textbox
'TextBox_SelectAll text1

datextbox.SelStart = 0
datextbox.SelLength = Len(datextbox)
datextbox.SetFocus

End Sub
Public Function TextBox_LineCount(datextbox As TextBox)

'This counts the amount of lines in a text box
'dim linecount as string
'linecount$ = TextBox_LineCount(text1)

Dim i As Integer
Dim count As Integer
Dim look As String

If datextbox.Text = "" Then Exit Function 'Exit if textbox is empty

count% = 1 'set count to one because it has to have atleast one line

For i% = 1 To Len(datextbox.Text)
    
    look$ = Mid(datextbox.Text, i%, 1) 'Search each character
        If look$ = Chr$(13) Then count% = count% + 1 'If character is return[chr$(13)] add 1 to line count

Next i%

TextBox_LineCount = count%

End Function
Public Sub TextBox_Menu(datextbox As TextBox, frm As Form, mnu As Menu)

'If you right click a textbox the standard editing menu appears
'This allows you to replace it with a menu you have created
'I recomend you use it in this fashion...
'Example:
'        In mousedown procedure of a textbox add this
'        If button = 2 then
'        Textbox_Menu text1, form1, mnufile

datextbox.enabled = False
datextbox.enabled = True
frm.PopupMenu mnu

End Sub
Public Sub Textbox_Undo(datextbox As TextBox)

'Undo function
'Example: TextBox_Undo text1

Dim Undoit As Long

On Error Resume Next

Undoit& = SendMessage(datextbox.hwnd, EM_UNDO, 0&, 0&)

End Sub
Public Sub TextBox_Spell(datextbox As TextBox, datext As String, Speed As String)

'This will spell a string into a textbox with a defined speed
'Speeds:    [time in between in letter placed]
'       1[1.00 seconds]
'       2[0.90 seconds]
'       3[0.80 seconds]
'       4[0.70 seconds]
'       5[0.60 seconds]
'       6[0.50 seconds]
'       7[0.40 seconds]
'       8[0.30 seconds]
'       9[0.20 seconds]
'      10[0.10 seconds]

'Example: TextBox_Spell text1, mystory$, 8

Dim i As Integer
Dim speeda As Single
Dim spell As String

If Number_Valid(Speed$) = False Then Exit Sub

If Speed = 1 Then speeda! = 1 'define speeds
If Speed = 2 Then speeda! = 0.9
If Speed = 3 Then speeda! = 0.8
If Speed = 4 Then speeda! = 0.7
If Speed = 5 Then speeda! = 0.6
If Speed = 6 Then speeda! = 0.5
If Speed = 7 Then speeda! = 0.4
If Speed = 8 Then speeda! = 0.3
If Speed = 9 Then speeda! = 0.2
If Speed = 10 Then speeda! = 0.1

For i% = 1 To Len(datext$)

    spell$ = Mid(datext$, i%, 1) 'Set variable to letter
    datextbox.SelText = spell$ 'Add letter to textbox
    Program_Pause speeda! 'Timeout to  control thespeed

Next i%

End Sub
Public Sub TextBox_Find(datextbox As TextBox, findwhat As String, casesensitive As Boolean)

'Will highlight the first instance of a string in a textbox
'Example: TextBox_Find text1, "premier", true

Dim length As Integer
Dim Find As Integer
Dim rcount As Integer

On Error GoTo errorfix: 'If string not found goto label

length% = Len(findwhat$) 'Set variable to length of string to find
If casesensitive = True Then Find% = InStr(datextbox.Text, findwhat$) 'Find string [case sensitive]
If casesensitive = False Then Find% = InStr(LCase$(datextbox.Text), LCase$(findwhat$)) 'Find string
datextbox.SelStart = Find% - 1 'Selstart to beginning of string to find
datextbox.SelLength = length% 'Selength to length of string, find string is now selected
datextbox.SetFocus
Exit Sub

errorfix: MsgBox "Search text not found.", 64, "Done..."

End Sub
Public Sub TextBox_FindNext()



End Sub
Public Sub TextBox_Save(datextbox As TextBox, fullpath As String)

'Saves contents of a textbox
'Example: TextBox_Save text1, dir1.path & "\" & file1.file

Dim freenumber

File_CheckReadOnly fullpath$, 1
If File_Validity(fullpath$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile

Open fullpath$ For Output As #freenumber

    Print #freenumber, datextbox.Text 'Print text to file

Close #freenumber

File_SetNormal fullpath$

End Sub
Public Sub TextBox_Replace(datextbox As TextBox, replacewhat As String, replacewith As String, casesensitive As Boolean, messagebox As Boolean)

'Will replace all instances of a string with another string
'CaseSensitive will match case
'MessageBox will display a msgbox with the number of replacements made
'Example: Textbox_ReplaceAll text1.text, "their", "they're", false, true

Dim replacelength As Integer
Dim replacefind As Integer
Dim rcount As Integer

On Error GoTo errorfix: 'If string not found goto label

Do
    
    replacelength% = Len(replacewhat$) 'Set variable to length of string to replace
    If casesensitive = True Then replacefind% = InStr(datextbox.Text, replacewhat$) 'Find string to replace [case sensitive]
    If casesensitive = False Then replacefind% = InStr(LCase$(datextbox.Text), LCase$(replacewhat$)) 'Find string to replace
    datextbox.SelStart = replacefind% - 1 'Selstart to beginning of string to replace
    datextbox.SelLength = replacelength% 'Selength to length of string, replace string is now selected
    datextbox.SelText = replacewith$ 'Set selected string as string to replace with
        rcount% = rcount% + 1 'keep count of replacements

Loop Until InStr(datextbox.Text, replacewhat$) = 0 'Loop until search string = 0

If messagebox = True Then MsgBox "Search complete, " & rcount% & " replacements made.", vbInformation, "Search Complete..."

Exit Sub

errorfix: MsgBox "Search text not found.", 64, "Done..."

End Sub
Public Sub TextBox_ReplaceHighlighted(datextbox As TextBox, replacewhat As String, replacewith As String, casesensitive As Boolean, messagebox As Boolean)

'Same as textbox_replace but this replaces only highlighted text

Dim seldtext As String

'This is an optional line
'if datextbox.SelText = "" then msgbox "No text selected" ,64,"Error..."

If String_Replace(datextbox.Text$, replacewhat$, replacewith$, casesensitive, messagebox) = "" Then

    Exit Sub

Else

    seldtext$ = datextbox.SelText
    datextbox.SelText = String_Replace(seldtext$, replacewhat$, replacewith$, casesensitive, False)

End If

End Sub
Public Sub TextBox_Load(datextbox As TextBox, fullpath As String)

'Loads text into a text file
'Example:
'        Commondialog1.showopen
'        TextBox_Load text1, Commondialog1.filename

Dim datext As String
Dim freenumber

If File_Validity(fullpath$, 3) = False Then Exit Sub

freenumber = FreeFile

Open fullpath$ For Input As #freenumber

    datext$ = Input(LOF(freenumber), #freenumber)

Close #freenumber

datextbox.Text = datext$

End Sub
Public Sub Window_Close(window As Long)

'Closes a given window
'Example:
'        dim aol as long
'        aol% = findwindow("Aol Frame25", vbnullstring)
'        Window_Close aol%

Call PostMessage(window&, WM_CLOSE, 0&, 0&)

End Sub
Public Sub Window_Hide(hwnd As Long)

'Hides given window
'Example:
'        dim aol as long
'        aol% = findwindow("Aol Frame25", vbnullstring)
'        Window_Hide aol%

Call ShowWindow(hwnd&, SW_HIDE)

End Sub
Public Sub Window_Show(hwnd As Long)
    
'Shows a given window
'Example:
'        dim aol as long
'        aol% = findwindow("Aol Frame25", vbnullstring)
'        Window_Show aol%
    
Call ShowWindow(hwnd&, SW_SHOW)

End Sub
Public Sub Window_SetText(winder As Long, txt2set As String)

'Will set the text of an outside textbox to a given text
'Example:
'        dim notepad as long, edit as long
'        notepad& = FindWindow("Notepad", vbNullString)
'        edit& = FindWindowEx(parenthandle&, 0&, "Edit", vbNullString)
'        Window_SetText edit&, "premier is my boy"
'        If notepad is open that will set the text to premier is my boy

Dim DoIt As Long

On Error Resume Next

DoIt& = SendMessageByString(winder&, WM_SETTEXT, 0, txt2set$)

End Sub

