Attribute VB_Name = "modFakeChat"
Option Explicit

Public Sub DoChatStuff(strSN As String, strSaid As String, blnRTF As Boolean)
    Dim lngSpot As Long
    If strSN$ <> "" And strSaid$ <> "" Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 8
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelUnderline = False
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelColor = vbBlue
        If frmChat.txtChat.Text = "" Then
            frmChat.txtChat.SelText = strSN$ & ":" & Chr(9)
        Else
            frmChat.txtChat.SelText = vbCrLf & strSN$ & ":" & Chr(9)
        End If
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 10
        frmChat.txtChat.SelBold = False
        frmChat.txtChat.SelColor = vbBlack
        lngSpot& = Len(frmChat.txtChat.Text)
        If blnRTF = True Then
            frmChat.txtChat.SelRTF = strSaid$
        Else
            frmChat.txtChat.SelText = strSaid$
        End If
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelStart = lngSpot&
        frmChat.txtChat.SelLength = Len(frmChat.txtChat.Text) - lngSpot&
        frmChat.txtChat.SelHangingIndent = 1400
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
    End If
End Sub

Public Sub DoChatStuff2(strSN As String, strSaid As String, blnRTF As Boolean)
    Dim lngSpot As Long
    If strSN$ <> "" And strSaid$ <> "" Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 8
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelUnderline = False
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelColor = vbRed
        If frmChat.txtChat.Text = "" Then
            frmChat.txtChat.SelText = strSN$ & ":" & Chr(9)
        Else
            frmChat.txtChat.SelText = vbCrLf & strSN$ & ":" & Chr(9)
        End If
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 10
        frmChat.txtChat.SelBold = False
        frmChat.txtChat.SelColor = vbBlack
        lngSpot& = Len(frmChat.txtChat.Text)
        If blnRTF = True Then
            frmChat.txtChat.SelRTF = strSaid$
        Else
            frmChat.txtChat.SelText = strSaid$
        End If
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelStart = lngSpot&
        frmChat.txtChat.SelLength = Len(frmChat.txtChat.Text) - lngSpot&
        frmChat.txtChat.SelHangingIndent = 1400
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
    End If
End Sub

Public Sub RandomStuff()
    Dim intPhrase As Integer, intPerson As Integer
    Dim intReply As Integer
    Randomize
    intPhrase% = Int((6 * Rnd) + 1)
'intPerson% = Int(frmChat.lstNames.ListCount * Rnd)
    intReply% = Int((6 * Rnd) + 1)
    Select Case intPhrase
        Case 1
           ' Call DoChatStuff("PeaceX101", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\fswiss Arial;}{\f4\froman Times New Roman;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red0\green128\blue128;\red128\green128\blue0;}\deflang1033\pard\plain\f4\fs24\cf2\ul Welcome \plain\f4\fs24\cf3\ul " & frmChat.lstNames.List(intPerson%) & "\plain\f2\fs20\par }", True)
            Call Pause(0.4)
            Select Case intReply%
                Case 1
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "turn that damn bot off peace", False)
                Case 2
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "30 years old and still programming a welcome bot peace?", False)
                Case 3
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "stfu peace before i reach through your screen and...", False)
                Case 4
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "peace, you are so gay", False)
                Case 5
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "is peace a bot?", False)
                Case 6
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "thanks peace. i couldn't call this room lame if you weren't here", False)
            End Select
        Case 2
        '    Call DoChatStuff("Izekial83", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\froman Times New Roman;}{\f4\fswiss Arial;}{\f5\froman Arial;}{\f6\fswiss Tahoma;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red128\green128\blue0;\red0\green128\blue128;\red0\green0\blue128;}\deflang1033\pard\plain\f6\fs24\cf4\b ^\plain\f3\fs24\cf4\b  izekial's ga\plain\f3\fs24\cf1\b y prog\plain\f2\fs20\par }", True)
            Call Pause(0.2)
            Call DoChatStuff("Izekial83", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\froman Times New Roman;}{\f4\fswiss Arial;}{\f5\froman Arial;}{\f6\fswiss Tahoma;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red128\green128\blue0;\red0\green128\blue128;\red0\green0\blue128;}\deflang1033\pard\plain\f6\fs24\cf4\b ^\plain\f3\fs24\cf4\b  100% copied\plain\f3\fs24\cf1\b  code\plain\f2\fs20\par }", True)
            Call Pause(0.2)
            Call DoChatStuff("Izekial83", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\froman Times New Roman;}{\f4\fswiss Arial;}{\f5\froman Arial;}{\f6\fswiss Tahoma;}{\f7\fswiss Tahoma;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red128\green128\blue0;\red0\green128\blue128;\red0\green0\blue128;}\deflang1033\pard\plain\f7\fs24\cf4\b ^\plain\f3\fs24\cf4\b  coded by ize\plain\f3\fs24\cf1\b kial83\plain\f2\fs20\par }", True)
        Case 3
            Call DoChatStuff("MacroBoy", "PReSs 555 eF JeW WaNnA gOin a PHat kNEw gRwP", False)
            Call Pause(0.2)
            Call DoChatStuff("MacroBoy", "555", False)
        Case 4
            Call DoChatStuff("MaGuSHaVoK", "Wanna Make A Prog Wit Me??? Anybody??? Please???", False)
        Case 5
            Call DoChatStuff("It Be Mi", "this room blows", True)
            Call Pause(0.4)
          '  Call DoChatStuff(frmChat.lstNames.List(intPerson%), "right on mi", False)
        Case 6
            'do nothing
    End Select
End Sub

Public Sub Pause(Duration As Variant)
    Dim Current As Variant
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
