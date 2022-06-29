Attribute VB_Name = "modCompare"
Option Explicit

Public Sub LoadString(txtString As String, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtString$ = TextString$
End Sub

Public Function GetBasTitle(bString As String) As String
    Dim Spot1 As Long, Spot2 As Long
    Spot1& = InStr(bString$, "Attribute VB_Name = " & Chr(34))
    Spot1& = Spot1& + 20&
    Spot2& = InStr(Spot1& + 1, bString$, Chr(34))
    If Spot1& = 0& Or Spot2& = 0& Then
        GetBasTitle$ = "unknown"
        Exit Function
    End If
    GetBasTitle$ = Mid(bString$, Spot1& + 1&, Spot2& - Spot1& - 1&)
End Function

Public Function LineCount(MyString As String) As Long
    Dim Spot As Long, Count As Long
    If Len(MyString$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        LineCount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    LineCount& = LineCount& + 1
End Function

Public Function LineFromString(MyString As String, Line As Long) As String
    Dim theline As String, Count As Long
    Dim fSpot As Long, lSpot As Long, DoIt As Long
    Count& = LineCount(MyString$)
    If Line& > Count& Then
        Exit Function
    End If
    If Line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = Replace(theline$, Chr(13), "")
        theline$ = Replace(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        fSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            lSpot& = fSpot&
            fSpot& = InStr(fSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If fSpot = 0 Then
            fSpot = Len(MyString$)
        End If
        theline$ = Mid(MyString$, lSpot&, fSpot& - lSpot& + 1)
        theline$ = Replace(theline$, Chr(13), "")
        theline$ = Replace(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function

Public Function GetSubTitle(SubString As String) As String
    Dim SubSpot As Long, FunctionSpot As Long, pSpot As Long
    Dim UseSpot As Long
    pSpot& = InStr(SubString$, "(")
    SubSpot& = InStr(SubString$, "Sub ")
    FunctionSpot& = InStr(SubString$, "Function ")
    If SubSpot& = 0& Then SubSpot& = 1000&
    If FunctionSpot& = 0& Then FunctionSpot& = 1000&
    If SubSpot& < FunctionSpot& Then
        UseSpot& = SubSpot& + 4&
    Else
        UseSpot& = FunctionSpot& + 9&
    End If
    GetSubTitle$ = Mid(SubString$, UseSpot&, pSpot& - UseSpot&)
End Function

Public Function GetComments(cString As String) As String
    Dim CurLine As String, lCount As Long, DoIt As Long
    Dim TempString As String
    lCount& = LineCount(cString$)
    For DoIt& = 1& To lCount&
        DoEvents
        CurLine$ = LineFromString(cString$, DoIt&)
        CurLine$ = Replace(CurLine$, Chr(9), "")
        CurLine$ = Trim(CurLine$)
        If InStr(CurLine$, Chr(39)) = 1 Then
            If TempString$ = "" Then
                TempString$ = CurLine$
            Else
                TempString$ = TempString$ & vbCrLf & CurLine$
            End If
        End If
    Next DoIt&
    GetComments$ = TempString$
End Function

Public Function GetVariables(SubString As String) As String
    Dim CurLine As String, sCount As Long, tString As String
    Dim DoThis As Long, vCount As Long, vSpot As Long
    Dim vLine As String, vString As String, DoV As Long
    sCount& = LineCount(SubString$)
    For DoThis& = 1& To sCount&
        DoEvents
        CurLine$ = LineFromString(SubString$, DoThis&)
        CurLine$ = Replace(CurLine$, Chr(9), "")
        CurLine$ = Trim(CurLine$)
        If InStr(1, CurLine$, "dim", 1) = 1& Then
            If tString$ = "" Then
                tString$ = CurLine$
            Else
                tString$ = tString$ & vbCrLf & CurLine$
            End If
        End If
    Next DoThis
    tString$ = Replace(tString$, "Dim ", "")
    tString$ = Replace(tString$, ", ", vbCrLf)
    vCount& = LineCount(tString$)
    For DoV& = 1& To vCount&
        DoEvents
        vLine$ = LineFromString(tString$, DoV&)
        vSpot& = InStr(vLine$, " As")
        If vSpot& <> 0& Then
            vLine$ = Left(vLine$, vSpot& - 1&)
        End If
        vLine$ = Trim(vLine$)
        If vString$ = "" Then
            vString$ = vLine$
        Else
            vString$ = vString$ & vbCrLf & vLine$
        End If
    Next DoV&
    GetVariables$ = vString$
End Function

Public Function GetArguments(SubString As String) As String
    Dim FirstLine As String, aCount As Long, aLine As String
    Dim aString As String, Spot1 As Long, Spot2 As Long
    Dim tString As String, DoA As Long, aSpot As Long
    FirstLine$ = LineFromString(SubString$, 1&)
    Spot1& = InStr(FirstLine$, "(")
    Spot2& = InStr(FirstLine$, ")")
    tString$ = Mid(FirstLine$, Spot1& + 1, Spot2& - Spot1& - 1&)
    tString$ = Replace(tString$, ", ", vbCrLf)
    aCount& = LineCount(tString$)
    For DoA& = 1& To aCount&
        DoEvents
        aLine$ = LineFromString(tString$, DoA&)
        aSpot& = InStr(aLine$, " As")
        If aSpot& <> 0& Then
            aLine$ = Left(aLine$, aSpot& - 1&)
        End If
        aLine$ = Trim(aLine$)
        If aString$ = "" Then
            aString$ = aLine$
        Else
            aString$ = aString$ & vbCrLf & aLine$
        End If
    Next DoA&
    GetArguments$ = aString$
End Function

Public Function KillTabs(cString As String) As String
    Dim CurLine As String, lCount As Long, DoIt As Long
    Dim TempString As String
    lCount& = LineCount(cString$)
    For DoIt& = 1& To lCount&
        DoEvents
        CurLine$ = LineFromString(cString$, DoIt&)
        CurLine$ = Replace(CurLine$, Chr(9), "")
        CurLine$ = Trim(CurLine$)
        If CurLine$ <> "" Then
            If TempString$ = "" Then
                TempString$ = CurLine$
            Else
                TempString$ = TempString$ & vbCrLf & CurLine$
            End If
        End If
    Next DoIt&
    KillTabs$ = TempString$
End Function

Public Function KillComments(cString As String) As String
    Dim CurLine As String, lCount As Long, DoIt As Long
    Dim TempString As String
    lCount& = LineCount(cString$)
    For DoIt& = 1& To lCount&
        DoEvents
        CurLine$ = LineFromString(cString$, DoIt&)
        CurLine$ = Replace(CurLine$, Chr(9), "")
        CurLine$ = Trim(CurLine$)
        If InStr(CurLine$, Chr(39)) <> 0& Then
            If InStr(CurLine$, Chr(39)) = 1 Then
                CurLine$ = ""
            Else
                Do
                    CurLine$ = Left(CurLine$, InStr(CurLine$, Chr(39)) - 1)
                Loop Until InStr(CurLine$, Chr(39)) = 0
            End If
        End If
        If CurLine$ <> "" Then
            If DoIt& = 1& Then
                If InStr(LineFromString(cString$, DoIt&), Chr(39)) <> 1 Then
                    If InStr(LineFromString(cString$, DoIt&), Chr(39)) <> 0 Then
                        TempString$ = Left(LineFromString(cString$, DoIt&), InStr(LineFromString(cString$, DoIt&), Chr(39)) - 1)
                    Else
                        TempString$ = LineFromString(cString$, DoIt&)
                    End If
                End If
            Else
                If InStr(LineFromString(cString$, DoIt&), Chr(39)) <> 1 Then
                    If InStr(LineFromString(cString$, DoIt&), Chr(39)) <> 0 Then
                        TempString$ = TempString$ & vbCrLf & Left(LineFromString(cString$, DoIt&), InStr(LineFromString(cString$, DoIt&), Chr(39)) - 1)
                    Else
                        TempString$ = TempString$ & vbCrLf & LineFromString(cString$, DoIt&)
                    End If
                End If
            End If
        End If
    Next DoIt&
    KillComments$ = TempString$
End Function

Public Function KillDims(SubString As String) As String
    Dim CurLine As String, sCount As Long, tString As String
    Dim DoThis As Long, vCount As Long, vSpot As Long
    Dim vLine As String, vString As String, DoV As Long
    sCount& = LineCount(SubString$)
    For DoThis& = 1& To sCount&
        DoEvents
        CurLine$ = LineFromString(SubString$, DoThis&)
        CurLine$ = Replace(CurLine$, Chr(9), "")
        CurLine$ = Trim(CurLine$)
        If InStr(1, CurLine$, "dim", 1) <> 1& Then
            If tString$ = "" Then
                tString$ = LineFromString(SubString$, DoThis&)
            Else
                tString$ = tString$ & vbCrLf & LineFromString(SubString$, DoThis&)
            End If
        End If
    Next DoThis
    KillDims$ = tString$
End Function

Public Function ReplaceVariables(SubStr As String) As String
    Dim TempStr As String, VarStr As String, DoIt As Long
    TempStr$ = SubStr$
    VarStr$ = GetVariables(TempStr$)
    For DoIt& = 1& To LineCount(TempStr$)
        TempStr$ = Replace(TempStr$, LineFromString(VarStr$, DoIt&), "*")
    Next DoIt&
    ReplaceVariables$ = TempStr$
End Function

Public Function ReplaceArguements(SubStr As String) As String
    Dim TempStr As String, VarStr As String, DoIt As Long
    TempStr$ = SubStr$
    VarStr$ = GetArguments(TempStr$)
    For DoIt& = 1& To LineCount(TempStr$)
        TempStr$ = Replace(TempStr$, LineFromString(VarStr$, DoIt&), "*")
    Next DoIt&
    ReplaceArguements$ = TempStr$
End Function

Public Function ReplaceTitle(SubStr As String) As String
    Dim TempStr As String, VarStr As String, DoIt As Long
    TempStr$ = SubStr$
    VarStr$ = GetSubTitle(TempStr$)
    TempStr$ = Replace(TempStr$, VarStr$, "*")
    ReplaceTitle$ = TempStr$
End Function
