Attribute VB_Name = "modReplace"
Option Explicit

'this function is for you those misfortunate souls who do not have
'visual basic 6. 'replace' is a new function in vb6, and i often
'use it in my examples. before vb6 though, i had written my own
'replace function, which was originally called 'replacestring'. if
'you are using an earlier version of visual basic, simply add this
'module or copy this function into the project.

'dos - dos@hider.com
'www.hider.com/dos

Public Function Replace(ByVal strMain As String, strFind As String, strReplace As String) As String
    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    Replace$ = strNew$
End Function
