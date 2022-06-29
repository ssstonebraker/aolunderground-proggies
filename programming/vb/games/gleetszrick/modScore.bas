Attribute VB_Name = "modScore"
Option Explicit

' ---------------------------------------------------------------------
' Global constants, variables, and others used within the game.
'
' *********************************************
' | @ Written by Pranay Uppuluri. @           |
' | @ Copyright (c) 1997-98 Pranay Uppuluri @ |
' *********************************************
'
' VB game example Break-Thru! by Mark Pruett ported
' to Visual Basic DirectX.
'
' Thanks for Patrice Scribe's DirectX.TLB for DirectX 3.0 or Higher,
' his dixuSprite Class, and his dixu module, this game looks
' to be easy to code.
'
' You can visit Patrice's home page at:
'
'           http://www.chez.com/scribe/  *OR*
'           http://ourworld.compuserve.com/homepages/pscribe/
'
' If it wasn't for his effort, I would have had to do a lot
' more coding than this!
'
' modScore.BAS
' Contains code for High Score Testing, saving, etc.
' ---------------------------------------------------------------------

' The maximum number of High Scores tracked.
Public Const MAX_HISCORES = 8

' The CRegSetting Class for modifying the Windows 32 bit Registry
Public reg As CRegSettings

' Used when calling SaveSetting, GetSetting, and DeleteSetting
' methods in the CRegSetting Class.
Public Const SECTION = "High Scores"
Public Const ENTRY = "Score"

' Data Structure that stores high scores
Type tScores
    pName As String
    lngScore As Long
End Type

' This should ALWAYS be set to MAX_HISCORES + 1.
Public Hi(1 To 9) As tScores

' The current number of high scores stored in Hi() array.
Public Num_HiScores As Long

' The new score to test.
Public gNewScore As Long

' This public Boolean tells the form to just display
' the high scores (no player name data entry).
Public gDisplayOnly As Boolean

Public Sub AddScoreAndSave(ByVal NewName As String, ByVal NewScore As Long)
' ------------------------------------------------------------------------
' Add this new score to the list of high scores and save everything back
' to the Registry database.
' ------------------------------------------------------------------------
Dim i As Long
Dim j As Long
Dim temp As tScores

    ' Add the new score to the end of the Hi() array.
    Hi(Num_HiScores + 1).pName = NewName
    Hi(Num_HiScores + 1).lngScore = NewScore
    
    ' Bubble-sort the scores in descening order (highest first)...
    For j = 1 To Num_HiScores + 1
        For i = 2 To Num_HiScores + 1
            If Hi(i).lngScore > Hi(i - 1).lngScore Then
                temp = Hi(i - 1)
                Hi(i - 1) = Hi(i)
                Hi(i) = temp
            End If
        Next i
    Next j
    
    If Num_HiScores < MAX_HISCORES Then Num_HiScores = Num_HiScores + 1
    
    ' Write the scores back to the Registry
    For i = 1 To Num_HiScores
        WriteScore i, Format(Hi(i).lngScore & ";" & Trim(Hi(i).pName))
    Next i
    
End Sub

Public Sub GetScores()
' ---------------------------------------------------------------
' Read scores from the registry and store them in the Hi()
' array. In the registry, the score and the player's name are
' stored together, seperated by a semicolon like this:
'
' 2535;Pranay Uppuluri
'
' We seperate the two pieces of data after reading them in.
' ----------------------------------------------------------------
Dim i As Long
Dim rc As String
Dim pos As Long
Dim AString As String

    For i = 1 To MAX_HISCORES
        AString = Space(255)
        
        ' Retrieve the data from the registry
        rc = reg.GetSetting(SECTION, ENTRY & Format(i))
        
        If Len(rc) > 0 Then
            ' rc tells us the length of the returned string,
            ' so we truncate the string at that length.
            AString = Left(rc, Len(rc))
            
            ' Seperate the player's name and score.
            pos = InStr(AString, ";")
            If pos > 0 Then
                Hi(i).lngScore = Left(AString, pos - 1)
                Hi(i).pName = Mid(AString, pos + 1)
            End If
        Else
            Num_HiScores = i - 1
            Exit Sub
        End If
    Next i
    Num_HiScores = MAX_HISCORES
End Sub

Public Function IsAHiScore(ByVal NewScore As Long) As Boolean
' -----------------------------------------------------------
' Returns True if NewScore is a High Score, False otherwise.
' -----------------------------------------------------------
Dim i As Long

    ' Assume that it's not a high score.
    IsAHiScore = False
    
    If Num_HiScores > 0 Then
        ' If we've only equalled the lowest high score,
        ' then don't bother...
        If Num_HiScores = MAX_HISCORES And (NewScore = Hi(Num_HiScores).lngScore) Then
            Exit Function
        End If
    End If
    
    ' If we haven't filled up the High Scores table,
    ' then this must be a new high score.
    If Num_HiScores < MAX_HISCORES Then
        IsAHiScore = True
        Exit Function
    End If
    
    ' Compare this new score to the existing high scores.
    For i = 1 To Num_HiScores
        If Hi(i).lngScore < NewScore Then
            IsAHiScore = True
            Exit For
        End If
    Next i
End Function

Sub WriteScore(ByVal EntryNum As Long, ByVal AString As String)
' --------------------------------------------------------------
' Write a high score back to the registry.
' --------------------------------------------------------------
Dim rc As Long

    rc = reg.SaveSetting(SECTION, ENTRY & Format(EntryNum), AString)
End Sub
