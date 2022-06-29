Attribute VB_Name = "Module1"

' Search sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com

Option Explicit

Global bCancel As Boolean

Function DirCount() As Integer
    Dim A$, N As Integer
    N = 0
    A$ = Dir$("*.*", vbDirectory)
    Do While Len(A$)
        DoEvents
        If bCancel Then Exit Function
        If GetAttr(A$) And vbDirectory Then
            Select Case A$
                Case ".", ".."
                Case Else
                    N = N + 1
            End Select
        End If
        A$ = Dir$
    Loop
    DirCount = N
End Function

Sub FindText(Drv$, FileMask$, Files$(), NumFiles As Integer, Srch$)
    Dim Cdir$
    Cdir$ = CurDir$
    ChDrive Drv$
    ChDir "\"
    NumFiles = 0
    If Srch$ = "" Then Exit Sub
    ReDim Files$(0)
    
    Call SearchInDir(Drv$ & ":\", FileMask$, Files$(), NumFiles, Srch$)
    
    ChDrive Cdir$
    ChDir Cdir$
    If bCancel Then Exit Sub
End Sub

Function LoadDirs(Dirs$())
    Dim A$, N As Integer
    N = 0
    A$ = Dir$("*.*", vbDirectory)
    Do While Len(A$)
        DoEvents
        If bCancel Then Exit Function
        If GetAttr(A$) And vbDirectory Then
            Select Case A$
                Case ".", ".."
                Case Else
                    N = N + 1
                    Dirs$(N) = A$
            End Select
        End If
        A$ = Dir$
    Loop
End Function


Sub SearchInDir(Dr$, FileMask$, Files$(), NumFiles, Srch$)
    Dim NumDirs As Integer, CurNumFiles As Integer, i As Integer, j As Integer
    NumDirs = DirCount
    If NumDirs Then
        ReDim Dirs$(1 To NumDirs)
        Call LoadDirs(Dirs$())
        If bCancel Then Exit Sub
    End If
    
    CurNumFiles = NumFiles
    Call SearchForText(FileMask$, Files$(), NumFiles, Srch$)
    If Right$(Dr$, 1) <> "\" Then Dr$ = Dr$ + "\"
    For i = CurNumFiles + 1 To NumFiles
        Files$(i) = Dr$ & Files$(i)
        Form1!List1.AddItem Files$(i)
    Next
    
    For j = 1 To NumDirs
        ChDir Dirs$(j)
        Call SearchInDir(CurDir$, FileMask$, Files$(), NumFiles, Srch$)
        ChDir ".."
        If bCancel Then Exit Sub
    Next
End Sub
    
Sub SearchForText(FileMask$, Files$(), NumFiles, Srch$)
    On Local Error Resume Next
    Dim A$, Bf As Integer, B$, C$, i As Long, S As Long, N As Long, Found As Boolean, Cdir$, j As Integer, CurNumFiles As Integer
    Bf = FreeFile
    A$ = Dir$(FileMask$)
    Do While Len(A$)
        Form1!Label2.Caption = "Searching " & A$
        DoEvents
        If bCancel Then Exit Sub
        Found = False
        Open A$ For Binary Shared As Bf
        S = LOF(Bf)
        N = S \ 4096
        C$ = ""
        If N Then
            For i = 1 To N
                B$ = Space$(4096)
                Get Bf, , B$
                B$ = C$ + B$
                C$ = Right$(B$, Len(Srch$))
                If InStr(UCase$(B$), Srch$) Then
                    NumFiles = NumFiles + 1
                    ReDim Preserve Files$(NumFiles)
                    Files$(NumFiles) = A$
                    Found = True
                    Exit For
                End If
            Next
        End If
        If Not Found Then
            N = S Mod 4096
            If N Then
                B$ = Space$(N)
                Get Bf, , B$
                B$ = C$ + B$
                If InStr(UCase$(B$), Srch$) Then
                    NumFiles = NumFiles + 1
                    ReDim Preserve Files$(NumFiles)
                    Files$(NumFiles) = A$
                    Found = True
                End If
            End If
        End If
        Close Bf
        A$ = Dir$
    Loop
    
End Sub


