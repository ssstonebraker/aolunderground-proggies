Attribute VB_Name = "LoadTimes"
Function LoadTimes(FileName)
        free = FreeFile
    Open FileName For Random As free
    Close free
    Open FileName For Input As free
i = FileLen(FileName)
X = Input(i, free)
    Close free

    Open FileName For Output As #1
X = Val(X) + 1
    Print #1, X
    Close #1
        LoadTimes = X
End Function




Function Wait(HLong)
'Same as timeout
Current = Timer
Do While Timer - Current < Val(HLong)
DoEvents
Loop
End Function


Function FileInput2(FileName As String)
free = FreeFile
Open FileName For Input As free
    i = FileLen(FileName)
    X = Input(i - 2, free)
Close free
    FileInput2 = X
End Function
Function FileInput(FileName As String)
free = FreeFile
Open FileName For Input As free
    i = FileLen(FileName)
    X = Input(i, free)
Close free
    FileInput = X
End Function

Function FileLoadList(FileName As String, Lis As listbox)
On Error Resume Next
Open FileName For Input As #1
Do While Not EOF(1)
 Line Input #1, Ln$
Lis.AddItem Ln$
Loop
Close #1
End Function



Function FileSaveList(FileName As String, Lis As listbox)
free = FreeFile
Open FileName For Output As free
For X = 0 To Lis.ListCount
Print #1, Lis.List(X)
Next X
Close #1
End Function

