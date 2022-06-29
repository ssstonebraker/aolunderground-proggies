Attribute VB_Name = "integerx32"
Option Explicit

Public Sub AddWinner(Winner As String, Points As Long, List As ListBox)
    Dim lngCount As Long, strItem As String, lngPoints As Long
    Dim intLengh As Integer
    For lngCount& = 0 To List.ListCount - 1
        strItem$ = List.List(lngCount&)
        If Left$(strItem$, Len(Winner$)) = Winner$ Then
            intLengh% = Len(strItem$) - (Len(Winner$) + 3)
            lngPoints& = Val(Mid$(strItem$, InStr(strItem$, " - ") + 3, intLengh%))
            lngPoints& = lngPoints& + Points&
            List.List(lngCount&) = Winner$ & " - " & lngPoints&
            Exit Sub
        End If
    Next lngCount&
    List.AddItem Winner$ & " - " & Points&
End Sub
Public Function isItemOn(Item As String, List As ListBox) As Boolean
    Dim lngCount As Long
    For lngCount& = 0 To List.ListCount - 1
        If List.List(lngCount&) = Item Then
            isItemOn = True
            Exit Function
        End If
    Next lngCount&
End Function
Public Sub loadSystemFonts(List As ListBox)
    Dim lngCount As Long
    For lngCount& = 0 To Screen.FontCount - 1
        List.AddItem Screen.Fonts(lngCount&)
    Next lngCount&
End Sub
Public Sub loadPrinterFonts(List As ListBox)
    Dim lngCount As Long
    For lngCount& = 0 To Printer.FontCount - 1
        List.AddItem Printer.Fonts(lngCount&)
    Next lngCount&
End Sub
Public Sub ShuffleList(List As ListBox)
    Dim lngIndex As Long, strItem As String, lngItemData As Long
    
    If List.Sorted Then Exit Sub
    If List.ListCount < 2 Then Exit Sub
    Randomize
    
    For lngIndex& = 1 To List.ListCount - 1
        strItem$ = List.List(lngIndex&)
        lngItemData& = List.ItemData(lngIndex&)
        List.RemoveItem (lngIndex&)
        List.AddItem strItem$, Int((lngIndex& + 1) * Rnd)
        List.ItemData(List.NewIndex) = lngItemData&
    Next
End Sub
Public Sub SortWinners(List As ListBox)
    Dim lngCount As Long, lngCount2 As Long, strHold As String
    Dim FirstVal As Long, SecondVal As Long
    
    For lngCount& = List.ListCount - 1 To 0 Step -1
        For lngCount2& = 0 To lngCount&
            If lngCount2& <> List.ListCount - 1 Then
                FirstVal& = CLng(Mid$(List.List(lngCount&), InStr(List.List(lngCount&), " - ") + 3))
                SecondVal& = CLng(Mid$(List.List(lngCount2&), InStr(List.List(lngCount2&), " - ") + 3))
                If FirstVal& > SecondVal& Then
                    strHold$ = List.List(lngCount&)
                    List.List(lngCount&) = List.List(lngCount2&)
                    List.List(lngCount2&) = strHold$
                End If
            End If
        Next lngCount2&
    Next lngCount&
End Sub
Public Sub loadListBox(Path As String, List As ListBox)
    Dim strItem As String
    
    Open Path$ For Input As #1
    Do While Not EOF(1)
        Line Input #1, strItem$
        If Not (ListItem$ = "") Then
            List.AddItem strItem$
        End If
    Loop
    Close #1
End Sub
Public Sub saveListBox(Path As String, List As ListBox)
    Dim strItem As String, lngCount As Long
    
    Open Path$ For Output As #1
    For lngCount& = 0 To List.ListCount - 1
        Print #1, List.List(lngCount&) + Chr(13)
    Next lngCount&
    Close #1
End Sub
