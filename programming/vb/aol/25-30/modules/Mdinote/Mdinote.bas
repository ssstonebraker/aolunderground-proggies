Option Explicit

Global Const modal = 1
Global Const CASCADE = 0
Global Const TILE_HORIZONTAL = 1
Global Const TILE_VERTICAL = 2
Global Const ARRANGE_ICONS = 3

Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type
Global FState()  As FormState
Global Document() As New frmNotePad
Global gFindString, gFindCase As Integer, gFindDirection As Integer
Global gCurPos As Integer, gFirstTime As Integer
Global ArrayNum As Integer

' API functions used to read and write to MDINOTE.INI.
' Used for handling the recent files list.
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer

Function AnyPadsLeft () As Integer
    Dim i As Integer

    ' Cycle throught the document array.
    ' Return True if there is at least one
    ' open document remaining.
    For i = 1 To UBound(Document)
        If Not FState(i).Deleted Then
            AnyPadsLeft = True
            Exit Function
        End If
    Next

End Function

Sub CenterForm (frmParent As Form, frmChild As Form)
' This procedure centers a child form over a parent form.
' Calling this routine loads the dialog. Use the Show method
' to display the dialog after calling this routine ( ie MyFrm.Show 1)

Dim l, t
  ' get left offset
  l = frmParent.Left + ((frmParent.Width - frmChild.Width) / 2)
  If (l + frmChild.Width > screen.Width) Then
    l = screen.Width = frmChild.Width
  End If

  ' get top offset
  t = frmParent.Top + ((frmParent.Height - frmChild.Height) / 2)
  If (t + frmChild.Height > screen.Height) Then
    t = screen.Height - frmChild.Height
  End If

  ' center the child formfv
  frmChild.Move l, t

End Sub

Sub EditCopyProc ()
    ' Copy selected text to Clipboard.
    ClipBoard.SetText frmMDI.ActiveForm.ActiveControl.SelText
End Sub

Sub EditCutProc ()
    ' Copy selected text to Clipboard.
    ClipBoard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Delete selected text.
    frmMDI.ActiveForm.ActiveControl.SelText = ""
End Sub

Sub EditPasteProc ()
    ' Place text from Clipboard into active control.
    frmMDI.ActiveForm.ActiveControl.SelText = ClipBoard.GetText()
End Sub

Sub FileNew ()
    Dim fIndex As Integer

    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = "Untitled:" & fIndex
    Document(fIndex).Show

    ' Make sure toolbar edit buttons are visible
    frmMDI!imgCutButton.Visible = True
    frmMDI!imgCopyButton.Visible = True
    frmMDI!imgPasteButton.Visible = True
    
End Sub

Function FindFreeIndex () As Integer
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    ' Cycle throught the document array. If one of the
    ' documents has been deleted, then return that
    ' index.
    For i = 1 To ArrayCount
        If FState(i).Deleted Then
            FindFreeIndex = i
            FState(i).Deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.

    ReDim Preserve Document(ArrayCount + 1)
    ReDim Preserve FState(ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function

Sub FindIt ()
    Dim start, pos, findstring, sourcestring, msg, response, Offset
    
    If (gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart) Then
        Offset = 1
    Else
        Offset = 0
    End If

    If gFirstTime Then Offset = 0

    start = frmMDI.ActiveForm.ActiveControl.SelStart + Offset
        
    If gFindCase Then
        findstring = gFindString
        sourcestring = frmMDI.ActiveForm.ActiveControl.Text
    Else
        findstring = UCase(gFindString)
        sourcestring = UCase(frmMDI.ActiveForm.ActiveControl.Text)
    End If
            
    If gFindDirection = 1 Then
        pos = InStr(start + 1, sourcestring, findstring)
    Else
        For pos = start - 1 To 0 Step -1
            If pos = 0 Then Exit For
            If Mid(sourcestring, pos, Len(findstring)) = findstring Then Exit For
        Next
    End If

    ' If string is found
    If pos Then
        frmMDI.ActiveForm.ActiveControl.SelStart = pos - 1
        frmMDI.ActiveForm.ActiveControl.SelLength = Len(findstring)
    Else
        msg = "Cannot find " & Chr(34) & gFindString & Chr(34)
        response = MsgBox(msg, 0, App.Title)
    End If
    
    gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart
    gFirstTime = False

End Sub

Sub GetRecentFiles ()
  Dim retval, key, i, j
  Dim IniString As String

  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)

  ' Get recent file strings from MDINOTE.INI
  For i = 1 To 4
    key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "MDINOTE.INI")
    If retval And Left(IniString, 8) <> "Not Used" Then
      ' Update the MDI form's menu.
      frmMDI.mnuRecentFile(0).Visible = True
      frmMDI.mnuRecentFile(i).Caption = IniString
      frmMDI.mnuRecentFile(i).Visible = True
  
      ' Iterate through all the notepads and update each menu.
      For j = 1 To UBound(Document)
        If Not FState(j).Deleted Then
          Document(j).mnuRecentFile(0).Visible = True
          Document(j).mnuRecentFile(i).Caption = IniString
          Document(j).mnuRecentFile(i).Visible = True
        End If
      Next j
    End If
  Next i

End Sub

Sub OptionsToolbarProc (CurrentForm As Form)
    CurrentForm.mnuOToolbar.Checked = Not CurrentForm.mnuOToolbar.Checked
    If TypeOf CurrentForm Is MDIForm Then
    Else
        frmMDI.mnuOToolbar.Checked = CurrentForm.mnuOToolbar.Checked
    End If
    If CurrentForm.mnuOToolbar.Checked Then
        frmMDI.picToolbar.Visible = True
    Else
        frmMDI.picToolbar.Visible = False
    End If
End Sub

Sub WriteRecentFiles (OpenFileName)
  Dim i, j, key, retval
  Dim IniString As String
  IniString = String(255, 0)

  ' Copy RecentFile1 to RecentFile2, etc.
  For i = 3 To 1 Step -1
    key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", key, "Not Used", IniString, Len(IniString), "MDINOTE.INI")
    If retval And Left(IniString, 8) <> "Not Used" Then
      key = "RecentFile" & (i + 1)
      retval = WritePrivateProfileString("Recent Files", key, IniString, "MDINOTE.INI")
    End If
  Next i
  
  ' Write openfile to first Recent File.
    retval = WritePrivateProfileString("Recent Files", "RecentFile1", OpenFileName, "MDINOTE.INI")

End Sub

