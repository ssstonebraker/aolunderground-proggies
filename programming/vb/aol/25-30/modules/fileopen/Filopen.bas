
Sub FOpenProc ()
    Dim RetVal
    On Error Resume Next
    Dim OpenFileName As String
    frmMDI.CMDialog1.Filename = ""
    frmMDI.CMDialog1.Action = 1
    If Err <> 32755 Then 'user pressed cancel
	OpenFileName = frmMDI.CMDialog1.Filename
	OpenFile (OpenFileName)
	UpdateFileMenu (OpenFileName)
    End If
End Sub

Function GetFileName ()
    'Displays a Save As dialog and returns a file name
    'or an empty string if the user cancels
    On Error Resume Next
    frmMDI.CMDialog1.Filename = ""
    frmMDI.CMDialog1.Action = 2
    If Err <> 32755 Then      'User cancelled dialog
	GetFileName = frmMDI.CMDialog1.Filename
    Else
	GetFileName = ""
    End If
End Function

Function OnRecentFilesList (FileName) As Integer
  Dim i

  For i = 1 To 4
    If frmMDI.mnuRecentFile(i).Caption = FileName Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next i
    OnRecentFilesList = False
End Function

Sub OpenFile (FileName)
    Dim NL, TextIn, GetLine
    Dim fIndex As Integer

    NL = Chr$(13) + Chr$(10)
    
    On Error Resume Next
    ' open the selected file
    Open FileName For Input As #1
    If Err Then
	MsgBox "Can't open file: " + FileName
	Exit Sub
    End If
    ' change mousepointer to an hourglass
    screen.MousePointer = 11
    
    ' change form's caption and display new text
    fIndex = FindFreeIndex()
    document(fIndex).Tag = fIndex
    document(fIndex).Caption = UCase$(FileName)
    document(fIndex).Text1.Text = Input$(LOF(1), 1)
    FState(fIndex).Dirty = False
    document(fIndex).Show
    Close #1
    ' reset mouse pointer
    screen.MousePointer = 0
End Sub

Sub SaveFileAs (FileName)
On Error Resume Next
    Dim Contents As String

    ' open the file
    Open FileName For Output As #1
    ' put contents of the notepad into a variable
    Contents = frmMDI.ActiveForm.Text1.Text
    ' display hourglass
    screen.MousePointer = 11
    ' write variable contents to saved file
    Print #1, Contents
    Close #1
    ' reset the mousepointer
    screen.MousePointer = 0
    ' set the Notepad's caption

    If Err Then
	MsgBox Error, 48, App.Title
    Else
	frmMDI.ActiveForm.Caption = UCase$(FileName)
	' reset the dirty flag
	FState(frmMDI.ActiveForm.Tag).Dirty = False
    End If
End Sub

Sub UpdateFileMenu (FileName)
	Dim RetVal
	' Check if OpenFileName is already on MRU list.
	RetVal = OnRecentFilesList(FileName)
	If Not RetVal Then
	  ' Write OpenFileName to MDINOTEPAD.INI
	  WriteRecentFiles (FileName)
	End If
	' Update menus for most recent file list.
	GetRecentFiles
End Sub

