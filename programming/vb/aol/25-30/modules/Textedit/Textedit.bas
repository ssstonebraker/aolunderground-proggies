Dim ArrayNum As Integer ' Index value for the menu control array mnuFileArray.
Global Filename As String ' This variable keeps track of the filename information for opening and closing files.
Const MB_YESNO = 4, MB_ICONQUESTION = 32, IDNO = 7, MB_DEFBUTTON2 = 256

Sub CloseFile (Filename As String)
Dim F As Integer
On Error GoTo CloseError                ' If there is an error, display the error message below.
    
    If Dir(Filename) <> "" Then         ' File already exists, so ask if overwriting is desired.
        response = MsgBox("Overwrite existing file?", MB_YESNO + MB_QUESTION + MB_DEFBUTTON2)
        If response = IDNO Then Exit Sub
    End If
    F = FreeFile
    Open Filename For Output As F       ' Otherwise, open the file name for output.
    Print #F, frmEditor!txtEdit.Text    ' Print the current text to the opened file.
    Close F                             ' Close the file
    Filename = "Untitled" ' Reset the caption of the main form
    Exit Sub
CloseError:
    MsgBox "Error occurred trying to close file, please retry.", 48
    Exit Sub
End Sub

Sub DoUnLoadPreCheck (unloadmode As Integer)
    If unloadmode = 0 Or unloadmode = 3 Then
            Unload frmAbout
            Unload frmEditor
            End
    End If
End Sub

Sub OpenFile (Filename As String)
Dim F As Integer
    If "Text Editor: " + Filename = frmEditor.Caption Then  ' Avoid opening the file if already loaded.
        Exit Sub
    Else
        On Error GoTo ErrHandler
            F = FreeFile
            Open Filename For Input As F                    ' Open file selected on File Open About.
            frmEditor!txtEdit.Text = Input$(LOF(F), F)
            Close F                                         ' Close file.
            frmEditor!mnuFileItem(3).Enabled = True         ' Enable the Close menu item
            UpdateMenu
            frmEditor.Caption = "Text Editor: " + Filename
            Exit Sub
    End If
ErrHandler:
        MsgBox "Error encountered while trying to open file, please retry.", 48, "Text Editor"
        Close F
        Exit Sub
End Sub

Sub UpdateMenu ()
    frmEditor!mnuFileArray(0).Visible = True            ' Make the initial element visible / display separator bar.
    ArrayNum = ArrayNum + 1                             ' Increment index property of control array.
    ' Check to see if Filename is already on menu list.
    For i = 0 To ArrayNum - 1
        If frmEditor!mnuFileArray(i).Caption = Filename Then
            ArrayNum = ArrayNum - 1
            Exit Sub
        End If
    Next i
    
    ' If filename is not on menu list, add menu item.
    Load frmEditor!mnuFileArray(ArrayNum)               ' Create a new menu control.
    frmEditor!mnuFileArray(ArrayNum).Caption = Filename ' Set the caption of the new menu item.
    frmEditor!mnuFileArray(ArrayNum).Visible = True     ' Make the new menu item visible.
End Sub

