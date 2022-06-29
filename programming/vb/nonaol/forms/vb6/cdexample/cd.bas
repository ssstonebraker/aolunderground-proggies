Attribute VB_Name = "CD"
'Module also made by me (JBF)
'Visit TZN PROGRAMMING
'Web: http://tznproggin.cjb.net
'Date: 11/27/00

'Returns False if the Save command has been canceled,
'True otherwise.
Function SaveTextControl(TB As Control, CD As CommonDialog, Filename As String) As Boolean
    Dim filenum As Integer
    On Error GoTo ExitNow
    
    CD.Filter = "All Files (*.*)|*.*|Text Files|*.txt"
    CD.FilterIndex = 2
    CD.DefaultExt = "txt"
    CD.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    CD.DialogTitle = "Select the destination file"
    CD.Filename = Filename
    'Exit if user presses Cancel.
    CD.CancelError = True
    CD.ShowSave
    Filename = CD.Filename
    
    'Write the control's contents.
    filenum = FreeFile()
    Open Filename For Output As #filenum
    Print #filenum, TB.Text;
    Close #filenum
    'Signal Success
    SaveTextControl = True
ExitNow:

End Function

'Returns False if the Save command has been canceled,
'True otherwise.
Function LoadTextControl(TB As Control, CD As CommonDialog, Filename As String) As Boolean
    Dim filenum As Integer
    On Error GoTo ExitNow
    
    CD.Filter = "All Files (*.*)|*.*|Text Files|*.txt"
    CD.FilterIndex = 2
    CD.DefaultExt = "txt"
    CD.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNNoReadOnlyReturn
    CD.DialogTitle = "Select the source file"
    CD.Filename = Filename
    'Exit if user presses Cancel.
    CD.CancelError = True
    CD.ShowOpen
    Filename = CD.Filename
    
    'Read the file's contents into the control.
    filenum = FreeFile()
    Open Filename For Input As #filenum
    TB.Text = Input$(LOF(filenum), filenum)
    Close #filenum
    'Signal Success
    LoadTextControl = True
ExitNow:

End Function

'Returns False if the Save command has been canceled,
'True otherwise.
Function SelectMultipleFiles(CD As CommonDialog, Filter As String, Filenames() As String) As Boolean
    On Error GoTo ExitNow
    
    CD.Filter = "All Files (*.*)|*.*" & Filter
    CD.FilterIndex = 1
    CD.Flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNExplorer
    CD.DialogTitle = "Select one or more files"
    CD.MaxFileSize = 10240
    CD.Filename = ""
    'Exit if user presses Cancel.
    CD.CancelError = True
    CD.ShowOpen
    
    'Parse the result to get filenames.
    Filenames() = Split(CD.Filename, vbNullChar)
    'Signal Success
    SelectMultipleFiles = True
ExitNow:

End Function
