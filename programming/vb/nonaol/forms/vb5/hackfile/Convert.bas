Attribute VB_Name = "Convert"
'Convert Example bY stUccO
'***************************************************************
'Windows API/Global Declarations for :Split any file into smaller
'     files
'***************************************************************
'*************************************
'*** PASTE THIS CODE INTO A MODULE ***
'*************************************
Option Explicit


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type


Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


Public Type SHITEMID
    mkidcb As Long
    abID As Byte
    End Type


Public Type ITEMIDLIST
    idlmkid As SHITEMID
    End Type


Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
    End Type


Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
    Public Const BIF_RETURNONLYFSDIRS = &H1


Function GetOpenFileNameDLG(Filter As String, Title As String, DefaultExt As String, WindowHnd As Long) As String


    On Error GoTo handelopenfile
    Dim OpenFile As OPENFILENAME, Tempstr As String
    Dim Success As Long, FileTitleLength%
    Filter = Find_And_Replace(Filter, "|", Chr(0))
    If Right$(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = WindowHnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = Filter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrTitle = Title
    OpenFile.lpstrDefExt = DefaultExt
    OpenFile.flags = 0
    Success = GetOpenFileName(OpenFile)


    If Success = 0 Then
        GetOpenFileNameDLG = ""
    Else
        Tempstr = OpenFile.lpstrFile
        GetOpenFileNameDLG = Trim(Tempstr)
    End If


    Exit Function
handelopenfile:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Function
End Function



Function Find_And_Replace(ByRef TextLine As String, ByRef SourceStr As String, ByRef ReplaceStr As String) As String


    On Error GoTo handelfindandreplace
    Dim DoAnother As Boolean, PosFound As Integer, ReturnStr As String


    DoAnother = True
        ReturnStr = TextLine


        While DoAnother
            PosFound = InStr(1, ReturnStr, SourceStr)


            If PosFound > 0 Then
                ReturnStr = Mid$(ReturnStr, 1, PosFound - 1) & ReplaceStr & Mid$(ReturnStr, PosFound + Len(SourceStr))
            Else


                DoAnother = False
                End If


            Wend


            Find_And_Replace = ReturnStr
handelfindandreplace:
            Exit Function
        End Function



