Attribute VB_Name = "FoRBiDon"


'sup, thanks for using mt sample pwsd
'it was made be me, ÐøøM, and also by Stock
'it is very simple to scan a file
'Example [Call Scan_For(text1.text,"Searchstring")
'Belove i have listed a list of searchstrings
'you can search for,.




'|¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'|SearchStrings:
'|
'|main.idx
'|win.ini
'|@juno.com
'|@hotmail.com
'|@freemail.com
'|@msn.com
'|@FreeYellow.com
'|win.com
'|autoexec.bat
'|_________________


' -ÐøøM
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
       Public Const LB_DIR = &H18D
       Public Const DDL_READWRITE = &H0
       Public Const DDL_READONLY = &H1
       Public Const DDL_HIDDEN = &H2
       Public Const DDL_SYSTEM = &H4
       Public Const DDL_DIRECTORY = &H10
       Public Const DDL_ARCHIVE = &H20
       Public Const DDL_DRIVES = &H4000
       Public Const DDL_EXCLUSIVE = &H8000
       Public Const DDL_POSTMSGS = &H2000
       Public Const DDL_FLAGS = DDL_ARCHIVE Or DDL_DIRECTORY
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Sub AddFilesToList(list1 As ListBox)

       Dim r As Long
       Dim pathSpec As String
       '     'fill the listbox yeee haah lol -FoRBiDon
       pathSpec = "c:\windows\system\*.*"
       r = SendMessageStr(list1.hWnd, LB_DIR, DDL_FLAGS, pathSpec)
End Sub
Sub File_Copy(File$, DestFile$)
If Not File_IfileExists(File$) Then Exit Sub
FileCopy File$, DestFile$
End Sub
Sub File_Delete(File$)
Dim NoFreeze%
If Not File_IfileExists(File$) Then Exit Sub
Kill File$
NoFreeze% = DoEvents()
End Sub
Function File_IfileExists(ByVal sFileName As String) As Integer
'Example: If Not File_ifileexists("win.com") then...
Dim i As Integer
On Error Resume Next
i = Len(dir$(sFileName))
    If Err Or i = 0 Then
        File_IfileExists = False
        Else
            File_IfileExists = True
    End If

End Function
Function Scan_Deltree(File As String)

'example : Call Scan_Deltree(text1.text)

Dim FileLenn As Variant
Dim FileLennn As Variant
Dim l003A As Variant
Dim l003E As Variant
Dim l0042 As String
Dim l0044 As Single
Dim l0046 As Single
Dim l0048 As Single
Dim l004a As Single
Dim l004c As Single
Dim l004e As Single
Dim l0050 As Single
Dim l0052 As Single
Dim l0054 As Single
Dim l0056 As Single
Dim l0058 As Single
Dim l005A As Variant
Dim l0045!
Open File For Binary As #2
DoEvents
FileLenn = LOF(2)
FileLennn = FileLenn
l003A = 1
While FileLennn >= 0
    If FileLennn > 32000 Then
        l003E = 32000
    ElseIf FileLennn = 0 Then
        l003E = 1
    Else
        l003E = FileLennn
    End If
    l0042$ = String$(l003E, " ")
    Get #2, l003A, l0042$
    l0044! = InStr(1, l0042$, "deltree \y", 1)
    l0045! = InStr(1, l0042$, "MZÿ C:\*.*", 1)
If l0044! Then deltreescan = True
Close: Exit Function
'If 10044! Then MsgBox "Deltree = true", , "Forbidons PWS detector"
'If 10045! Then MsgBox "deltree = false", , "Forbidons PWS detector"

If Not l0044! Then deltreescan = False
Close: Exit Function
Wend
End Function


Function IfDirExists(TheDirectory)

Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(dir$(TheDirectory))
If Err Or Check = 0 Then
    IfDirExists = False
Else
    IfDirExists = True
End If
End Function
Function FreeProcess()
Dim DooM

Do: DoEvents
DooM = DooM + 1
If DooM = 50 Then Exit Do
Loop
End Function
Sub Directory_Delete(dir)
'This deletes a directory automatically from your HD
RmDir (dir)
End Sub
Sub Directory_Create(dir)

'Call Directory_Create("C:\NewDir")
MkDir dir
End Sub
Private Function Scan_For(sFile$, ByVal sWhat$)
Dim VariantA As Variant
Dim VariantB As Variant
Dim VariantC As Variant
Dim VariantD As Variant
Dim SingleA As Single
Dim StringA As String
Dim EnterKey As String

On Error Resume Next
Open sFile$ For Binary As #1
    EnterKey$ = Chr$(13) + Chr$(10)
    msg$ = ""
    VariantA = LOF(1)
    VariantB = VariantA
    VariantC = 1

    If VariantB > 32000 Then
        VariantD = 32000
    ElseIf VariantB = 0 Then
        VariantD = 1
    Else
        VariantD = VariantB
    End If

    StringA$ = String$(VariantD, " ")
    Get #1, VariantC, StringA$

    SingleA! = InStr(1, StringA$, sWhat$, 1)

    If SingleA! Then
        Scan_For = 0
    Else
        Scan_For = 1
    End If
Close #1
End Function
Function ScanFile1(FileName$, SearchString As String, Label As Label) As Long
'ok, Filename is the File to scan
'SearchSring is the string to search for.
'and label like tells if it is a virus or not
'this is a example:
'call scanfile (text1.text,"Deltree y c:",label1)
'thats all
'i have also listed a list of virus searchstring below
'Main.idx, Deltree, Kill C:, Win.ini
'@Juno.com, @Hotmail.com, @FreeMail.com
'Deltree y c:, .Com, Deltree.com
'that is only a FEW

'-FoRBiDon
'
'
Dim free
Dim X
Dim Text$


free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #free
For X = 1 To LOF(free) Step 32000
    Text$ = Space(32000)
    Get #free, X, Text$
    Debug.Print X
    If InStr(1, Text$, SearchString$, 1) Then
    
MsgBox "Virus Found!"

        Where = InStr(1, Text$, SearchString$, 1)
        ScanFile1 = (Where + X) - 1
        Close #free
        Exit For
    End If
    Next X
    
    
    
    If Not InStr(1, Text$, SearchString$, 1) Then
MsgBox "No virus Found"

  End If
  
    
Close #free
End Function
Sub Directory_Delete2(dirnAmes$)
If Not IfDirExists(dirnAmes$) Then MsgBox dirnAmes$ & Chr(13) & "Bad Dir File Name!", 16, "Error": Exit Sub
On Error GoTo ErrorInDeletion
Kill dirnAmes$
Exit Sub
ErrorInDeletion:
MsgBox Error$
Resume Exitinga
Exitinga:
Exit Sub
End Sub
