Attribute VB_Name = "Syn32"
'Syn32 v 0.25
'by martyr

'12-13-99


'*see Sub_About for updates*
'started 10-20-99

'this module does not use option explicit

'DECLARATIONS

'--finding windows
Public Declare Function FindChild& Lib "user32" Alias "FindWindowExA" (ByVal hWnd1&, ByVal hWnd2&, ByVal lpsz1$, ByVal lpsz2$)
Public Declare Function FindParent& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function IsWinVis& Lib "user32" Alias "IsWindowVisible" (ByVal hWnd&)
Public Declare Function GetWin& Lib "user32" Alias "GetWindow" (ByVal hWnd&, ByVal wCmd&)
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'--sending messages to windows
Public Declare Function PostIt& Lib "user32" Alias "PostMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lparam As Any)
Public Declare Function SendIt& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lparam As Any)
Public Declare Function SendItByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lparam&)
Public Declare Function SendItByString& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lparam$)
Public Declare Function SetWinPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'--hiding/showing windows
Public Declare Function ShowWin& Lib "user32" Alias "ShowWindow" (ByVal hWnd&, ByVal nCmdShow&)

'--getting text from windows
Public Declare Function WinTxt& Lib "user32" Alias "GetWindowTextA" (ByVal hWnd&, ByVal lpString$, ByVal cch&)
Public Declare Function WinTxtLen& Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd&)

'--for upchat
Public Declare Function EnableWin& Lib "user32" Alias "EnableWindow" (ByVal hWnd&, ByVal fEnable&)

'start taken from izekial32.bas
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const LB_GETITEMDATA = &H199
'end of taken from izekial32.bas

'CONSTANTS

'--left button on mouse
Public Const Down = &H201
Public Const Up = &H202
Public Const Click2 = &H203

'--window messages
Public Const SetTxt = &HC
Public Const GetTxt = &HD
Public Const GetTxtLen = &HE
Public Const Char = &H102
Public Const CloseWin = &H10
Public Const HideIt = &H0
Public Const Min = &H6
Public Const Norm = &H1
Public Const ShowIt = &H5
Public Const NextWin = &H2
Public Const FirstWin = &H0
Public Const WinTop = -(&H1)
Public Const NoMove = &H2
Public Const NoSize = &H1

'--checkbox options
Public Const GetCheck = &HF0

'--listbox options
Public Const SetIndex = &H186
Public Const GetCount = &H18B
Public Const GetLBTxt = &H189
Public Const GetLBTxtLen = &H18A

'--virtual keys
Public Const Retrn = &HD
Public Const Space = &H20
Public Const vkDown = &H28
Public Const vkUp = &H26

'TYPES

Type ChunkSize
    S12000 As String * 12000
    S6000 As String * 6000
    S3000 As String * 3000
    S1500 As String * 1500
    S500 As String * 500
    S100 As String * 100
    S25 As String * 25
    S5 As String * 5
    S1 As String * 1
End Type

'VARIABLES

Dim Bytes As ChunkSize
'

Public Sub AOL_ClearChat()
    Dim AoRoom As Long
    Dim AoTxt As Long
        AoRoom& = AOL_FindChat&
            If AoRoom& = 0& Then Exit Sub
            AoTxt& = Find_SubChild2&(AOL_FindChat&, "richcntl")
                Send_Text AoTxt&, ""
End Sub



Public Function File_EOLChars(ByVal File As String) As String
    On Local Error Resume Next
    Dim Free As Integer
    Dim BufferSize As Long
    Dim Buffer As String
    
    Free = FreeFile
        Open File For Input As #Free
            BufferSize = LOF(Free)
                If BufferSize > 1024 Then
                    Buffer = Input(1024, #Free)
                Else
                    Buffer = Input(BufferSize, #Free)
                End If
            If InStr(Buffer$, vbCrLf) Then
                File_EOLChars = vbCrLf
            ElseIf InStr(Buffer$, vbCr) Then
                File_EOLChars = vbCr
            ElseIf InStr(Buffer$, vbLf) Then
                File_EOLChars = vbLf
            Else
                File_EOLChars = vbNullString
            End If
        Close #Free

End Function

Public Function File_Exists(ByVal File As String) As Boolean
'want to know if a file exists?
'then use this.
    Dim FileNum As Integer
        On Error Resume Next
            FileNum = FreeFile
                Open File For Input As FileNum
                    If Err.Number <> 0& Then
                        File_Exists = False
                    Else
                        File_Exists = True
                    End If
                Close FileNum
End Function

Public Function File_LineCount(ByVal File As String) As Long
'this will get the line count of a text file
'the only real use for this i guess is if you're
'making a bot that loads from lists...
    On Local Error Resume Next
        Dim Lines As Long
        Dim intStart As Integer
        Dim intClose As Integer
        Dim Free As Integer
        Dim Buffer As String
        Dim Blocks As Long
        Dim Block As Long
        Dim LastBlockSize As Long
        Dim EOLChars As String
        Dim EOLCharsLen As Integer
        Dim FilePtr As Long
        Dim FileSize As Long
        
        Const BlockSize As Integer = 31477
        
        EOLChars = File_EOLChars(File)
        EOLCharsLen = Len(EOLChars)
        Buffer = Space(BlockSize)
        
        Free = FreeFile
        Open File For Input As #Free
            FileSize = LOF(Free)
            FilePtr = 1
            Blocks = (FileSize \ BlockSize) + 1
            For Block = 1 To Blocks
                If Block = Blocks Then
                    LastBlockSize = FileSize - Seek(Free) + 1
                    Buffer = Input(LastBlockSize, Free)
                Else
                    Buffer = Input(BlockSize, Free)
                End If
                DoEvents
                intStart = 1
                Do
                    intClose = InStr(intStart, Buffer, EOLChars)
                        If intClose Then
                            Lines = Lines + 1
                            intStart = intClose + EOLCharsLen
                        Else
                            FilePtr = FilePtr + intStart - 1
                            Seek #Free, FilePtr
                        End If
                Loop While intClose
            Next
        Close #Free
    File_LineCount = Lines

End Function


Public Sub File_Move(ByVal StartWhere As String, EndWhere As String)
'okay this will move startwhere to endwhere
    If File_Exists(StartWhere) = True Then
        FileCopy StartWhere, EndWhere
        Kill StartWhere
    End If
End Sub


Public Function Files_Merge(ByVal File As String, ByVal FileType As String, Optional NumOfSegs As Integer) As Integer
'this function will merge files back together ;c)
    Dim TotalBytes As Long
    Dim DestinationFile As String
    Dim SegFile As String
    Dim SegNum As Integer
    Dim Segs As Integer
    Dim BytesDone As Long
    Dim filename As String
    Dim FilePath As String
    Dim FileNameNoExt As String
    Dim ErrorCode As Integer
    
    On Error GoTo ErrorHandler
        If Len(File) = 0 Or Len(Dir(File)) = 0 Then
            ErrorCode = 1
            GoTo ErrorHandler
        End If
    Do
        i = i + 1
        j = InStr(Len(File) - i, File, "\", vbTextCompare)
    Loop Until j > 0
        filename = Right$(File, Len(File) - j)
        FilePath = Left$(File, j)
        j = InStr(1, filename, ".", vbTextCompare)
    If j = 0 Then
        FileNameNoExt = filename
    Else
        FileNameNoExt = Left$(filename, j - 1)
    End If
        Do
            Segs = Segs + 1
            Select Case SegNum
                Case Is < 10
                    SegFile = FilePath & FileNameNoExt & ".00" & CStr(Segs)
                    MsgBox SegFile
                Case 10 To 99
                    SegFile = FilePath & FileNameNoExt & ".0" & CStr(Segs)
                Case 100 To 999
                    SegFile = FilePath & FileNameNoExt & "." & CStr(Segs)
            End Select
                If Len(Dir(SegFile)) = 0 Then Exit Do
                TotalBytes = TotalBytes + FileLen(SegFile)
        Loop
                Segs = Segs - 1
    If Segs = 0 Then
        ErrorCode = 2
        GoTo ErrorHandler
    End If
        If InStr(FileType, ".") <> 0 Then
            DestinationFile = FilePath & FileNameNoExt & FileType
        Else
            DestinationFile = FilePath & FileNameNoExt & "." & FileType
        End If
    If Len(Dir(DestinationFile)) <> 0 Then
        ErrorCode = 3
        GoTo ErrorHandler
    End If
        Open DestinationFile For Binary Access Write As #1 Len = 1
            Do
                SegNum = SegNum + 1
                        Select Case SegNum
                            Case Is < 10
                                File = FilePath & FileNameNoExt & ".00" & CStr(SegNum)
                            Case 10 To 99
                                File = FilePath & FileNameNoExt & ".0" & CStr(SegNum)
                            Case 100 To 999
                                File = FilePath & FileNameNoExt & "." & CStr(SegNum)
                        End Select
                    Open File For Binary Access Read As #2 Len = 1
                        RemainingBytes = FileLen(File)
       
                        Do
                             
                             Select Case RemainingBytes
                                 Case Is >= 12000
                                     Get #2, , Bytes.S12000
                                     Put #1, , Bytes.S12000
                                     RemainingBytes = RemainingBytes - 12000
                                     BytesDone = BytesDone + 12000
                                     DoEvents
                                 Case 6000 To 11999
                                     Get #2, , Bytes.S6000
                                     Put #1, , Bytes.S6000
                                     RemainingBytes = RemainingBytes - 6000
                                     BytesDone = BytesDone + 6000
                                     DoEvents
                                 Case 3000 To 5999
                                     Get #2, , Bytes.S3000
                                     Put #1, , Bytes.S3000
                                     RemainingBytes = RemainingBytes - 3000
                                     BytesDone = BytesDone + 3000
                                     DoEvents
                                 Case 1500 To 2999
                                     Get #2, , Bytes.S1500
                                     Put #1, , Bytes.S1500
                                     RemainingBytes = RemainingBytes - 1500
                                     BytesDone = BytesDone + 1500
                                     DoEvents
                                 Case 500 To 1499
                                     Get #2, , Bytes.S500
                                     Put #1, , Bytes.S500
                                     RemainingBytes = RemainingBytes - 500
                                     BytesDone = BytesDone + 500
                                     DoEvents
                                 Case 100 To 499
                                     Get #2, , Bytes.S100
                                     Put #1, , Bytes.S100
                                     RemainingBytes = RemainingBytes - 100
                                     BytesDone = BytesDone + 100
                                     DoEvents
                                 Case 25 To 99
                                     Get #2, , Bytes.S25
                                     Put #1, , Bytes.S25
                                     RemainingBytes = RemainingBytes - 25
                                     BytesDone = BytesDone + 25
                                     DoEvents
                                 Case 5 To 24
                                     Get #2, , Bytes.S5
                                     Put #1, , Bytes.S5
                                     RemainingBytes = RemainingBytes - 5
                                     BytesDone = BytesDone + 5
                                     DoEvents
                                 Case 1 To 4
                                     Get #2, , Bytes.S1
                                     Put #1, , Bytes.S1
                                     RemainingBytes = RemainingBytes - 1
                                     BytesDone = BytesDone + 1
                                     DoEvents
                                 Case Is = 0
                                     Close 2
                                     DoEvents
                                     Exit Do
                             End Select
                             
                             Percent = Int((BytesDone / TotalBytes) * 100)
                             DoEvents
                         Loop
        
            Loop Until SegNum = Segs
    Close 1
        NumOfSegs = Segs
        Files_Merge = 0
        Exit Function
ErrorHandler:
    Select Case ErrorCode
        Case Is = 0
            Reset
            Files_Merge = 4
        Case Else
            Files_Merge = ErrorCode
    End Select
    
    Exit Function

End Function

Public Sub About()

'electronic mail: kaosdemon2@hotmail.com
'instant message: i be martyr

'this module was written and tested in
'Visual Basic 6.0 Enterprise Edition

'and for anyone that reads this about section thanx.
'i'm running outta ideas so send anything in to me.
'no matter how stupid it may sound ;c)

'12-13-99 *update*
'okay i changed this update quite a few times now ;c)
'but i'm going to finally hand out my module in its
'very very early stages. i haven't really done much
'testing on it so i'm sure there's a million bugs in it.
'but thats okay because this is only .25 of what i want to
'have in it. i'm going to be adding zip and unzip features
'along with my File_Merge function in the next version of this module.
'and hopefully a whole lot more ;c)

'and on a side note:
'i know i've spoke out against joining groups numerous times
'i just joined AAA kuz magnet started it and i figure kewl,
'so no one's in charge, we just help each other. so i guess
'its not really an "aol group", its more of a "help each other
'learn programming group". and with that said i'm going to write
'a tutorial on how to make a chat scan. not the source, but an
'example. enjoy the early stages of my module,and please
'e-mail me with any problems!

'12-04-99 *update*
'well i'm done with this module for now
'so i'm just going to do basic error testing.

'11-13-99 *update*
'okay today i got my own computer so i'm
'going to work on this module more, so
'i'm hoping to make it full, with as many
'options as i want, for aol/aim only.
'   *any sub that does not have AIM_
'   before it will can be used for aol.
'none of the AIM stuff has been tested.
'i'm on of those kinda programmers that
'just writes down all the windows in a txt
'file and use that to make programs...
'so i have a txt file with all the aim wins
'if you ever want a copy, e-mail me
'aol wins are easy to remember ;c)
'aol's windows
'aol frame25 (main parent win)
'mdiclient (aol child, but parent to all aol children)
'aol child (every damn win)
'aol toolbar (child to aol, parent to other toolbar)
'_aol_toolbar (child to toolbar, parent to all those buttons you see in the toolbar)
'_aol_icon (aol's buttons)
'_aol_image (aol's images i guess, not sure)
'richcntl (on aol children, textboxes and such)


'11-11-99 *update*
'it turns out some of my subs did not work
'when finding aol's children, so i had to
'redo a couple of functions, Find_Chat&
'and Find_Child&.
'GetWin& was added to the declarations
'NextWin was added to the constants
'added a function from izekial32.bas
'added ignoring subs

'okay this module was intended only to be an example
'but i ended up having fun with it, so i made
'it for a full release, mainly AIM stuff. and some very
'basic aol stuff, half way through the module i changed
'my code styling again, so some subs have "&" instead of long
'and stuff like that.

'how i got my handle:
'well my handle used to be KaosDemon2, but
'i decided that handle was getting lame kuz
'everyone was into the whole Kaos and Demon
'thing. so i stayed thinking towards evil
'and came up with Kain, the first person
'to kill anyone, according to the bible.
'i spelt the name different for effect but
'many people use that handle, so i said fuck
'this and went back to KaosDemon2 for a while,
'but i still wanted something that wasn't
'like everyone else. so i started listening
'to some music to kinda take my mind off shit.
'and i was listening to my favorite band,TOOL,
'and i heard "get off the fucking cross, we
'need the space for the next fool martyr"
'and i realized that no one in their right
'mind wants to be a martyr, so that was gonna
'be my handle.


'how i got into programming:
'well geez its been over a year now, and i
'still enjoy programming. the main influence
'on me was FBI '98, that was my favorite
'aol 3.0 proggie. and well from there
'i just started programming with vb4pe.
'lol i thought that modules made me the shit.



End Sub

Public Function File_GetType(ByVal File As String) As String
'this function will get the last characters of a filename
'this can be used with the File_Merge function to make sure
'you're making the correct file back ;c)
    Dim Find As Integer
    Dim FindTxt As String
        
    If Len(Dir(File)) = 0& Then: File_GetType$ = "": Exit Function
        Find% = Get_ChrCount(File, ".")
        If Find$ = 1 Then
            File_GetType$ = Mid(File$, InStr(File, "."))
        Else
            While Find% > 1
                FindTxt$ = Mid(File, InStr(File, "."))
                Find% = Get_ChrCount(FindTxt$, ".")
            Wend
            File_GetType$ = Mid(FindTxt$, InStr(File, "."))
End Function

Public Function File_Split(ByVal File As String, ByVal SegSize As Long, Optional ByVal NumOfSegs As Integer) As Integer
'okay i'm just starting to understand this more
'so i have the File_Split down and i'm currently
'working on the merge file.
'but i need to find a way to put "tags" in the file.
'like the mp3's do. so that the file can be re-built
'without knowing what the file was before hand.

'wait a little while for the merge function.
'but i've tested this numerous times, and it will
'tell you if there was an error or not by the value
'it returns.

'ex:
'   Dim RetVal as Integer
'   Dim Segz as Integer
'
'   RetVal = File_Split(somefile,100,ProgressBar1,Segz)
'   If RetVal = 0 then
'       MsgBox "File was split successfully into " & Segz & " files."
'   Else
'       MsgBox "An error has occurred"
'   End If

'so if you find a way to successfully merge the files
'back together and make the original file that was
'taken apart, e-mail me!
'i can do it w/ ini, but i don't want to do it that way

'this function will NOT split the file into 1000+ diff segments, because errors could occur
    Dim SourceBytes As Long
    Dim SourceFile As String
    Dim Destination As String
    Dim SegNum As Integer
    Dim RemainingBytes As Long
    Dim BytesDone As Long
    Dim FilePath As String
    Dim filename As String
    Dim FileNameNoExt As String
    Dim ErrorCode As Integer
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrorHandler
        If Len(File) = 0& Or Len(Dir(File)) = 0& Then
            ErrorCode = 1
            GoTo ErrorHandler
        End If
    If SegSize = 0 Then
        ErrorCode = 2
        GoTo ErrorHandler
    End If
        Do
            i = i + 1
            j = InStr(Len(File) - i, File, "\", vbTextCompare)
        Loop Until j > 0
    filename$ = Right$(File, Len(File) - j)
    FilePath$ = Left$(File, j)
    j = InStr(1, filename$, ".", vbTextCompare)
        If j = 0 Then
            FileNameNoExt = filename
        Else
            FileNameNoExt = Left(filename, j - 1)
        End If
    SourceBytes = FileLen(File)
        If SourceBytes / SegSize >= 1000 Then
            ErrorCode = 3
            GoTo ErrorHandler
        End If
    Open File For Binary Access Read As #1 Len = 1
        Do
            SegNum = SegNum + 1
                Select Case SegNum
                    Case Is < 10
                        Destination = FilePath & FileNameNoExt & ".00" & CStr(SegNum)
                    Case 10 To 99
                        Destination = FilePath & FileNameNoExt & ".0" & CStr(SegNum)
                    Case 100 To 999
                        Destination = FilePath & FileNameNoExt & "." & CStr(SegNum)
                End Select
                    Open Destination For Binary Access Write As #2 Len = 1
                        If SourceBytes - BytesDone < SegSize Then
                            RemainingBytes = SourceBytes - BytesDone
                        Else
                            RemainingBytes = SourceBytes
                        End If
                            Do
                                Select Case RemainingBytes
                                    Case Is >= 12000
                                        Get #1, , Bytes.S12000
                                        Put #2, , Bytes.S12000
                                        RemainingBytes = RemainingBytes - 12000
                                        BytesDone = BytesDone + 12000
                                        DoEvents
                                    Case 6000 To 11999
                                        Get #1, , Bytes.S6000
                                        Put #2, , Bytes.S6000
                                        RemainingBytes = RemainingBytes - 6000
                                        BytesDone = BytesDone + 6000
                                        DoEvents
                                    Case 3000 To 5999
                                        Get #1, , Bytes.S3000
                                        Put #2, , Bytes.S3000
                                        RemainingBytes = RemainingBytes - 3000
                                        BytesDone = BytesDone + 3000
                                        DoEvents
                                    Case 1500 To 2999
                                        Get #1, , Bytes.S1500
                                        Put #2, , Bytes.S1500
                                        RemainingBytes = RemainingBytes - 1500
                                        BytesDone = BytesDone + 1500
                                        DoEvents
                                    Case 500 To 1499
                                        Get #1, , Bytes.S500
                                        Put #2, , Bytes.S500
                                        RemainingBytes = RemainingBytes - 500
                                        BytesDone = BytesDone + 500
                                        DoEvents
                                    Case 100 To 499
                                        Get #1, , Bytes.S100
                                        Put #2, , Bytes.S100
                                        RemainingBytes = RemainingBytes - 100
                                        BytesDone = BytesDone + 100
                                        DoEvents
                                    Case 25 To 99
                                        Get #1, , Bytes.S25
                                        Put #2, , Bytes.S25
                                        RemainingBytes = RemainingBytes - 25
                                        BytesDone = BytesDone + 25
                                        DoEvents
                                    Case 5 To 24
                                        Get #1, , Bytes.S5
                                        Put #2, , Bytes.S5
                                        RemainingBytes = RemainingBytes - 5
                                        BytesDone = BytesDone + 5
                                        DoEvents
                                    Case 1 To 4
                                        Get #1, , Bytes.S1
                                        Put #2, , Bytes.S1
                                        RemainingBytes = RemainingBytes - 1
                                        BytesDone = BytesDone + 1
                                        DoEvents
                                    Case Is = 0
                                        Close 2
                                        DoEvents
                                        Exit Do
                                End Select
                                    Percent = Int((BytesDone - SourceBytes) * 100)
                                    DoEvents
                            Loop
            Exit Do
        Loop
    Close 1
        NumOfSegs = SegNum
        File_Split = 0
        Exit Function

ErrorHandler:
    Select Case ErrorCode
        Case Is = 0
            Reset
            File_Split = 4
        Case Else
            File_Split = ErrorCode
    End Select
Exit Function

End Function

Public Sub Form_OnTop(ByVal hWnd As Long)
'okay this is a pretty simple sub
'ex:
'   SetWinPos Main.hWnd

    If hWnd& = 0& Then Exit Sub
        SetWinPos hWnd, WinTop, 0, 0, 0, 0, NoMove Or NoSize
End Sub

Public Sub Mod_AddFunctions(ByVal Where As String, ByVal Cntrl As ComboBox)
'this will add all the functions from module to a combo box
'i was gonna make a program to compare modules, and maybe i'll add
'the subs and functions to this module ;c)
    Dim Txt As String
    Dim GetName As String
    Dim TempParen As Integer
    
    If Len(Dir$(Where$)) = 0& Then
        Cntrl.Text = "File does not exist"
        Exit Sub
    End If
    
    Open Where$ For Input As #1
    Do While Not EOF(1)
        Input #1, Txt$
            If InStr(Txt$, "Attribute VB_Name = ") Then
                GetName$ = Mid(Txt$, 22, Len(Txt$) - 22)
                Cntrl.Text = GetName$ & " - Functions"
                Exit Do
            End If
    Loop
    Close #1
    
    Open Where$ For Input As #1
    While Not EOF(1)
        Input #1, Txt$
            If InStr(Txt$, "Declare ") = 0& And InStr(Txt$, "End") = 0& And InStr(Txt$, "Exit") = 0& And InStr(Txt$, " Function ") <> 0& Then
                If InStr(Txt$, "Public") <> 0& Then
                    GetName$ = Mid(Txt$, 17)
                    TempParen% = InStr(GetName$, "(")
                    GetName$ = Mid(GetName$, 1, TempParen% - 1)
                ElseIf InStr(Txt$, "Private") <> 0& Then
                    GetName$ = Mid(Txt$, 18, Len(Txt$) - 2)
                    TempParen% = InStr(GetName$, "(")
                    GetName$ = Mid(GetName$, 1, TempParen% - 1)
                End If
                    Cntrl.AddItem GetName$
            End If
    Wend
    Close #1
End Sub


Public Sub AIM_AddBuddies(ByVal Obj As Control)
'this will add your buddies to a control
    Dim AIM&
    Dim AmTab&
    Dim AmTree&
    Dim Count&
    Dim Kount&
    Dim Sn$
    
    AIM& = FindParent&("_oscar_buddylistwin", vbNullString)
        If AIM& = 0& Then Exit Sub
        AmTab& = Find_SubChild2&(AIM&, "_oscar_tabgroup")
            AmTree& = Find_SubChild2&(AmTab&, "_oscar_tree")
                Count& = Get_ListCount&(AmTree&)
            For Kount& = 0 To Count&
                Sn$ = AIM_GetListText$(AmTree&, Kount&)
                    If Find_2&(Sn$, "(", ")") = 666& Then Call Obj.AddItem(Sn$)
            Next Kount&
End Sub

Public Sub AIM_AddChat(ByVal Obj As Control, Optional ByVal hWnd As Long)
'this will add the topmost chat to a control
    Dim AmChat&
    Dim AmTree&
    Dim Count&
    Dim Kount&
    Dim Sn$
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        If AmChat& = 0& Then Exit Sub
        AmTree& = Find_SubChild2&(AmChat&, "_oscar_tree")
            If hWnd& <> 0& Then AmTree& = hWnd&
        Count& = Get_ListCount&(AmTree&)
            For Kount& = 0 To Count&
                Sn$ = AIM_GetListText(Kount&)
                    If InStr(1, LCase(Sn$), LCase(AIM_UserSn$), vbTextCompare) = 0& Then Call Obj.AddItem(Sn$)
            Next Kount&
    'for those of you wondering about the "Optional ByVal hWnd as long"
    'that saves me some time with the AIM_AddChatAll sub ;c)
End Sub


Public Sub AIM_AddChatAll(ByVal Obj As Control)
'this will add all the open chats to a control
    Dim AmChat&
    Dim AmTree&
    Dim Count&
    Dim Kount&
    Dim Sn$
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        AmChat& = GetWin&(AmChat&, FirstWin)
        While AmChat& <> 0&
            AmTree& = Find_SubChild2&(AmChat&, "_oscar_tree")
                AIM_AddChat Obj, AmTree&
            AmChat& = GetWin&(AmChat&, NextWin)
        Wend
End Sub


Public Sub AIM_AddGroup(ByVal strgroup As String, ByVal Obj As Control)
'this will add a group from your buddy list
'to the control of your choice
    Dim AIM&
    Dim AmTab&
    Dim AmTree&
    Dim Count&
    Dim Kount&
    Dim Chount& 'really weird way to spell count ;)
    Dim Sn$
    
    AIM& = FindParent&("_oscar_buddylistwin", vbNullString)
        If AIM& = 0& Then Exit Sub
        AmTab& = Find_SubChild2&(AIM&, "_oscar_tabgroup")
            AmTree& = Find_SubChild2&(AmTab&, "_oscar_tree")
            Count& = Get_ListCount&(AmTree&)
                For Kount& = 0& To Count&
                    Sn$ = AIM_GetListText(AmTree&, Kount&)
                        If InStr(LCase(Sn$), LCase(strgroup)) <> 0& Then Exit For
                Next Kount&
                    If Kount& <> 0& And Kount& < Count& Then
                        For Chount& = (Kount& + 1) To Count&
                            Sn$ = AIM_GetListText(AmTree&, Chount&)
                                If Find_2&(Sn$, "(", ")") = 666 Then Exit For
                            Obj.AddItem Sn$
                        Next Chount&
End Sub


Public Sub AIM_ClearChat()
'this will clear the topmost chat
    Dim AmChat&
    Dim AmTxt&
    
        AmChat& = FindParent&("aim_chatwnd", vbNullString)
            If AmChat& = 0& Then Exit Sub
            AmTxt& = FindChild&(AmChat&, 0&, "wndate32class", vbNullString)
                Send_Text AmTxt&, ""
End Sub

Public Sub AIM_ClearChatAll()
'this will clear all the chats
    Dim AmChat&
    Dim AmTxt&
        
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        AmChat& = GetWin&(AmChat&, FirstWin)
            While AmChat& <> 0&
                AmTxt& = FindChild&(AmChat&, 0&, "wndate32class", vbNullString)
                    Send_Text AmTxt&, ""
                AmChat& = GetWin&(AmChat&, NextWin)
            Wend
End Sub


Public Sub AIM_Close()
'this will close aim
    Dim AIM&
    
    AIM& = FindParent&("_oscar_buddylistwin", vbNullString)
        If AIM& = 0& Then Exit Sub
        SendIt& AIM&, CloseWin, 0&, 0&
End Sub

Public Sub AIM_CloseChat()
'this will close the topmost chat
    Dim AmChat&
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        If AmChat& = 0& Then Exit Sub
        SendIt& , AmChat&, CloseWin, 0&, 0&
End Sub

Public Sub AIM_CloseChatAll()
'this will close all the chats
    Dim AmChat&
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        If AmChat& = 0& Then Exit Sub
        AmChat& = GetWin&(AmChat&, FirstWin)
        While AmChat& <> 0&
            SendIt& AmChat&, CloseWin, 0&, 0&
            AmChat& = FindParent&("aim_chatwnd", vbNullString)
        Wend
End Sub

Public Sub AIM_CloseChatByName(ByVal RoomName$)
'this will close a chat by name
    Dim AmChat&
    Dim ChatName$
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        AmChat& = GetWin&(AmChat&, FirstWin)
        If AmChat& = 0& Then Exit Sub
            While AmChat& <> 0&
                ChatName$ = Get_Caption$(AmChat&)
                If InStr(LCase(ChatName$), LCase(RoomName$)) <> 0& Then Call SendIt&(AmChat&, CloseWin, 0&, 0&): Exit Sub
                AmChat& = GetWin&(AmChat&, NextWin)
            Wend
End Sub

Public Sub AIM_CloseIM()
'this will close the topmost im
    Dim AmIM&
    
    AmIM& = FindParent&("aim_imessage", vbNullString)
        If AmIM& = 0& Then Exit Sub
        SendIt& AmIM&, CloseWin, 0&, 0&
End Sub



Public Sub AIM_CloseIMAll()
'this will close all the ims
    Dim AmIM&
    
    AmIM& = FindParent&("aim_imessage", vbNullString)
        If AmIM& = 0& Then Exit Sub
        AmIM& = GetWin&(AmIM&, FirstWin)
            While AmIM& <> 0&
                SendIt& AmIM&, CloseWin, 0&, 0&
                AmIM& = FindParent&("aim_imessage", vbNullString)
            Wend
End Sub


Public Sub AIM_CloseIMBySN(ByVal Sn$)
'this will close ain im by the sn
    Dim AmIM&
    Dim IMSn$
    
    AmIM& FindParent&("aim_imessage", vbNullString)
        If AmIM& = 0& Then Exit Sub
        AmIM& = GetWin&(AmIM&, FirstWin)
            While AmIM& <> 0&
                If LCase(AIM_GetIMsn$(AmIM&)) Like LCase(Sn$) Then Call SendIt&(AmIM&, CloseWin, 0&, 0&): Exit Sub
                AmIM& = GetWin&(AmIM&, NextWin)
            Wend
End Sub

Public Function AIM_FindChat(Optional ByVal RoomName As String) As Long
    Dim AmChat As Long
    Dim Name As String
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        If AmChat& = 0& Then AIM_FindChat& = 0&: Exit Function
            AmChat& = GetWin&(AmChat&, FirstWin)
        If Len(RoomName$) = 0& Then AIM_FindChat& = AmChat&: Exit Function
                While AmChat& <> 0&
                    Name$ = LCase(Get_Caption(AmChat&))
                        If InStr(Name$, LCase(RoomName$)) <> 0& Then AIM_FindChat& = AmChat&: Exit Function
                Wend
            AIM_FindChat& = 0&
End Function

Public Function AIM_FindIM&(Optional ByVal Sn$)
'this will find the topmost im
'or find an im by the persons sn
    Dim AmIM&
    Dim who$
    
    AmIM& = FindParent&("aim_imessage", vbNullString)
        If AmIM& = 0& Then AIM_FindIM& = 0&: Exit Function
        If Len(Sn$) = 0& Then AIM_FindIM& = AmIM&: Exit Function
                While AmIM& <> 0&
                    who$ = LCase(Get_Caption(AmIM&))
                        If InStr(who$, LCase(Sn$)) <> 0& Then AIM_FindIM& = AmIM&: Exit Function
                        AmIM& = FindChild&(0&, AmIM&, "aim_imessage", vbNullString)
                Wend
            AIM_FindIM& = 0&
End Function

Public Function AIM_GetIMsn$(Optional ByVal IM&)
'this was written to work with the new aim 3.5
'it will get the sn from a direct im
'gets the sn from the topmost im
'or form an im of your choice
    Dim AmIM&
    Dim IMSn$
    Dim Hyphen&
    
    AmIM& = AIM_FindIM&
    If IM& <> 0& Then AmIM& = IM&
        If AmIM& = 0& Then AIM_GetIMsn$ = "": Exit Function
        IMSn$ = Get_Caption$(AmIM&)
        Hyphen& = InStr(IMSn$, " - ")
            If Hyphen = 0& Then AIM_GetIMsn$ = "": Exit Function
        AIM_GetIMsn$ = Left$(IMSn$, Hyphen - 1)
End Function

Public Function AIM_GetListText$(ByVal LstHwnd&, ByVal index&)
'this will get text from an aim listbox
    Dim hWndTxtLen&
    Dim NullString$
    Dim TempTab&
    Dim Sn$
    
    SendItByNum& LstHwnd&, SetIndex, index&, 0&
    hWndTxtLen& = SendIt&(LstHwnd&, GetLBTxtLen, index&, 0&)
    NullString$ = String$(hWndTxtLen&, 0)
    SendItByString& LstHwnd&, GetLBTxt, index&, NullString
    TempTab& = InStr(NullString$, Chr(9))
    Sn$ = Right$(NullString$, Len(NullString$) - TempTab&)
    TempTab& = InStr(Sn$, Chr(9))
    Sn$ = Right$(Sn$, Len(Sn$) - TempTab&)
    AIM_GetListText$ = Sn$
End Function

Public Sub AIM_GetOpenIMsns(ByVal lst As Control)
'gets all your open im buddies and puts
'them in a listbox
    Dim AmIM&
    Dim IMSn$
    
    AmIM& = AIM_FindIM&
        While AmIM& <> 0&
            IMSn$ = AIM_GetIMsn$(AmIM&)
            lst.AddItem IMSn$
            AmIM& = GetWin&(AmIM&, NextWin)
        Wend
End Sub

Public Sub AIM_GoPage(ByVal Link$)
'goes to the page that you want
    Dim AIM&
    Dim AmEdit&
    Dim AmIcon&
    
    AIM& = FindParent&("_oscar_buddylistwin", vbNullString)
        If AIM& = 0& Then Exit Sub
            AmEdit& = FindChild&(AIM&, 0&, "edit", vbNullString)
                Send_Text AmEdit&, Link$
            AmIcon& = FindChild&(AIM&, 0&, "_oscar_iconbtn", vbNullString)
                Send_Click AmIcon&
End Sub

Public Sub AIM_IgnoreChatByIndex(ByVal index As Long, Optional ByVal RoomName As String)
    Dim AmChat As Long
    Dim AmTree As Long
    Dim AmButton As Long
    
    If Len(RoomName$) = 0& Then AmChat& = AIM_FindChat&
    If Len(RoomName$) <> 0& Then AmChat& = AIM_FindChat&(RoomName$)
        If AmChat& = 0& Then Exit Sub
        AmTree& = Find_SubChild2&(AmChat&, "_oscar_tree")
            SendIt& AmTree&, SetIndex, index&, 0&
        AmButton& = Find_SubChild2&(AmChat&, "_oscar_iconbtn", 2&)
            Send_Click AmButton&
End Sub

Public Sub AIM_IgnoreChatBySn(ByVal Sn As String, Optional ByVal RoomName As String)
    Dim AmChat As Long
    Dim AmTree As Long
    
End Sub


Public Sub AIM_SendChat(ByVal msg$)
'sends text to the chat room
    Dim AmChat&
    Dim AmTxt&
    Dim AmIcon&
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        If AmChat& = 0& Then Exit Sub
        AmTxt& = Find_SubChild2&(AmChat&, "wndate32class", 2)
        Send_Text AmTxt&, msg$
        AmIcon& = Find_SubChild2&(AmChat&, "_oscar_iconbtn", 4)
        Send_Click AmIcon&
End Sub

Public Sub AIM_SendChatAll(ByVal msg$)
'this will send the Msg to all the chats
    Dim AmChat&
    Dim AmTxt&
    Dim AmIcon&
    
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        If AmChat& = 0& Then Exit Sub
        AmChat& = GetWin&(AmChat&, FirstWin)
            While AmChat& <> 0&
                AmTxt& = Find_SubChild2&(AmChat&, "wndate32class", 2)
                Send_Text AmTxt&, msg$
                AmIcon& = Find_SubChild2&(AmChat&, "_oscar_iconbtn", 4)
                Send_Click AmIcon&
                AmChat& = GetWin&(AmChat&, NextWin)
            Wend
End Sub


Public Sub AIM_SendChatByName(ByVal ChatName$, ByVal msg$)
'this will send Msg to the Chat by its name
    Dim AmChat&
    Dim AmTxt&
    Dim AmIcon&
    Dim Where$
        
    AmChat& = FindParent&("aim_chatwnd", vbNullString)
        If AmChat& = 0& Then Exit Sub
        AmChat& = GetWin&(AmChat&, FirstWin)
            While AmChat& <> 0&
                Where$ = LCase(Get_Caption(AmChat&))
                If InStr(Where$, LCase(ChatName$)) <> 0& Then
                    AmTxt& = Find_SubChild2&(AmChat&, "wndate32class", 2)
                    Send_Text AmTxt&, msg$
                    AmIcon& = Find_SubChild2&(AmChat&, "_oscar_iconbtn", 4)
                    Send_Click AmIcon&
                    AmChat& = GetWin&(AmChat&, NextWin)
                End If
            Wend
End Sub


Public Sub AIM_SendIM(ByVal Sn$, ByVal msg$)
'this will open an im and send a message to that person
    Dim AIM&
    Dim AmIM&
    Dim AmButton&
    
    AIM& = FindParent&("_oscar_buddylistwin", vbNullString)
        If AIM& = 0& Then Exit Sub
    AIM_GoPage "aim:goim?screenname=" & Replace_String$(Sn$, " ", "+") & "&message=" & Replace_String$(msg$, " ", "+")
    AmIM& = AIM_FindIM&(Sn$)
        AmButton& = Find_SubChild2&(AmIM&, "_oscar_iconbtn")
        Send_Click AmButton&
End Sub

Public Sub AIM_SendIMAll(ByVal msg$)
'this will cycle through all the open
'imz and send the same message to all
'of them
    Dim AmIM&
    Dim AmTxt&
    Dim AmButton&
    
    AmIM& = FindParent&("aim_imessage", vbNullString)
        If AmIM& = 0& Then Exit Sub
        AmIM& = GetWin&(AmIM&, FirstWin)
            While AmIM& <> 0&
                AmTxt& = Find_SubChild2&(AmIM&, "wndate32class", 2)
                    Send_Text AmTxt&, msg$
                AmButton& = Find_SubChild2&(AmIM&, "_oscar_iconbtn")
                    Send_Click AmButton&
                AmIM& = GetWin&(AmIM&, NextWin)
            Wend
End Sub


Public Sub AIM_SendIMBySn(ByVal Sn$, ByVal msg$)
'this sends the message to an open im by
'the persons sn
    Dim AmIM&
    Dim AmTxt&
    Dim AmButton&
    
    AmIM& = AIM_FindIM&(Sn$)
        If AmIM& = 0& Then Exit Sub
        AmTxt& = Find_SubChild2&(AmIM&, "wndate32class", 2&)
            Send_Text AmTxt&, msg$
        AmButton& = Find_SubChild2&(AmIM&, "_oscar_iconbtn")
            Send_Click AmButton&
End Sub

Public Sub AIM_SendIM2(ByVal Sn$, ByVal msg$)
'this is the most efficient way of sending an im ;c)
'it will send the im to the person no matter what
'unless they aren't online...
    Dim AmIM&
    
    AmIM& = AIM_FindIM&(Sn$)
        If AmIM& = 0& Then Call AIM_SendIM(Sn$, msg$): Exit Sub
    AIM_SendIMBySn Sn$, msg$
End Sub


Public Function AIM_UserSn$()
'this will get the user's sn
    Dim AIM&
    Dim Txt$
    Dim buddylist&
    
    AIM& = FindParent&("_oscar_buddylistwin", vbNullString)
        If AIM& = 0& Then AIM_UserSn$ = "": Exit Function
    Txt$ = Get_Caption$(AIM&)
    buddylist& = InStr(Txt$, "'s Buddy List")
    Txt$ = Left$(Txt$, buddylist& - 1)
    AIM_UserSn$ = Txt$
End Function

Public Function Find_1(ByVal Str As String, ByVal What As String) As Long
'this is used by me ;c)
'but it just finds text from a within a string
    Find_1& = InStr(Str, What)
End Function

Public Function Find_2(ByVal Str As String, ByVal What1 As String, ByVal What2 As String) As Long
'this is a bit more advanced than Find_1
'it will look for two diff strings
'which comes in handy when working with aim ;c)
    If InStr(Str$, What$) <> 0& And InStr(Str$, What2$) <> 0& Then Find_2& = 666: Exit Function
        Find_2& = 0&
End Function


Public Function AOL_FindChatTxt2&()
'this will find the text box where you type
    Dim AoChild&
    Dim AoTxt&
    
    AoChild& = Find_Chat&
        If AoRoom& = 0& Then Find_ChatTxt2& = 0&: Exit Function
        AoTxt& = Find_SubChild2&(AoChild&, "richcntl", 2&)
            AOL_FindChatTxt2& = AoTxt&
End Function

Public Function Get_ChrCount(ByVal Text As String, ByVal Chr2Count As String) As Long
'this will count the number of a specific
'character in a string
'i was messing with win captions when i wrote this ;c)
    Dim Txt As String
    Dim Space As Long
    Dim Count As Long
        
    Txt$ = Text$
    Space& = InStr(Txt$, Chr2Count$)
        If Space& = 0& Then Exit Function
            Count& = 1&
                Do While Space& <> 0&
                    Txt$ = Mid$(Txt$, Space& + 1&)
                    Space& = InStr(Txt$, Chr2Count$)
                        If Space& = 0& Then Exit Do
                    Count& = Count& + 1&
                Loop
            Get_ChrCount& = Count&
End Function

Public Function Get_ListCount(ByVal LstHwnd As Long) As Long
'counts the number of items in a listbox
    If LstHwnd& = 0& Then Get_ListCount& = 0&: Exit Function
    Get_ListCount& = Int(SendItByNum&(LstHwnd, GetCount, 0&, 0&) - 1)
End Function

Public Sub Main()

'this is used for testing a module ;c)
'and i would highly recommend having a program
'load form a module rather than a form
'but hey, thats your choice not mine

'and as far as programming goes, if you want
'to work on a module w/o loading a form
'this sub will load up when the module is in
'a project all alone.
Dim sez As Integer
a = Files_Merge("c:\test\izekial32.001", ".bas", sez)
    MsgBox a
    MsgBox sez
End Sub

Public Sub Mod_AddSubs(ByVal Where As String, ByVal Cntrl As ComboBox)
'this will add all the subs from module to a combo box
    Dim Txt As String
    Dim GetName As String
    Dim TempParen As Integer
    
    If Len(Dir$(Where$)) = 0& Then
        Cntrl.Text = "File does not exist"
        Exit Sub
    End If
    
    Open Where$ For Input As #1
    Do While Not EOF(1)
        Input #1, Txt$
            If InStr(Txt$, "Attribute VB_Name = ") Then
                GetName$ = Mid(Txt$, 22, Len(Txt$) - 22)
                Cntrl.Text = GetName$ & " - Subs"
                Exit Do
            End If
    Loop
    Close #1
    
    Open Where$ For Input As #1
    While Not EOF(1)
        Input #1, Txt$
            If InStr(Txt$, "Declare ") = 0& And InStr(Txt$, "End") = 0& And InStr(Txt$, "Exit") = 0& And InStr(Txt$, " Sub ") <> 0& Then
                If InStr(Txt$, "Public") <> 0& Then
                    GetName$ = Mid(Txt$, 12)
                    TempParen% = InStr(GetName$, "(")
                    GetName$ = Mid(GetName$, 1, TempParen% - 1)
                ElseIf InStr(Txt$, "Private") <> 0& Then
                    GetName$ = Mid(Txt$, 13, Len(Txt$) - 2)
                    TempParen% = InStr(GetName$, "(")
                    GetName$ = Mid(GetName$, 1, TempParen% - 1)
                End If
                    Cntrl.AddItem GetName$
            End If
    Wend
    Close #1
End Sub

Public Function Replace_String$(ByVal Txt$, ByVal FindStr$, ByVal NewStr$)
'okay this is a basic function
'it comes with vb6 but i know some people don't
'have vb6 sooo i use this instead of the Replace function
    Dim FindIt&
    Dim TextLen&
    
    FindIt& = 1
    TextLen& = Len(NewStr$)

        While FindIt& > 0&
            FindIt& = InStr(FindIt&, Txt$, FindStr$)
            If FindIt& > 0& Then
                Txt$ = Left(Txt$, FindIt& - 1) + NewStr$ + Mid(Txt$, FindIt& + Len(FindStr))
                FindIt& = FindIt& + TextLen&
            End If
        Wend
            Replace_String$ = Txt$
End Function

Public Function AOL_FindChat&()
'finds the aol chat room
    Dim AOL&
    Dim AoMDI&
    Dim AoChild&
    Dim AoList&
    Dim AoTxt&
    Dim AoTxt2&
    
    AOL& = FindParent&("aol frame25", vbNullString)
        If AOL& = 0& Then AOL_FindChat& = 0&: Exit Function
        AoMDI& = FindChild&(AOL&, 0&, "mdiclient", vbNullString)
            AoChild& = FindChild&(AoMDI&, 0&, "aol child", vbNullString)
            AoChild& = GetWin(AoChild&, FirstWin)
                AoList& = Find_SubChild2&(AoChild&, "_aol_listbox")
                AoTxt& = Find_SubChild2&(AoChild&, "richcntl")
                AoTxt2& = Find_SubChild2&(AoChild&, "richcntl", 2&)
                    If (AoList& <> 0&) And (AoTxt& <> 0&) And (AoTxt2& <> 0&) Then
                        AOL_FindChat& = AoChild&
                        Exit Function
                    Else
                        While AoChild& <> 0&
                            DoEvents
                            AoChild& = GetWin&(AoChild&, NextWin)
                                AoList& = Find_SubChild2&(AoChild&, "_aol_listbox")
                                AoTxt& = Find_SubChild2&(AoChild&, "richcntl")
                                AoTxt2& = Find_SubChild2&(AoChild&, "richcntl", 2&)
                            If (AoList& <> 0&) And (AoTxt& <> 0&) And (AoTxt2& <> 0&) Then
                                AOL_FindChat& = AoChild&
                                Exit Function
                            End If
                        Wend
                    End If
            AOL_FindChat& = 0&
End Function


Public Function AOL_FindChild&(Optional ByVal Title$)
'okay this function was originally done by SiR
'i rewrote the coding to my own, because
'i liked the function and how it saved time
    Dim AOL&
    Dim AoMDI&
    Dim AoChild&
    
    AOL& = FindParent&("aol frame25", vbNullString)
        If AOL& = 0& Then AOL_FindChild& = 0&: Exit Function
        AoMDI& = FindChild&(AOL&, 0&, "mdiclient", vbNullString)
            AoChild& = FindChild&(AoMDI&, 0&, "aol child", vbNullString)
            If Len(Title$) = 0& Then AOL_FindChild& = AoChild&: Exit Function
                If InStr(LCase$(Get_Caption$(AoChild&)), LCase$(Title$)) <> 0& Then AOL_FindChild& = AoChild&: Exit Function
            While AoChild& <> 0&
                DoEvents
                AoChild& = GetWin&(AoChild&, NextWin)
                    If InStr(LCase$(Get_Caption$(AoChild&)), LCase$(Title$)) <> 0& Then AOL_FindChild& = AoChild&: Exit Function
            Wend
    AOL_FindChild& = 0&
End Function

Public Function AOL_FindSendMail&()
'finds the aol child that sends mail
    AOL_FindSendMail& = Find_Child&("write mail")
End Function

Public Function Get_Caption$(ByVal hWnd&)
'gets the caption from a form
    Dim hWndTxtLen&
    Dim NullString$
        If hWnd& = 0& Then Get_Caption$ = "": Exit Function
        hWndTxtLen& = WinTxtLen&(hWnd&)
        NullString$ = String$(hWndTxtLen&, Chr$(0))
        WinTxt& hWnd&, NullString$, hWndTxtLen& + 1
        Get_Caption$ = NullString$
End Function

Public Function Get_ListTextByIndex$(ByVal lst&, ByVal index&)
    'this function is izekials coding
    'i just made it a lil diff
    'his coding though
    On Error Resume Next
    Dim rlist&
    Dim sthread&
    Dim mthread&
    Dim screenname$
    Dim itmhold&
    Dim psnHold&
    Dim rbytes&
    Dim cprocess&
    
    rlist& = lst&
    sthread& = GetWindowThreadProcessId&(rlist&, cprocess&)
    mthread& = OpenProcess&(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& <> 0& Then
                screenname$ = String$(4, vbNullChar)
            itmhold& = SendIt&(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
                ReadProcessMemory mthread&, itmhold&, screenname$, 4, rbytes&
                CopyMemory psnHold&, ByVal screenname$, 4
            psnHold& = psnHold& + 6
            screenname$ = String$(16, vbNullChar)
                ReadProcessMemory mthread&, psnHold&, screenname$, Len(screenname$), rbytes&
                screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            getlistitemtext$ = screenname$
                CloseHandle mthread&
        End If
End Function

Public Sub AOL_IgnoreChatSn(ByVal Sn$, Optional ByVal ignore As Boolean)
'ignore a person in the chat room
'by their full or partial sn
    Dim AoRoom&
    Dim AoRoomList&
    Dim Count&
    Dim Kount&
    Dim Person$

    AoRoom& = Find_Chat&
        If AoRoom& = 0& Then Exit Sub
    AoRoomList& = Find_SubChild2&(AoRoom&, "_aol_listbox")
        Count& = Get_ListCount&(AoRoomList&)
        For Kount& = 0 To Count&
            Person$ = Get_ListTextByIndex$(AoRoomList&, Kount&)
                If InStr(1, Sn$, Person$, 1) <> 0& And Sn$ <> UserSn$ Then Call AOL_IgnoreChatIndex(Kount&, ignore): Exit Sub
        Next Kount&
End Sub


Public Sub AOL_IgnoreIMSn(ByVal Sn$, ByVal ignore As Boolean)
'this will turn ims off/on to the sn of your choice
    Dim What$
        Select Case ignore
            Case True
                What$ = "$IM_On"
            Case False
                What$ = "$IM_Off"
        End Select
    AOL_SendIM Sn$ & " " & What$, " "
End Sub


Public Sub AOL_SendChat(ByVal msg$, Optional ByVal SaveText As Boolean)
'sends text to the chat window
    Dim AoRoom&
    Dim AoTxt&
    Dim SaveTxt$
        AoRoom& = AOL_FindChat&
            If AoRoom& = 0& Then Exit Sub
            AoTxt& = AOL_FindChatTxt2&
        If SaveText = True Then
            If Len(Get_Text$(AoTxt&)) <> 0& Then
                SaveTxt$ = Get_Text$(AoTxt&)
                Send_Text AoTxt&, ""
            Else
                SaveTxt$ = ""
            End If
        End If
        Send_Text AoTxt&, SaveTxt$
        Send_Enter AoTxt&
        If SaveText = True Then
            Send_Text AoTxt&, SaveTxt$
        End If
End Sub

Public Function Get_Text$(ByVal hWnd&)
'gets the text from a window
    Dim hWndTxtLen&
    Dim NullString$
        If hWnd& = 0& Then Get_Text$ = "": Exit Function
        hWndTxtLen& = SendIt&(hWnd&, GetTxtLen, 0&, 0&)
        NullString$ = String$(hWndTxtLen&, vbNullChar)
        SendIt& hWnd&, GetTxt, hWndTxtLen& + 1, NullString$
        Get_Text$ = NullString$
End Function

Public Sub AOL_MenuByChar(ByVal icon&, ByVal Character$)
'this will run the menu that drops
'from the icon of your choice on the
'AOL Toolbar
    
'this sub will not work for me....
    Dim AOL&
    Dim AoToolbar&
    Dim AoToolBar2&
    Dim AoIcon&
    Dim Mnu&
    Dim Mnu2&
    Dim MnuVis&
    Dim Count&
    
    AOL& = FindParent("aol frame25", vbNullString)
        If AOL& = 0& Then Exit Sub
        AoToolbar& = FindChild&(AOL&, 0&, "aol toolbar", vbNullString)
            AoToolBar2& = FindChild&(AoToolbar&, 0&, "_aol_toolbar", vbNullString)
                AoIcon& = Find_SubChild2&(AoToolBar2&, "_aol_icon", icon&)
            SendIt& AoIcon&, Down, 0&, 0&
            SendIt& AoIcon&, Up, 0&, 0&
                Mnu& = FindParent&("#32768", vbNullString)
                While MnuVis& = 0&
                    DoEvents
                    Mnu& = FindParent&("#32768", vbNullString)
                    MnuVis& = IsWinVis&(Mnu&)
                Wend
            PostIt& Mnu&, Char, Asc(Character$), 0&
End Sub

Public Sub Send_Click(ByVal hWnd&)
'sends a click to the deisired window
    If hWnd& = 0& Then Exit Sub
    SendIt& hWnd&, Down, 0&, 0&
    SendIt& hWnd&, Up, 0&, 0&
End Sub

Public Sub Send_Close(ByVal hWnd As Long)
    If hWnd& = 0& Then Exit Sub
    SendIt& hWnd&, CloseWin, 0&, 0&
End Sub

Public Sub AOL_SendKeyword(ByVal KW$)
'sends an AOL keyword
    Dim AOL&
    Dim AoToolbar&
    Dim AoToolBar2&
    Dim AoCombo&
    Dim AoEdit&
    
    AOL& = FindParent&("aol frame25", vbNullString)
        If AOL& = 0& Or Len(KW$) = 0& Then Exit Sub
        AoToolbar& = FindChild&(AOL&, 0&, "aol toolbar", vbNullString)
            AoToolBar2& = FindChild&(AoToolbar&, 0&, "_aol_toolbar", vbNullString)
                AoCombo& = FindChild&(AoToolBar2&, 0&, "_aol_combobox", vbNullString)
                    AoEdit& = FindChild&(AoCombo&, 0&, "edit", vbNullString)
                    Send_Text& AoEdit&, KW$
                    SendIt& AoEdit&, Char, Space, 0&
                    SendIt& AoEdit&, Char, Retrn, 0&
End Sub

Public Sub AOL_SendIM(ByVal Sn$, ByVal msg$)
'send an instant message to someone
    Dim AOL&
    Dim AoFindIM&
    Dim AoIM&
    Dim AoTxt&
    Dim AoIcon&
    Dim AoButton&
    Dim AoOk&
    
    AOL& = FindParent&("aol frame25", vbNullString)
        If AOL& = 0& Then Exit Sub
    AOL_SendKeyword "aol://9293:" & Sn$
    While AoFindIM& = 0&
        AoFindIM& = AOL_FindChild&("send instant message")
    Wend
    AoTxt& = Find_SubChild2&(AoFindIM&, "richcntl")
        Send_Text AoTxt&, msg$
    AoIcon& = Find_SubChild2&(AoFindIM&, "_aol_icon", 9)
        Send_Click AoIcon&
    While AoIM& = 0& Or AoOk& = 0&
        AoIM& = AOL_FindChild&("send instant message")
        AoOk& = FindParent&("#32770", "America Online")
    Wend
        If AoOk& <> 0& Then
            AoButton& = Find_SubChild2&(AoOk&, "button")
            Send_vkClick AoButton&
            Send_Close AoIM&
        End If
End Sub

Public Sub Send_vkClick(ByVal Button As Long)
'this is for clicking buttons
    If Button& = 0& Then Exit Sub
    PostIt& Button&, vkDown, Space, 0&
    PostIt& Button&, vkUp, Space, 0&
End Sub

Public Sub AOL_SendMail(ByVal Address$, ByVal subject$, ByVal msg$, Optional OffLine As Boolean = True)
'sends e-mail

'okay the offline thing is pretty neet to me
'if you select true then the mail will not send and close with an error
'if the user is not online
'if you select false then mail is sent online and if offline then
'the mail will just remain open and un-sent
'*true is the defualt setting
        Dim AOL&
        Dim AoMDI&
        Dim AoToolbar&
        Dim AoToolBar2&
        Dim AoIcon&
        Dim AoStatic&
        Dim AoChild&
        Dim AoEdit&
        Dim AoTxt&
        Dim Modal&
        
        AOL& = FindParent&("aol frame25", vbNullString)
                If AOL& = 0& Then Exit Sub
                AoToolbar& = FindChild&(AOL&, 0&, "aol toolbar", vbNullString)
                    AoToolBar2& = FindChild&(AoToolbar&, 0&, "_aol_toolbar", vbNullString)
                        AoIcon& = Find_SubChild2&(AoToolBar2&, "_aol_icon", 2&)
                        Send_Click AoIcon&
                        AoIcon& = 0
                            Do
                                AoChild& = Find_SendMail&
                                AoEdit& = Find_SubChild2&(AoChild&, "_aol_edit", 3&)
                                AoTxt& = FindChild&(AoChild&, 0&, "richcntl", vbNullString)
                                AoIcon& = Find_SubChild2&(AoChild&, "_aol_icon", 14&)
                            Loop Until AoChild& <> 0& And AoEdit& <> 0& And AoTxt& <> 0& And AoIcon& <> 0&
                        AoEdit& = Find_SubChild2&(AoChild&, "_aol_edit")
                            Send_Text AoEdit&, Address$
                        AoEdit& = Find_SubChild2&(AoChild&, "_aol_edit", 2&)
                            Send_Text AoEdit&, subject$
                        AoEdit& = Find_SubChild2&(AoChild&, "_aol_edit", 3&)
                            Send_Text AoEdit&, msg$
                            If Len(AOL_UserSn$) <> 0& Then
                                Send_Click AoIcon&
                                    AoIcon& = 0
                                Do
                                    Modal& = FindParent&("_aol_modal", vbNullString)
                                    AoIcon& = FindChild&(Modal&, 0&, "_aol_icon", vbNullString)
                                Loop Until Modal& <> 0& And AoIcon& <> 0&
                                    AoChild& = Find_Child&("write mail")
                                    If Modal& <> 0& And AoChild& <> 0& Then
                                        Send_Click& AoIcon&
                                        Exit Sub
                                    ElseIf AoChild& = 0& And Modal& = 0& Then
                                        Exit Sub
                                    End If
                            End If
                    If OffLine = True Then
                        MsgBox "Error, not currently signed on to America Online's server.", 64, "Error:101"
                        SendIt& Find_Child&("write mail"), CloseWin, 0&, 0&
                    End If
End Sub

Public Sub AOL_UpChat(ByVal onoff As Boolean)
'allows you to use aol while uploading
    Dim AOL&
    Dim Modal&
    
    Modal& = Find_Modal&
    AOL& = FindParent&("aol frame25", vbNullString)
    Select Case onoff
        Case True
            EnableWin Modal&, 0&
            ShowWin Modal&, Min
            EnableWin AOL&, 1&
        Case False
            EnableWin Modal&, 1&
            ShowWin Modal&, Norm
            EnableWin AOL&, 0&
    End Select

End Sub

Public Function AOL_FindModal&()
'finds the upload modal
    Dim Modal&
    Dim MoStatic&
    Dim MoStatic2&
    Dim MoCheck&
    Dim MoGauge&
    Dim MoGauge2&
    Dim MoButton&
    Dim MoButton2&
    
    Modal& = FindParent&("_aol_modal", vbNullString)
        If Modal& = 0& Then AOL_FindModal& = 0&: Exit Function
        MoStatic& = FindChild&(Modal&, 0&, "_aol_static", vbNullString)
        MoStatic2& = FindChild&(Modal&, MoStatic2&, "_aol_static", vbNullString)
        MoCheck& = FindChild&(Modal&, 0&, "_aol_checkbox", vbNullString)
        MoGauge& = FindChild&(Modal&, 0&, "_aol_gauge", vbNullString)
        MoGauge2& = FindChild&(Modal&, MoGauge&, "_aol_gauge", vbNullString)
        MoButton& = FindChild&(Modal&, 0&, "_aol_button", vbNullString)
        MoButton2& = FindChild&(Modal&, MoButton2&, "_aol_button", vbNullString)
        If MoStatic& <> 0& And MoStatic2& <> 0& And MoCheck& <> 0& And MoGauge& <> 0& And MoGauge2& <> 0& And MoButton& <> 0& And MoButton2& <> 0& Then AOL_FindModal& = Modal&: Exit Function
            While Modal& <> 0&
                Modal& = GetWin&(Modal&, NextWin)
                MoStatic& = FindChild&(Modal&, 0&, "_aol_static", vbNullString)
                MoStatic2& = FindChild&(Modal&, MoStatic2&, "_aol_static", vbNullString)
                MoCheck& = FindChild&(Modal&, 0&, "_aol_checkbox", vbNullString)
                MoGauge& = FindChild&(Modal&, 0&, "_aol_gauge", vbNullString)
                MoGauge2& = FindChild&(Modal&, MoGauge&, "_aol_gauge", vbNullString)
                MoButton& = FindChild&(Modal&, 0&, "_aol_button", vbNullString)
                MoButton2& = FindChild&(Modal&, MoButton2&, "_aol_button", vbNullString)
                If MoStatic& <> 0& And MoStatic2& <> 0& And MoCheck& <> 0& And MoGauge& <> 0& And MoGauge2& <> 0& And MoButton& <> 0& And MoButton2& <> 0& Then AOL_FindModal& = Modal&: Exit Function
            Wend
        AOL_FindModal& = 0&
End Function

Public Function Find_SubChild2&(ByVal child&, ByVal Class$, Optional ByVal SubNum& = 1)
'finds a child of an AOL child by
'using the AoChilds handle and childs
'class
'this function doesn't not have to be used w/ aol

    Dim AoChild&
    Dim AoSubChild&
    Dim AoSubClass$
    Dim Count&
    
    If child& = 0& Then Find_SubChild2& = 0&: Exit Function
    If Len(Class$) = 0& Then Find_SubChild2& = 0&: Exit Function
    If SubNum& > Find_MaxSubChildren&(child&, Class$) Then Find_SubChild2& = 0&: Exit Function
    AoChild& = child&
    AoSubClass$ = Class$
        AoSubChild& = FindChild&(child&, 0&, AoSubClass$, vbNullString)
            If SubNum& = 1 Then Find_SubChild2& = AoSubChild&: Exit Function
            For Count& = 2 To SubNum&
                AoSubChild& = FindChild&(child&, AoSubChild&, AoSubClass$, vbNullString)
            Next
            Find_SubChild2& = AoSubChild&
End Function

Public Function Find_SubChild&(ByVal ChildCaption$, ByVal Class$, Optional ByVal SubNum& = 1)
'finds a child from an AOL child window
'by the AoChilds caption and the class
'of the child.
    Dim AoChild&
    Dim AoSubChild&
    Dim AoSubClass$
    Dim Count&
    
    If Len(ChildCaption$) = 0& Then Find_SubChild& = 0&: Exit Function
    If Len(Class$) = 0& Then Find_SubChild& = 0&: Exit Function
    If Find_MaxSubChildren&(Find_Child&(ChildCaption$), Class$) < SubNum& Then Find_SubChild& = 0&: Exit Function
    AoChild& = Find_Child&(ChildCaption$)
        If AoChild& = 0& Then Find_SubChild& = 0&: Exit Function
    AoSubClass$ = Class$
        AoSubChild& = FindChild&(AoChild&, 0&, AoSubClass$, vbNullString)
            If SubNum& = 1 Then Find_SubChild& = AoSubChild&: Exit Function
            For Count& = 2 To SubNum&
                AoSubChild& = FindChild&(AoChild&, AoSubChild&, AoSubClass$, vbNullString)
            Next
            Find_SubChild& = AoSubChild&
End Function
Public Sub Send_Text(ByVal hWnd&, ByVal Txt$)
'sends text
    If hWnd& = 0& Then Exit Sub
    SendIt& hWnd&, SetTxt, 0&, Txt$
End Sub

Public Sub Send_Enter(ByVal hWnd&)
'sends the enter key to the window of
'your choice
    If hWnd& = 0& Then Exit Sub
    SendIt& hWnd&, Char, 13, 0&
End Sub

Public Function AOL_UserSn$()
'gets the sn from AOL
    Dim AOL&
    Dim AoWelcome&
    Dim AoWelCaption$
    
    AOL& = FindParent&("aol frame25", vbNullString)
        If AOL& = 0& Then Exit Function
    AoWelcome& = AOL_FindChild&("welcome")
    AoWelCaption$ = Get_Caption$(AoWelcome&)
    AoWelCaption$ = Right$(AoWelCaption$, Len(AoWelCaption$) - InStr(AoWelCaption$, ", "))
        If Len(AoWelCaption$) = 0& Then AOL_UserSn$ = "": Exit Function
    AoWelCaption$ = Mid$(AoWelCaption$, 1, Len(AoWelCaption$) - 1)
    AOL_UserSn$ = AoWelCaption$
End Function

Public Function Find_MaxSubChildren&(ByVal child&, ByVal Class$)
'designed to find the maximum number of
'children on an AOL child
    Dim ClassTest&
    Dim SubChild&
    Dim Count&
    
    If child& = 0& Then Find_MaxSubChildren& = 0&: Exit Function
    ClassTest& = FindChild&(child&, 0&, Class$, vbNullString)
    If ClassTest& = 0& Then Find_MaxSubChildren& = 0&: Exit Function
    SubChild& = FindChild&(child&, 0&, Class$, vbNullString)
    Count& = 1
        While SubChild& <> 0&
            SubChild& = FindChild&(child&, SubChild&, Class$, vbNullString)
            If SubChild& = 0& Then Find_MaxSubChildren& = Count&: Exit Function
            Count& = Count& + 1
        Wend
    Find_MaxSubChildren& = Count&
End Function

Public Sub ListBox_Load(ByVal Where$, ByVal lst As ListBox)
'thanx to COREAPI.hlp by anubis
    Dim Txt$
    
    If Len(Dir$(Where$)) = 0& Then Exit Sub
    Open Where$ For Input As #1
    While Not EOF(1)
        Input #1, Txt$
        lst.AddItem Txt$
    Wend
    Close #1
End Sub

Public Sub ListBox_Save(ByVal Where$, ByVal lst As ListBox)
'thanx to COREAPI.hlp by anubis
    Dim Count%
    
    Open Where$ For Output As #1
    For Count% = 0 To lst.ListCount - 1
        Print #1, thelist.List(Count%)
    Next Count%
    Close #1
End Sub

Public Sub Pause(ByVal StopTime&)
'thanx to wgf for this sub
    Dim Time&
    
    Time& = Timer
    While Timer - Time& >= StopTime&
        DoEvents
    Wend
End Sub

Public Sub AOL_IgnoreChatIndex(ByVal index&, Optional ByVal ignore As Boolean)
'ignore someone in the chat room
'by the index of their sn
'this sub is used in Ignore_ChatSn$
    Dim AoRoom&
    Dim AoRoomList&
    Dim Count&
    Dim Sn$
    Dim AoChild&
    Dim AoCheckbox&
    Dim SetCheckVal&
    Dim GetCheckVal&
    
    SetCheckVal& = 0&
    AoRoom& = Find_Chat&
        If AoRoom& = 0& Then Exit Sub
    AoRoomList& = FindChild&(AoRoom&, 0&, "_aol_listbox", vbNullString)
        Count& = SendIt&(AoRoomList&)
        If Int(index&) > Int(Count& - 1) Then Exit Sub
            SendIt& AoRoomList&, SetIndex, index&, 0&
            SendIt& AoRoomList&, Click2, 0&, 0&
        Sn$ = Get_ListTextByIndex$(AoRoomList&, index&)
    While AoChild& = 0&
        AoChild& = Find_Child&(Sn$)
    Wend
        AoCheckbox& = Find_SubChild2&(AoChild&, "_aol_checkbox")
            If ignore = True Then SetCheckVal& = 1&
        GetCheckVal& = SendIt&(AoCheckbox&, GetCheck, 0&, 0&)
            While SetCheckVal& <> GetCheckVal&
                PostIt& AoCheckbox&, Down&, 0&, 0&
                PostIt& AoCheckbox&, Up&, 0&, 0&
                GetCheckVal& = SendIt&(AoCheckbox&, GetCheck&, 0&, 0&)
            Wend
        SendIt& AoChild&, CloseWin, 0&, 0&
End Sub

