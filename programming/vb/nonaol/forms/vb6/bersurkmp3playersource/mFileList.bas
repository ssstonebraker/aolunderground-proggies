Attribute VB_Name = "mFileList"
Option Explicit
                                                                                    
            
Public Sub subFileList(sFolderPath As String)

    Dim lReturn As Long
    Dim lNextFile As Long
    Dim sPath As String
    Dim WFD As WIN32_FIND_DATA
    Dim lstItem As ListItem
    Dim lstSubItem As ListSubItem
    Dim sFileName As String
    Dim oFileList As ListView
        Set oFileList = frmExplore.FileList
        sPath$ = sFolderPath$ & "*.mp3"
    Dim lFileLoop As Long
   
    With oFileList
        
        .Visible = False
        .ListItems.Clear
    
        lReturn& = FindFirstFile(sPath$, WFD) & Chr$(0)
        frmExplore.MousePointer = 11
                
        Do
                       
            If Not (WFD.dwFileAttributes And vbDirectory) = vbDirectory Then
        
                sFileName$ = mProcFunc.ftnStripNullChar(WFD.cFileName)
            
                If sFileName > Trim("") Then
                        Set lstItem = .ListItems.Add(, , sFileName$)
                        Set lstSubItem = lstItem.ListSubItems.Add(, , Format(WFD.nFileSizeLow, "#,0"))
                End If
            
            End If
        
            lNextFile& = FindNextFile(lReturn&, WFD)
        
        Loop Until lNextFile& <= Val(0)

        frmExplore.MousePointer = 0
    
        lNextFile& = FindClose(lReturn&)
        
        For lFileLoop = 1 To .ListItems.Count

            If InStrRev(LCase(.ListItems(lFileLoop).Text), ".mp3", , vbTextCompare) Then
                .ListItems(lFileLoop).ForeColor = RGB(60, 60, 140)
            
            End If
        
        Next
        
        .Visible = True
    
    End With

End Sub

