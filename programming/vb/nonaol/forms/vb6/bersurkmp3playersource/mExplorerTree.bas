Attribute VB_Name = "mExplorerTree"

Option Explicit

Public Sub subShowFolderList(oFolderList As ListBox, oExplorerTree As TreeView, sDriveLetter As String, vParentID As Variant)
    
    Dim nNode As Node
    Dim lReturn As Long
    Dim lNextFile As Long
    Dim sPath As String
    Dim WFD As WIN32_FIND_DATA
    Dim sFolderName As String
    Dim x As Long
    Set oFolderList = frmExplore.List1
    Set oExplorerTree = frmExplore.Explorer
       
        

    sPath$ = (sDriveLetter & "*.*") & Chr$(0)
    
    lReturn& = FindFirstFile(sPath$, WFD)
    
    Do

        If (WFD.dwFileAttributes And vbDirectory) Then
            
            sFolderName$ = mProcFunc.ftnStripNullChar(WFD.cFileName)
            If sFolderName$ <> "." And sFolderName$ <> ".." Then
                
                If WFD.dwFileAttributes <> 16 Then
                    oFolderList.AddItem sFolderName$ & "~A~"
                Else
                    oFolderList.AddItem sFolderName$ & "~~~"
                End If

            End If
        End If
        
        lNextFile& = FindNextFile(lReturn&, WFD)
    
    Loop Until lNextFile& = False
  
    lNextFile& = FindClose(lReturn&)

    For x = 0 To oFolderList.ListCount - 1

        If Right(oFolderList.List(x), 3) = "~A~" Then
            Set nNode = oExplorerTree.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
            nNode.ForeColor = RGB(120, 120, 120)
        Else
            Set nNode = oExplorerTree.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
        End If
    
    Next x

    oFolderList.Clear

End Sub




