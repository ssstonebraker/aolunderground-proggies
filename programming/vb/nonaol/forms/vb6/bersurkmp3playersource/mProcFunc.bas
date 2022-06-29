Attribute VB_Name = "mProcFunc"
Public Function ftnStripNullChar(sInput As String) As String

    Dim X As Integer
    
    X = InStr(1, sInput$, Chr$(0))

    If X > 0 Then
        ftnStripNullChar = Left(sInput$, X - 1)
    End If

End Function


Public Function ftnReturnNodePath(sExplorerPath As String) As String

    Dim iSearch(1) As Integer
    Dim sRootPath As String
    
    iSearch%(0) = InStr(1, sExplorerPath$, "(", vbTextCompare)
    iSearch%(1) = InStr(1, sExplorerPath$, ")", vbTextCompare)
    
    If iSearch%(0) > 0 Then
        sRootPath$ = Mid(sExplorerPath$, iSearch%(0) + 1, 2)
    End If
    
    If iSearch%(1) > 0 Then
        ftnReturnNodePath$ = sRootPath$ & Mid(sExplorerPath$, iSearch%(1) + 1, Len(sExplorerPath$)) & "\"
    End If
    
End Function

Public Sub subSetLightColour(sLightSource As String)

    Dim X As Integer

    Select Case sLightSource$

        Case "speCommandLight"
            For X = 0 To 5

            Next
            
    End Select
End Sub

Public Sub subSendMCIMessage(sCommand As String)
    
    Dim lERReturn(1) As Long
    Dim sError As String * 256
    
    
    lERReturn&(1) = mciSendString(sCommand$, 0, 0, hwnd)
    
    mciGetErrorString lERReturn&(1), sError$, Len(sError$)
    frmMain.Caption = sError$
    
End Sub

Public Function ftnRandomSelect() As Integer
       
    Dim iSearch As Integer
    Dim iRandomNo As Integer
       
    Randomize
    Do
        
        With frmMain.lstFiles
            
            iRandomNo% = Int((.ListItems.Count * Rnd) + 1)
            
            If .ListItems(iRandomNo%).ListSubItems(3).Text = "P" Then
            Else
                ftnRandomSelect% = iRandomNo%
                mVariables.iRandomCount = mVariables.iRandomCount + 1
            Exit Function
            End If
        
        End With

    Loop Until mVariables.iRandomCount >= frmMain.lstFiles.ListItems.Count

End Function


Public Sub subSetVolume(sVolumeSet As String)

    Select Case sVolumeSet$
                
        Case "Increase"
            
            If mVariables.iVolumeSetting < 995 Then
                
                mVariables.iVolumeSetting = mVariables.iVolumeSetting + 166
    
                mciSendString "setaudio mp3 volume to " & mVariables.iVolumeSetting, 0, 0, 0

            End If
        
        Case "Decrease"
        
            If mVariables.iVolumeSetting > 331 Then
                
                mVariables.iVolumeSetting = mVariables.iVolumeSetting - 166
                                
                mciSendString "setaudio mp3 volume to " & mVariables.iVolumeSetting, 0, 0, 0
            
            End If

    End Select

    Call subVolumeInd(mVariables.iVolumeSetting)

End Sub


Private Sub subVolumeInd(iVolume As Integer)
        
        
    With frmMain
        
        If iVolume% <= 166 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = &H80&
            .VolumeInd(2).FillColor = &H80&
            .VolumeInd(3).FillColor = &H80&
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 166 And iVolume% <= 332 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = &H80&
            .VolumeInd(3).FillColor = &H80&
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 332 And iVolume% <= 498 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = &H80&
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 498 And iVolume% <= 664 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = RGB(250, 0, 0)
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 664 And iVolume% <= 830 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = RGB(250, 0, 0)
            .VolumeInd(4).FillColor = RGB(250, 0, 0)
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 830 And iVolume% <= 996 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = RGB(250, 0, 0)
            .VolumeInd(4).FillColor = RGB(250, 0, 0)
            .VolumeInd(5).FillColor = RGB(250, 0, 0)
        End If
        
    End With

End Sub

