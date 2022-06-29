VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Sub subLoadTreeView()

    Dim TreeList As TreeView
        Set TreeList = frmExplore.Explorer
    Dim iDriveNum As Integer
    Dim sDriveType As String
    Dim fso, d As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
    Dim x As Integer
    Dim RetStr(1) As String
    Dim nNode As Node
    Dim sComputerName As String
        sComputerName$ = mVariables.sComputerName
                        
    RetStr$(0) = ftnShowDriveList

    With TreeList

        Set nNode = .Nodes.Add(, , sComputerName$, sComputerName$, "mycomputer", "mycomputer")
        
        
        For x = 1 To Len(RetStr$(0))
            
            RetStr$(1) = Mid(RetStr$(0), x, 1)
            
            Set d = fso.GetDrive(RetStr$(1))
            iDriveNum% = x
                       
            Select Case d.drivetype
                
                Case 0: sDriveType = "Unknown"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvunknown")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvunknown")
                    End If

                Case 1: sDriveType = "Removable"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":", "drvremove")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvremove")

                    End If
                    
                Case 2: sDriveType = "Fixed"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvfixed")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvfixed")
                    End If

                Case 3: sDriveType = "Network"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvremote")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvremote")
                    End If
                    
                Case 4: sDriveType = "CD-ROM"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvcd")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvcd")
                    End If
                    
                Case 5: sDriveType = "Ram Disk"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvram")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvram")
                    End If
            
            End Select
            
        Next x

    End With

End Sub


Private Function ftnShowDriveList()
  
    Dim fso, d, dc As Object
    Dim sDriveLetter As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set dc = fso.Drives
    
    For Each d In dc
        sDriveLetter$ = sDriveLetter$ & d.driveletter
    Next
    
    ftnShowDriveList = sDriveLetter$

End Function

