Attribute VB_Name = "ModOLECont"
Option Explicit

Public MDINew As Integer

Sub NewObject()
    MDINew = True
    NewOleForm
    If MDIfrm.ActiveForm.Ole1.OLEType = vbOLENone Then
        Unload MDIfrm.ActiveForm
    End If
End Sub

Sub DisplayInstructions()
    ' Declare local variables.
    Dim MsgText
    Dim PB
    ' Initialize the paragraph break variable.
    PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
    ' Display the instructions.
    MsgText = "To insert a new object, choose New from the File menu, and then select an object from the Insert Object dialog box."
    MsgText = MsgText & PB & "Once you have saved an inserted object using the Save As command, you can use the Open command on the File menu to view the object in subsequent sessions."
    MsgText = MsgText & PB & "To edit an object, double-click the object to display the editing environment for the application from which the object originated."
    MsgText = MsgText & PB & "Click the object with the right mouse button to view the object's verbs."
    MsgText = MsgText & PB & "Use the Copy, Delete, and Paste Special commands to copy, delete, and paste objects."
    MsgText = MsgText & PB & "Choose the Update command to update the contents of the insertable object."
    MsgBox MsgText, 64, "OLE Container Control Demo Instructions"
End Sub

Sub NewOleForm()
    Dim Newform As New frmOLE
    Newform.Show
    ' Only display the Insert Object dialog box if the user chose New from the File menu.
    If MDINew Then
        MDIfrm.ActiveForm.Ole1.InsertObjDlg
    End If
    
    UpdateCaption
End Sub

Sub OpenObject()
    MDINew = False
    NewOleForm
    OpenSave ("Open")
    If MDIfrm.ActiveForm.Ole1.OLEType = vbOLENone Then
        Unload MDIfrm.ActiveForm
    End If
End Sub

' Opening a new file will only work with a file that contains a valid OLE Automation object.
' To see this work, follow this procedure while the application is running.
' 1) From the File menu, choose New, and then specify an object.
' 2) Edit the object, and then choose Save As from the File menu.
' 3) Click the menu-control box for the object to close it.
' 4) From the File menu, choose Open, and then select the file you just saved.
Sub OpenSave(Action As String)
    Dim Filenum
    Filenum = FreeFile
    ' Set the common dialog options and filters.
    MDIfrm.ActiveForm.CommonDialog1.Filter = _
      "Insertable objects (*.OLE)|*.OLE|All files (*.*)|*.*"
    MDIfrm.ActiveForm.CommonDialog1.FilterIndex = 1
  
    MDIfrm.ActiveForm.Ole1.FileNumber = Filenum

On Error Resume Next

    Select Case Action
        Case "Save"
            ' Display the Save As dialog box.
            MDIfrm.ActiveForm.CommonDialog1.ShowSave
            If Err Then
                ' User chose Cancel.
                If Err = 32755 Then
                    Exit Sub
                Else
                    MsgBox "An unanticipated error occurred with the Save As dialog box."
                End If
            End If
            ' Open and save the file.
            Open MDIfrm.ActiveForm.CommonDialog1.filename For Binary As Filenum
            If Err Then
                MsgBox (Error)
                    Exit Sub
            End If
                MDIfrm.ActiveForm.Ole1.SaveToFile Filenum
            If Err Then MsgBox (Error)

        Case "Open"
            ' Display File Open dialog box.
            MDIfrm.ActiveForm.CommonDialog1.ShowOpen
            If Err Then
                ' User chose Cancel.
                If Err = 32755 Then
                    Exit Sub
                Else
                    MsgBox "An unanticipated error occurred with the Open File dialog box."
                End If
            End If
            ' Open the file.
            Open MDIfrm.ActiveForm.CommonDialog1.filename For Binary As Filenum
            If Err Then
                Exit Sub
            End If
            ' Display the hourglass mouse pointer.
            Screen.MousePointer = 11
            MDIfrm.ActiveForm.Ole1.ReadFromFile Filenum
            If (Err) Then
                If Err = 30015 Then
                    MsgBox "Not a valid object."
                Else
                    MsgBox Error$
                End If
                Unload MDIfrm.ActiveForm
            End If
            ' If no errors occur during open, activate the object.
            MDIfrm.ActiveForm.Ole1.DoVerb -1

        ' Set the form properties now that the OLE container control contains an object.
        UpdateCaption
        ' Restore the mouse pointer.
        Screen.MousePointer = 0
    End Select
  
    Close Filenum
End Sub

Sub UpdateCaption()
    ' Set the form properties now that it contains an object.
    MDIfrm.ActiveForm.Caption = MDIfrm.ActiveForm.Ole1.Class + " Object"
    On Error Resume Next
End Sub

