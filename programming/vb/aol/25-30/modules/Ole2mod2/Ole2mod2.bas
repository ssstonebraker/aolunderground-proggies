Option Explicit

Global MDINew As Integer

Sub NewObject ()
  MDINew = True
  NewOleForm
  If MDIfrm.ActiveForm.Ole1.OLEType <> OLE_NONE Then
    MDIfrm.ActiveForm.Ole1.Action = OLE_ACTIVATE
  Else
    Unload MDIfrm.ActiveForm
  End If
End Sub

Sub NewOleForm ()
Dim Newform As New frmOLE
Newform.Show
UpdateCaption
End Sub

Sub OpenObject ()
  MDINew = False
  NewOleForm
  OpenSave ("Open")
  If MDIfrm.ActiveForm.Ole1.OLEType = OLE_NONE Then
    Unload MDIfrm.ActiveForm
  End If
End Sub

Sub OpenSave (Action As String)
Dim Filenum
Filenum = FreeFile


  ' Set common dialog options.
  MDIfrm.ActiveForm.CMDialog1.Filter = "OLE 2.0 Objects|*.OLE"
  MDIfrm.ActiveForm.CMDialog1.FilterIndex = 1
  
  MDIfrm.ActiveForm.Ole1.FileNumber = Filenum

On Error Resume Next

Select Case Action
Case "Save"
  ' Display Save As dialog.
  MDIfrm.ActiveForm.CMDialog1.Action = 2
  If Err Then
    ' user pressed cancel
    If Err = 32755 Then
      Exit Sub
    Else
      MsgBox "An unanticipated error occured with the Save As dialog."
    End If
  End If
  ' Open and save the file.
  Open MDIfrm.ActiveForm.CMDialog1.Filename For Binary As Filenum
  If Err Then
    MsgBox (Error)
    Exit Sub
  End If
  MDIfrm.ActiveForm.Ole1.Action = OLE_SAVE_TO_FILE
  If Err Then MsgBox (Error)

Case "Open"
  ' Display File Open dialog.
  MDIfrm.ActiveForm.CMDialog1.Action = 1
  If Err Then
    ' user pressed cancel
    If Err = 32755 Then
      Exit Sub
    Else
      MsgBox "An unanticipated error occured with the Open As dialog."
    End If
  End If
  ' Open the file.
  Open MDIfrm.ActiveForm.CMDialog1.Filename For Binary As Filenum
  If Err Then
    Exit Sub
  End If
  ' Display hourglass.
  Screen.MousePointer = 11
  MDIfrm.ActiveForm.Ole1.Action = OLE_READ_FROM_FILE
  If (Err) Then
    If Err = 30015 Then
      MsgBox "Not a valid OLE object."
    Else
      MsgBox Error$
    End If
    Unload MDIfrm.ActiveForm
  End If

  ' Set form properties now that OLE control contains an object.
  UpdateCaption
  ' Restore mouse pointer.
  Screen.MousePointer = 0
End Select
  
Close Filenum
End Sub

Sub UpdateCaption ()
  Dim Verb
  ' Set Form properties now that it contains an object.
  MDIfrm.ActiveForm.Caption = MDIfrm.ActiveForm.Ole1.Class + " Object"
  MDIfrm.ActiveForm.mnuObject.Caption = MDIfrm.ActiveForm.Ole1.Class + " " + MDIfrm.ActiveForm.mnuObject.Caption

  On Error Resume Next
  For Verb = 1 To VerbMax
    Load MDIfrm.ActiveForm.mnuVerbs(Verb)
    If Err = 360 Then ' Object already loaded.
      Unload MDIfrm.ActiveForm.mnuVerbs(Verb)
      Load MDIfrm.ActiveForm.mnuVerbs(Verb)
      Err = 0
    End If
  Next Verb
  MDIfrm.ActiveForm.mnuVerbs(0).Visible = False
End Sub

