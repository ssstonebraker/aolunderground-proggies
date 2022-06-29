Option Explicit

' Define a string variable to hold the name and path of the database
Global gsDatabaseName As String

' Define a variable to be used by the system and bug edit forms to show
' if the user cancelled any changes
Global gCancel As Integer

' Define a new type to hold the bug details information, as is held in the bugs table
Type BugDetails

    nBugID As Long
    nSystemID As Long
    varCreatedOn As Variant
    nCleared As Integer
    sNotes As String
    sFileName As String * 50
    sProcedure As String * 50
    sDescription As String * 35

End Type

Sub EditBug (CurrentBug As BugDetails, sType As String)
    Dim nIndex As Integer
    
    screen.MousePointer = 11

    Load frmBugs

    frmBugs!lblSystem.Caption = frmMainForm!cboSystemName.Text
    frmBugs!lblBugID.Caption = CurrentBug.nSystemID
    frmBugs!txtCreatedOn.Text = CurrentBug.varCreatedOn
    
    If CurrentBug.nCleared Then
        frmBugs!chkCleared.Value = 1
    Else
        frmBugs!chkCleared.Value = 0
    End If

    frmBugs!txtNotes.Text = CurrentBug.sNotes
    frmBugs!txtProcedure.Text = CurrentBug.sProcedure
    frmBugs!txtDescription.Text = CurrentBug.sDescription

    ' Now build up the files combo box.

    frmBugs!cboFilename.Clear
    For nIndex = 0 To frmMainForm!cboFileName.ListCount - 1
    
        frmBugs!cboFilename.AddItem frmMainForm!cboFileName.List(nIndex)

    Next nIndex

    frmBugs!cboFilename.Text = CurrentBug.sFileName

    screen.MousePointer = 0

    frmBugs.Show 1

    If gCancel = True Then
        Unload frmBugs
        Exit Sub
    End If

    screen.MousePointer = 11

    CurrentBug.nSystemID = frmBugs!lblBugID.Caption
    CurrentBug.varCreatedOn = frmBugs!txtCreatedOn.Text
    CurrentBug.nCleared = frmBugs!chkCleared.Value
    CurrentBug.sNotes = frmBugs!txtNotes.Text
    CurrentBug.sProcedure = frmBugs!txtProcedure.Text
    CurrentBug.sDescription = frmBugs!txtDescription.Text
    CurrentBug.sFileName = frmBugs!cboFilename.Text

    If sType = "NEW" Then

        ' Add a new record
        frmMainForm!datBugs.Recordset.AddNew
        frmMainForm!datBugs.Recordset.Update
        frmMainForm!datBugs.Recordset.Bookmark = frmMainForm!datBugs.Recordset.LastModified
        CurrentBug.nBugID = frmMainForm!datBugs.Recordset.Fields("Bug_ID")
        Call UpdateBugRecord(CurrentBug)

    Else
            
        ' Update the existing one
        Call UpdateBugRecord(CurrentBug)

    End If

    Unload frmBugs

    screen.MousePointer = 0

End Sub

Sub SetDatabaseName (frmForm As Form, ByVal nRefresh As Integer)

    '-------------------------------------------------------------------------------------------
    '   Name    :   SetDatabaseName - procedure to set up the datacontrols on the form specified
    '
    '   Notes   :   This routine sets up the datacontrols with the name of the database
    '           :   obtained during the main form's Load Event.
    '
    '   Params  :   frmForm     The form containing the data controls to set
    '           :   nRefresh    True/False flag indicating if the controls should be refreshed
    '-------------------------------------------------------------------------------------------

    ' Define a variable to be used to count the controls on the form
    Dim nControl As Integer


    ' Loop through the controls on the form
    For nControl = 0 To frmForm.Controls.Count - 1

        ' If the current control is a data control then
        If TypeOf frmForm.Controls(nControl) Is Data Then

            ' Set the DatabaseName property of the data control
            frmForm.Controls(nControl).DatabaseName = gsDatabaseName

            ' If the calling code has asked to refresh the data control, refresh it
            If nRefresh Then frmForm.Controls(nControl).Refresh

        End If

    Next

End Sub

Sub UpdateBugRecord (CurrentBug As BugDetails)
    
    '-------------------------------------------------------------------------------------------
    '   Name    :   UpdateBugRecord - Updates the current record (EDIT)
    '
    '   Notes   :   This updates the bug record indicated in the CurrentBug structure. This means
    '           :   simply finding the specified bug, editing, writing, updating and then
    '           :   rebuilding the grid
    '-------------------------------------------------------------------------------------------

    ' Find the specified bug
    frmMainForm.datBugs.Recordset.FindFirst "Bug_ID = " & CurrentBug.nBugID

    ' There is no question that the bug will be found, since it was there a few minutes ago.
    ' Start an Edit
    frmMainForm.datBugs.Recordset.Edit

    ' Copy the values from the CurrentBug structure to the record itself.
    frmMainForm.datBugs.Recordset.Fields("System_ID") = CurrentBug.nSystemID
    frmMainForm.datBugs.Recordset.Fields("CreatedOn") = CurrentBug.varCreatedOn
    frmMainForm.datBugs.Recordset.Fields("Cleared") = CurrentBug.nCleared
    frmMainForm.datBugs.Recordset.Fields("Notes") = CurrentBug.sNotes
    frmMainForm.datBugs.Recordset.Fields("FileName") = CurrentBug.sFileName
    frmMainForm.datBugs.Recordset.Fields("Procedure") = CurrentBug.sProcedure
    frmMainForm.datBugs.Recordset.Fields("Description") = Left$(CurrentBug.sDescription, 35)

    ' Update the database
    frmMainForm.datBugs.Recordset.Update

End Sub

