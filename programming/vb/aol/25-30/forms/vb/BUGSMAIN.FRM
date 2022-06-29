VERSION 2.00
Begin Form frmMainForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Bug Tracker"
   ClientHeight    =   4350
   ClientLeft      =   375
   ClientTop       =   2550
   ClientWidth     =   8700
   Height          =   5040
   Left            =   315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8700
   Top             =   1920
   Width           =   8820
   Begin PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   3540
      Left            =   120
      ScaleHeight     =   3510
      ScaleWidth      =   8400
      TabIndex        =   5
      Top             =   660
      Width           =   8430
      Begin CommandButton mnuQuit 
         Caption         =   "&Quit"
         Height          =   345
         Left            =   7125
         TabIndex        =   13
         Top             =   3045
         Width           =   1035
      End
      Begin CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   3330
         TabIndex        =   12
         Top             =   3045
         Width           =   1035
      End
      Begin CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   345
         Left            =   2235
         TabIndex        =   11
         Top             =   3045
         Width           =   1035
      End
      Begin CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   1155
         TabIndex        =   10
         Top             =   3045
         Width           =   1035
      End
      Begin CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   345
         Left            =   90
         TabIndex        =   9
         Top             =   3030
         Width           =   1035
      End
      Begin Grid grdBugs 
         BackColor       =   &H00C0C0C0&
         Cols            =   4
         FixedCols       =   0
         FixedRows       =   0
         ForeColor       =   &H00FF0000&
         Height          =   2280
         Left            =   75
         Rows            =   10
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   660
         Width           =   8055
      End
      Begin CheckBox chkShowUncleared 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show Uncleared Bugs"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6030
         TabIndex        =   7
         Top             =   45
         Width           =   2280
      End
      Begin CheckBox chkShowCleared 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show Cleared Bugs"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   45
         Width           =   2040
      End
   End
   Begin Data datBugs 
      Caption         =   "Bugs Data Control"
      Connect         =   ""
      DatabaseName    =   "D:\BG2VB\BUGS.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   195
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "BUGS"
      Top             =   5280
      Width           =   3135
   End
   Begin Data datSystems 
      Caption         =   "Systems Data Control"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   195
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "System"
      Top             =   5010
      Visible         =   0   'False
      Width           =   3135
   End
   Begin PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H0000FFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   150
      ScaleHeight     =   510
      ScaleWidth      =   8370
      TabIndex        =   0
      Top             =   60
      Width           =   8400
      Begin ComboBox cboFileName 
         Height          =   300
         Left            =   5520
         TabIndex        =   4
         Top             =   90
         Width           =   2655
      End
      Begin ComboBox cboSystemName 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   120
         Width           =   2655
      End
      Begin Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4500
         TabIndex        =   3
         Top             =   150
         Width           =   1020
      End
      Begin Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "System Name"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   105
         TabIndex        =   2
         Top             =   150
         Width           =   1290
      End
   End
   Begin Menu mnuSystem 
      Caption         =   "&System"
      Begin Menu mnuSNew 
         Caption         =   "&New System"
      End
      Begin Menu mnuSEdit 
         Caption         =   "&Edit System Details"
      End
      Begin Menu mnuSDelete 
         Caption         =   "&Delete System"
      End
   End
   Begin Menu mnuBugs 
      Caption         =   "&Bugs"
      Begin Menu mnuBNew 
         Caption         =   "&New Bug"
      End
      Begin Menu mnuBEdit 
         Caption         =   "&Edit Bug"
      End
      Begin Menu mnuBClear 
         Caption         =   "&Clear Bug"
      End
      Begin Menu mnuBDelete 
         Caption         =   "&Delete Bug"
      End
   End
   Begin Menu mnuReports 
      Caption         =   "&Reports"
      Begin Menu mnuRBugs 
         Caption         =   "System &Bugs"
      End
   End
End
Option Explicit


Dim mnBugID As Long
Dim msFileName As String

Sub cboFileName_Click ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   cboFileName_Click   - Occurs when the user selects a new file name from
    '           :   the grid.
    '
    '   Notes   :   This routine rebuilds the bugs grid whenever the user selects a new file
    '           :   name. This causes only the bugs in the selected file to be displayed.
    '-------------------------------------------------------------------------------------------

    ' Rebuild the bugs grid. This will reflect the bugs in the currently selected file.
    Call RebuildBugsGrid
    
End Sub

Sub cboSystemName_Click ()
    
    '-------------------------------------------------------------------------------------------
    '   Name    :   cboSystemName_Click  - Event occurs when the user selects a new system
    '
    '   Notes   :   This routine rebuilds the bugs grid whenever the user selects a new system.
    '           :   This causes only the bugs in the selected system to be displayed.
    '-------------------------------------------------------------------------------------------

    ' Rebuild the filenames combo
    Call RebuildFileNames

End Sub

Sub chkShowCleared_Click ()
    
    '-------------------------------------------------------------------------------------------
    '   Name    :   chkShowCleared_Click    - User has click the Show Cleared check box
    '
    '   Notes   :   This will cause the grid to be rebuilt reflecting only those bugs in the
    '           :   current file and system which have been cleared.
    '-------------------------------------------------------------------------------------------

    ' Rebuild the bugs grid
    Call RebuildBugsGrid

End Sub

Sub chkShowUncleared_Click ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   chkShowUnCleared_Click    - User has click the Show Uncleared box
    '
    '   Notes   :   This will cause the grid to be rebuilt reflecting only those bugs in the
    '           :   current file and system which have NOT been cleared.
    '-------------------------------------------------------------------------------------------
    
    ' Rebuild the bugs grid
    Call RebuildBugsGrid

End Sub

Sub cmdClear_Click ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   cmdClear_Click  - User wants to clear/unclear a bug.
    '
    '   Notes   :   Inverts the clear field of the selected bug. The grid is rebuilt after that
    '           :   since the check boxes above the grid may be set such that this bug no longer
    '           :   matches the criteria.
    '-------------------------------------------------------------------------------------------

    If grdBugs.Row = 0 Then Exit Sub

    ' Find the bug in the current grid row
    datBugs.Recordset.FindFirst "Bugs.Bug_ID = " & mnBugID

    ' Start an edit
    datBugs.Recordset.Edit

    ' Set the cleared value to its opposite (True to False, False to True)
    datBugs.Recordset.Fields("Cleared") = Not datBugs.Recordset.Fields("Cleared")

    ' Update the record
    datBugs.Recordset.Update
    
    ' Rebuild the grid
    Call RebuildBugsGrid

End Sub

Sub cmdDelete_Click ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   cmdDelete_Click  - Delete the current bug
    '
    '   Notes   :   The user is first questioned to make sure that he or she knows what is about
    '           :   to happen. If the user still wants to continue then the bug is deleted, as
    '           :   opposed to being deleted.
    '-------------------------------------------------------------------------------------------

    Dim nresponse As Integer

    If grdBugs.Row = 0 Then Exit Sub

    ' Ask the user if he/she really wants to delete this bug
    nresponse = MsgBox("Do you really want to delete this bug?", 36, "Delete")

    ' If the answer is yes then
    If nresponse = 6 Then

        ' Find the record on the currently selected row of the grid
        datBugs.Recordset.FindFirst "Bugs.Bug_ID = " & mnBugID

        ' Delete it
        If Not datBugs.Recordset.NoMatch Then datBugs.Recordset.Delete

        Call RebuildBugsGrid

    End If

End Sub

Sub cmdEdit_Click ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   cmdEdit_Click   - User wants to edit the currently selected bug
    '
    '   Notes   :   This runs up the edit bugs form with the currently selected bug displayed.
    '           :   The bug id is passed in the bug ID global variable since this is the
    '           :   easiest and most efficient way of passing parameters to a form
    '-------------------------------------------------------------------------------------------
    
    Dim CurrentBug As BugDetails
    Dim nUpdate As Integer

    ' Run up the bug edit form as modal
    If SetBugType(CurrentBug, "") Then
    
        Call EditBug(CurrentBug, "EDIT")
    
        ' Rebuild the grid
        Call RebuildBugsGrid
    End If

End Sub

Sub cmdNew_Click ()
    
    '-------------------------------------------------------------------------------------------
    '   Name    :   cmdNew_Click    - Occurs when the user elects to add a new bug to the system
    '
    '   Notes   :   This causes the new bug form to be loaded, with the system name and file
    '           :   name being passed to the form within global variables (its faster and better
    '           :   practice than using properties across forms.
    '
    '           :   Validation is included in that the routine checks to see if the a system
    '           :   has currently been selected. If none has then the event will not run up
    '           :   the new bug form, since the program does not know which system the bug
    '           :   applies to.
    '-------------------------------------------------------------------------------------------
    
    Dim nUpdate As Integer
    Dim CurrentBug As BugDetails

    ' If no system has currently been selected then
    If cboSystemName.Text = "<All Systems>" Then

        ' Display an error message - you must select a system first
        MsgBox "You must select a system, before adding bugs to it", , "Error"

        ' Exit the subroutine
        Exit Sub

    End If

    ' If a filename has been selected then store it in the global filename variable.
    If cboFileName.Text <> "<All Files>" Then msFileName = cboFileName.Text

    ' Run up the bug edit form as modal
    If SetBugType(CurrentBug, "NEW") Then

        Call EditBug(CurrentBug, "NEW")
    
        ' Rebuild the bugs grid
        Call RebuildBugsGrid

    End If

End Sub

Sub Form_Load ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   Form_Load => Main form, frmMainfrm
    '
    '   Notes   :   This routine will only ever run once, that is when the program has just
    '           :   started up. It performs a number of steps
    '           :
    '           :       1.  Checks to see if it is already running and terminates if it is.
    '           :       2.  Sets up a global variable holding the full name of the database.
    '           :       3.  Initialises the data controls by placing the database name in them.
    '           :       4.  Initialises the combo boxes on the main form.
    '           :       5.  Initialises the grid on the main form.
    '-------------------------------------------------------------------------------------------

    ' Define a variable to be used for the message boxes
    Dim sMessage As String

    Start3d
    
    ' Check the App.Instance property to see if the program is already running. If it is then
    ' display an error message and close the application down.
    If app.PrevInstance Then
        sMessage = "Sorry, Bug Tracker is already running." & Chr$(13)
        sMessage = sMessage & "Try using Alt-Tab to switch to the other copy."
        MsgBox sMessage
        End
    End If

    ' Set up the global variable holding the database name
    gsDatabaseName = app.Path & "\bugs.mdb"

    ' Initialise the data controls
    Call SetDatabaseName(frmMainForm, True)
    
    ' Set up the column widths of the bugs grid
    Call SetGridColWidths

    ' Rebuild the system name combo box
    Call RebuildSystemCombo

    ' Show the form
    frmMainForm.Show

    frmMainForm.Left = (screen.Width - frmMainForm.Width) / 2
    frmMainForm.Top = (screen.Height - frmMainForm.Height) / 2

End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
    
    '-------------------------------------------------------------------------------------------
    '   Name    :   Form_QueryUnload - "just before the form shuts down" event
    '
    '   Notes   :   Routine called whenever the controlbox is double clicked, or the user
    '           :   hits the quit button on the main form.
    '-------------------------------------------------------------------------------------------

    Dim nresponse As Integer

    ' Ask if the user really wants to quit the application and didnt just press quit by mistake
    nresponse = MsgBox("Really quit?", 36, "Quit")

    ' Set the cancel variable according to what the user decided to do!
    If nresponse = 6 Then Cancel = False Else Cancel = True

End Sub

Sub Form_Unload (Cancel As Integer)

    End3d

End Sub

Sub grdBugs_RowColChange ()
    
    '-------------------------------------------------------------------------------------------
    '   Name    :   grdBugs_RowColChange    - Occurs whenever the user moves to a new cell
    '
    '   Notes   :   When the user moves to a new cell we have to assume that they have in fact
    '           :   moved to a new row. This way we can calculate the current bug ID of the
    '           :   row selected and store it in a global variable ready to be passed to the
    '           :   bug details form, if the user selects the Edit option.
    '-------------------------------------------------------------------------------------------

    If grdBugs.Row = 0 Then Exit Sub

    DoEvents
    
    ' Set the current column to the first column
    If grdBugs.Col <> 0 Then grdBugs.Col = 0
    
    ' Get the bug id of the currently selected bug and store it in the global variable
    mnBugID = Val(grdBugs.Text)

End Sub

Sub mnuBClear_Click ()

    ' Call the clear button click event
    Call cmdClear_Click

End Sub

Sub mnuBDelete_Click ()

    ' Call the delete click event
    Call cmdDelete_Click

End Sub

Sub mnuBEdit_Click ()

    ' Call the Edit button click event
    Call cmdEdit_Click

End Sub

Sub mnuBNew_Click ()

    ' Call the New bug command button click event
    Call cmdNew_Click

End Sub

Sub mnuQuit_Click ()

    ' Unload the form - all validation is held in QueryUnload
    Unload frmMainForm

End Sub

Sub mnuRBugs_Click ()

    Load frmReport
    frmReport.Show 1

End Sub

Sub mnuSDelete_Click ()

    Dim nresponse As Integer

    datSystems.Recordset.FindFirst "System_ID = " & cboSystemName.ItemData(cboSystemName.ListIndex)
    If datSystems.Recordset.NoMatch Then
        Beep
        MsgBox "You must select a system before you can delete it"
        Exit Sub
    
    Else

        nresponse = MsgBox("Are you sure you want to delete this system, " & Chr$(10) & "and all its bugs", 36, "Delete System")
        If nresponse = 6 Then

            screen.MousePointer = 11

            datBugs.RecordSource = "Select * from Bugs where Bugs.System_ID = " & datSystems.Recordset.Fields("System_ID")
            datBugs.Refresh
            
            On Error Resume Next

            datBugs.Recordset.MoveFirst
            Do Until datBugs.Recordset.EOF


                datBugs.Recordset.Delete
                datBugs.Recordset.MoveNext
                
            Loop
        
            On Error GoTo 0

            datSystems.Recordset.Delete
        
            Call RebuildSystemCombo

            screen.MousePointer = 0

        End If

    End If

End Sub

Sub mnuSEdit_Click ()

    datSystems.Recordset.FindFirst "System_ID = " & cboSystemName.ItemData(cboSystemName.ListIndex)
    
    If datSystems.Recordset.NoMatch Then
        Beep
        MsgBox "You must select a system first", , "System Error"
        Exit Sub
    End If
    
    Load frmSystem
    frmSystem!txtSystem.Text = datSystems.Recordset.Fields("System_Name")
    frmSystem!txtMakefile.Text = datSystems.Recordset.Fields("MakeFile")
    frmSystem!txtNotes.Text = datSystems.Recordset.Fields("notes")

    frmSystem.Show 1

    If gcancel = True Then
        Unload frmSystem
        Exit Sub
    End If

    datSystems.Recordset.Edit

    datSystems.Recordset.Fields("System_Name") = frmSystem!txtSystem.Text
    datSystems.Recordset.Fields("MakeFile") = frmSystem!txtMakefile.Text
    datSystems.Recordset.Fields("Notes") = frmSystem!txtNotes.Text

    datSystems.Recordset.Update

    Unload frmSystem

    Call RebuildFileNames

End Sub

Sub mnuSNew_Click ()
    
    Load frmSystem
    frmSystem.Show 1

    If gcancel = True Then
        Unload frmSystem
        Exit Sub
    End If

    datSystems.Recordset.AddNew

    datSystems.Recordset.Fields("System_Name") = frmSystem!txtSystem.Text
    datSystems.Recordset.Fields("MakeFile") = frmSystem!txtMakefile.Text
    datSystems.Recordset.Fields("Notes") = frmSystem!txtNotes.Text

    datSystems.Recordset.Update

    Unload frmSystem

    Call RebuildSystemCombo

End Sub

Sub RebuildBugsGrid ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   RebuildBugsGrid - Redraws the grid with a list of bugs in
    '
    '   Notes   :   This routine uses the currently selected system and file name, along with
    '           :   the values of the check box to build up an SQL Select statement to
    '           :   pull a list of all the bugs in the system matching those criteria. The
    '           :   resultant list is then fired into the grid from the out of site data
    '           :   control.
    '-------------------------------------------------------------------------------------------
    
    Dim sGridText As String
    Dim nGridRow As Integer
    Dim sSQL As String

    screen.MousePointer = 11

    ' Start the SQL statement off (Select * from Bugs where)
    sSQL = "Select * From Bugs Where "

    ' If a system is selected then feed that into the SQL statement
    If cboSystemName.Text <> "<All Systems>" Then
        sSQL = sSQL & "Bugs.System_ID = " & cboSystemName.ItemData(cboSystemName.ListIndex) & " and "
    End If

    ' If a file name is selected then feed that into the SQL statement
    If cboFileName.Text <> "<All Files>" Then
        sSQL = sSQL & "Bugs.FileName = '" & cboFileName.Text & "' and "
    End If

    If chkShowCleared.Value = 0 Or chkShowUnCleared = 0 Then

        ' Feed the value of the Show Cleared check box into the SQL statement
        If chkShowCleared.Value = 1 Then sSQL = sSQL & " Bugs.Cleared "
    
        ' Feed the value of the Show Uncleared check box into the SQL statement
        If chkShowUnCleared.Value = 1 Then sSQL = sSQL & " not Bugs.Cleared "

    End If
    
    If Right$(sSQL, 6) = "Where " Then sSQL = Left$(sSQL, Len(sSQL) - 6)

    If Right$(sSQL, 4) = "and " Then sSQL = Left$(sSQL, Len(sSQL) - 4)

    sSQL = sSQL & " Order By Bugs.Bug_ID"

    ' Place the SQL statement into the data control and refresh it
    datBugs.RecordSource = sSQL
    datBugs.Refresh

    ' Clear out the bugs grid
    grdBugs.Rows = 1

    ' Fill the grid with the records from the data control
    On Error Resume Next
    sGridText = ""
    datBugs.Recordset.MoveFirst

    Do Until datBugs.Recordset.EOF

        sGridText = "" & datBugs.Recordset.Fields("Bug_ID") & Chr$(9)
        sGridText = sGridText & datBugs.Recordset.Fields("CreatedOn") & Chr$(9)
        sGridText = sGridText & datBugs.Recordset.Fields("Description") & Chr$(9)
        If datBugs.Recordset.Fields("Cleared") Then
            sGridText = sGridText & "Yes"
        Else
            sGridText = sGridText & "No"
        End If

        grdBugs.AddItem sGridText

        datBugs.Recordset.MoveNext

    Loop

    On Error GoTo 0

    screen.MousePointer = 0

End Sub

Sub RebuildFileNames ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   RebuildFileNames - Redraws the filename combo box
    '
    '   Notes   :   This loads the mak file for the system and gets a list of Basic file names
    '           :   from it (files ending in .Bas and .Frm). These names are then placed
    '           :   into the combo box.
    '-------------------------------------------------------------------------------------------

    Dim varFileNumber As Variant
    Dim sLineInMak As String
    Dim nCharacter As Integer

    screen.MousePointer = 11

    varFileNumber = FreeFile

    ' Clear the combo box
    cboFileName.Clear

    ' Add the <All Files> item to the combo box.
    cboFileName.AddItem "<All Files>"

    ' Find the system record matching the system name in the systems combo box
    datSystems.Recordset.FindFirst "System_Id = " & cboSystemName.ItemData(cboSystemName.ListIndex)
    
    ' If the project field has been set to anything and the file can be found
    If Not datSystems.Recordset.NoMatch Then

        On Error GoTo EndOfBuild

        If Dir$(datSystems.Recordset.Fields("Makefile") & "") <> "" Then


            On Error GoTo 0
            ' Open the file as read only
            Open datSystems.Recordset.Fields("MakeFile") For Input As varFileNumber
    
            ' Step through each line of the mak file
            Do Until EOF(varFileNumber)

                Line Input #varFileNumber, sLineInMak
    
                ' If the line ends in .Bas, or .FRM then
                If UCase(Right$(sLineInMak, 4)) = ".BAS" Or UCase(Right$(sLineInMak, 4)) = ".FRM" Then
    
                    ' Lose the path to get the filename on its own and add it to the combo box
                    For nCharacter = Len(sLineInMak) To 1 Step -1

                        If Mid$(sLineInMak, nCharacter, 1) = "\" Then

                            cboFileName.AddItem Right$(sLineInMak, Len(sLineInMak) - nCharacter)
                            Exit For

                        End If

                    Next

                    If nCharacter = 0 Then cboFileName.AddItem Right$(sLineInMak, Len(sLineInMak) - nCharacter)

                End If

            Loop

            ' Close the mak file
            Close varFileNumber

        End If

    End If

    ' Select the first item in the combo box (this will trigger a click event which will in turn
    ' rebuild the bugs grid)
    
EndOfBuild:
    On Error GoTo 0
    
    cboFileName.ListIndex = 0

    screen.MousePointer = 0

    Exit Sub

End Sub

Sub RebuildSystemCombo ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   RebuildSystemCombo - Rebuilds the system name combo box
    '
    '   Notes   :   This routine steps through the datSystems data control, copying the system
    '           :   names into the System Combo box. An additional item is added to the head
    '           :   of the combo box which is used to indicate that the user wants to see bugs
    '           :   relating to all systems.
    '           :
    '           :   The ItemData collection of the combo box is used to hold the SystemID of each
    '           :   system to make retrieval nice and quick (Remember that the database has been
    '           :   set up with the System name as a non-unique value.
    '-------------------------------------------------------------------------------------------

    screen.MousePointer = 11

    ' Clear out the system name combo box
    cboSystemName.Clear

    ' Add the "All Systems" item to the head of the combo box
    cboSystemName.AddItem "<All Systems>"
    cboSystemName.ItemData(cboSystemName.ListCount - 1) = 0

    On Error Resume Next

    ' Step through all the systems in the database, adding them to the combo box
    datSystems.Recordset.MoveFirst

    Do Until datSystems.Recordset.EOF

        cboSystemName.AddItem datSystems.Recordset.Fields("System_Name")
        cboSystemName.ItemData(cboSystemName.ListCount - 1) = datSystems.Recordset.Fields("System_ID")

        datSystems.Recordset.MoveNext

    Loop

    ' Finally, select the first item in the list automatically. This will trigger a redraw
    ' of the grid and files combo box
    cboSystemName.ListIndex = 0

    On Error GoTo 0

    screen.MousePointer = 0

End Sub

Function SetBugType (CurrentBug As BugDetails, sType As String) As Integer
    
    '-------------------------------------------------------------------------------------------
    '   Name    :   SetBugType  - sets up the Bug Type structure passed as a parameter
    '
    '   Notes   :   This routine sets up the structure passed as a parameter from the fields
    '           :   in the current record.
    '-------------------------------------------------------------------------------------------
    

    If sType = "NEW" Then
        
        CurrentBug.nBugID = 0
        CurrentBug.nSystemID = cboSystemName.ItemData(cboSystemName.ListIndex)
        CurrentBug.varCreatedOn = Now
        CurrentBug.nCleared = False
        CurrentBug.sNotes = ""
        CurrentBug.sFileName = cboFileName.Text
        CurrentBug.sProcedure = ""
        CurrentBug.sDescription = ""

    Else
        
        ' First find the record currently selected in the grid
        datBugs.Recordset.FindFirst "Bug_ID = " & mnBugID
        
        If datBugs.Recordset.NoMatch Then
            Beep
            MsgBox "There is no bug currently selected", , "Error"
            SetBugType = False
            Exit Function
        End If

        CurrentBug.nBugID = datBugs.Recordset.Fields("Bug_ID")
        CurrentBug.nSystemID = datBugs.Recordset.Fields("System_ID")
        CurrentBug.varCreatedOn = datBugs.Recordset.Fields("CreatedOn")
        CurrentBug.nCleared = datBugs.Recordset.Fields("Cleared")
        CurrentBug.sNotes = datBugs.Recordset.Fields("Notes")
        CurrentBug.sFileName = datBugs.Recordset.Fields("FileName")
        CurrentBug.sProcedure = datBugs.Recordset.Fields("Procedure")
        CurrentBug.sDescription = datBugs.Recordset.Fields("Description")

    End If

    SetBugType = True

End Function

Sub SetGridColWidths ()

    '-------------------------------------------------------------------------------------------
    '   Name    :   SetGridColWidths    - Ses up the widths of the columns in the grid
    '
    '   Notes   :   The widths of the columns in a grid are specified in Twips, a independant
    '           :   measuring system that is difficult to relate to the the number of characters
    '           :   we need to display. What we do is use the width of the grid, divide it by
    '           :   the number of characters that will be displayed in total in all the cells
    '           :   and use that number to calculate the width of each column. Confused? Read on.
    '-------------------------------------------------------------------------------------------

    Dim nCharWidth As Integer
    
    nCharWidth = grdBugs.Width / 84

    ' First column width (Bug ID number)
    grdBugs.ColWidth(0) = 7 * nCharWidth

    ' Second column ( Date bug was created in database)
    grdBugs.ColWidth(1) = 20 * nCharWidth

    ' Third column ( Description of the bug )
    grdBugs.ColWidth(2) = 48 * nCharWidth

    ' Final column ( Bug cleared Yes/No )
    grdBugs.ColWidth(3) = 10 * nCharWidth

    grdBugs.Row = 0
    grdBugs.Col = 0
    grdBugs.FontBold = True
    grdBugs.Text = "ID"
    
    grdBugs.Col = 1
    grdBugs.Text = "Date"

    grdBugs.Col = 2
    grdBugs.Text = "Description"

    grdBugs.Col = 3
    grdBugs.Text = "Cleared"

End Sub

