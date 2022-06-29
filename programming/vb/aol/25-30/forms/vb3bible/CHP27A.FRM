VERSION 2.00
Begin Form formMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Access Project"
   ClientHeight    =   3948
   ClientLeft      =   648
   ClientTop       =   2160
   ClientWidth     =   4260
   Height          =   4716
   Icon            =   CHP27A.FRX:0000
   Left            =   588
   LinkTopic       =   "Form1"
   ScaleHeight     =   3948
   ScaleWidth      =   4260
   Top             =   1452
   Width           =   4380
   Begin Data dataLog 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Log"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1812
   End
   Begin FileListBox fileBox 
      Height          =   1944
      Left            =   2280
      MultiSelect     =   2  'Extended
      Pattern         =   "*.ico"
      TabIndex        =   9
      Top             =   1800
      Width           =   1812
   End
   Begin CommandButton cmndLoad 
      Caption         =   "&Load"
      Height          =   372
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   852
   End
   Begin CommandButton cmndSeek 
      Caption         =   "See&k"
      Height          =   372
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   972
   End
   Begin DriveListBox drivBox 
      Height          =   288
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1932
   End
   Begin DirListBox dirBox 
      Height          =   1155
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin CommandButton cmndSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3300
      TabIndex        =   4
      Top             =   960
      Width           =   795
   End
   Begin CommandButton cmndDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   795
   End
   Begin CommandButton cmndClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Top             =   960
      Width           =   795
   End
   Begin Data dataMain 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   372
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Icon"
      Top             =   960
      Width           =   912
   End
   Begin TextBox textPathName 
      DataField       =   "PathName"
      DataSource      =   "dataMain"
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   3852
   End
   Begin TextBox textFileName 
      DataField       =   "FileName"
      DataSource      =   "dataMain"
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   3252
   End
   Begin Line Line1 
      BorderWidth     =   3
      X1              =   240
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin Image imagIcon 
      DataField       =   "Icon"
      DataSource      =   "dataMain"
      Height          =   555
      Left            =   3600
      Top             =   60
      Width           =   555
   End
   Begin Menu menuBar 
      Caption         =   "&File"
      Index           =   0
      Begin Menu menuFile 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin Menu menuBar 
      Caption         =   "&Utilities"
      Index           =   1
      Begin Menu menuUtils 
         Caption         =   "&Compact"
         Index           =   0
      End
      Begin Menu menuUtils 
         Caption         =   "&Statistics"
         Index           =   1
      End
      Begin Menu menuUtils 
         Caption         =   "&Log"
         Index           =   2
      End
      Begin Menu menuUtils 
         Caption         =   "&Find"
         Index           =   3
      End
      Begin Menu menuUtils 
         Caption         =   "&Datasheet"
         Index           =   4
      End
   End
End

Dim Scan As Integer                'tracks if we're recursing or not
Dim TotalLoaded As Integer         'tracks number of new entries

Function appPath () As String

    'this function returns the path of the application

    'this isn't covered in the book; we added it to make
    'sure the files would run correctly from the floppy

    tempPath = App.Path
    If Right$(tempPath, 1) = "\" Then
	'remove the trailing \ character
	tempPath = Left$(tempPath, Len(tempPath) - 1)
    End If
    
    appPath = tempPath

End Function

Sub cmndClear_Click ()
     MousePointer = HOURGLASS                        'we're going to be a while
     dataMain.Database.Execute "DELETE * FROM Icon"  'delete all records
     dataMain.Refresh                                'update the recordset
     MousePointer = DEFAULT                          'and we're done
End Sub

Sub cmndDelete_Click ()
     Rem *** you could actually enable the delete with this line: ***
     Rem Kill textPathName.Text & "\" & textFileName.Text
     dataMain.Recordset.Delete                         'get rid of entry
     dataMain.Recordset.MovePrevious                   'and go to a record
End Sub

Sub cmndLoad_Click ()
     MousePointer = HOURGLASS             'we're going to be a while
     If Not Scan Then                     'if we're not recursing
      dataLog.Recordset.AddNew           'start a new log record
      dataLog.Recordset!Date = Now       'and put date and time in it
      dataLog.Recordset.Update           'save new record
      dataLog.Recordset.Edit             'prepare current log record for memo
     End If
     For i = 0 To fileBox.ListCount - 1   'go through entire list box
      If fileBox.Selected(i) Then        'if file was selected
	   dataMain.Recordset.AddNew     'create a new record for it
	   textFileName.Text = fileBox.List(i)     'get file name
	   textPathName.Text = dirBox.Path         'get path name
	   imagIcon.Picture = LoadPicture(textPathName & "\" & textFileName)
	   imagIcon.Refresh              'force image control to repaint
	   On Error Resume Next          'very first record update causes error
	   dataMain.Recordset.Update     'save the new record
	   dataMain.Recordset.MoveLast   'and move to it
	    m$ = textPathName.Text & "\" & textFileName.Text & "  added.  "
	   dataLog.Recordset.Fields("Memo").AppendChunk (m$) 'make a log entry
	   DoEvents                      'give the system time to catch up
	   On Error GoTo 0               'disable trap
	   TotalLoaded = TotalLoaded + 1 'remember how many we've loaded
      End If
     Next i                               'go to next file in list box
     If Not Scan Then                     'if we're not recursing
      dataMain.Refresh                   'update the recordset
       dataLog.Recordset.Update           'save the log entry
     End If
     FreeLocks                            'let the system refresh databases
     MousePointer = DEFAULT               'we're done
End Sub

Sub cmndSave_Click ()
     fn = textFileName.DataChanged        'filename same as record?
     pn = textPathName.DataChanged        'pathname same as record?
     If fn Or pn Then                     'if either has changed,
	  dataMain.Recordset.Update       'Save the record
     Else
	  cmndSave.Enabled = False
     End If
End Sub

Sub cmndSeek_Click ()
     'this routine recurses through all subdirectories below current one
     MousePointer = HOURGLASS             'we're going to be awhile
     Scan = True                          'tell other procedures to recurse
     TotalLoaded = 0                      'start fresh total
     BeginTrans                           'start a transaction block
     dataLog.Recordset.AddNew             'add a new log entry
     dataLog.Recordset!Date = Now         'and add in the current date and time
     dirBox_Change                        'start the recursion process
     dataLog.Recordset.Update             'all done! save the log entry
     msg$ = Str$(TotalLoaded) & " records loaded." & Chr$(10) & Chr$(13)
     If dataMain.Recordset.Transactions Then 'recordset support transactions?
       msg$ = msg$ & "Write changes to disk?"  'ask user to confirm
       ans = MsgBox(msg$, MB_YESNO + MB_ICONQUESTION, "Confirm Entries")
       If ans = IDYES Then                'yup, go ahead and add
	    CommitTrans                   'save changes to disk
       Else                               'nope, stop changes
	    Rollback                      'cancel all changes
       End If
     Else                                 'transactions NOT supported
       MsgBox msg$, , "Seek completed"    'just inform user
     End If
     dataMain.Refresh                     'regenerate recordset
     dataLog.Refresh                      'regenerate recordset
     MousePointer = DEFAULT               'we're done
     Scan = False                         'no more recursion
End Sub

Sub dataMain_Error (DataErr As Integer, Response As Integer)
     Select Case DataErr
       Case 3021                          'no current record
	    Response = 0                  'just ignore it
       Case Else
	    Debug.Print DataErr           'hmm... unanticipated error!
	    Response = 0                  'ignore it
     End Select
End Sub

Sub dataMain_Reposition ()
     cmndSave.Enabled = False             'new record, no changes to save!
End Sub

Sub dataMain_Validate (Action As Integer, Save As Integer)
     If cmndSave.Enabled = True And dataMain.EditMode <> 2 Then  'changed record
       On Error GoTo badFileName          'set a trap for the renaming process
       fn = dataMain.Recordset!filename   'get existing filename
       pn = dataMain.Recordset!PathName   'and existing pathname
       Name pn & "\" & fn As textPathName.Text & "\" & textFileName.Text
       dataMain.UpdateRecord              'aha! success! update the database
     Else
       Select Case Action
	    Case 7, 8, 9, 10, 11          'delete, find, bookmark, close, unload
	      Save = False                'don't bother saving changes
	    Case 5, 6                     'addnew, update
	      Save = True                 'yes, save changes
	    Case Else
       End Select
     End If
     cmndSave.Enabled = False             'no more changes to save!
Exit Sub

badFileName:                              'rename failed
     Beep                                 'enhance the user's mood
     dataMain.UpdateControls              'undo changes
     textFileName.SetFocus                'and go back to text box
Exit Sub
     
End Sub

Sub dirBox_Change ()
     fileBox.Path = dirBox.Path           'reset file list box
     ChDir dirBox.Path                    'and go to the new directory
     
     If Scan Then                         'are we recursing?
       'This recursively searches all subdirectories on drive for any *.ico file
       seekPath$ = dirBox.Path            'directory the box is on
       If Right$(seekPath$, 1) <> "\" Then seekPath$ = seekPath$ + "\" 'root
       If Dir$(seekPath$ & "*.ico") <> "" Then      'if there are any icon files
	    For i = 0 To fileBox.ListCount - 1      'go through each file
	      fileBox.Selected(i) = True  'and select it
	    Next i
	    cmndLoad_Click                'load the files into the database
       End If
       numDirs = dirBox.ListCount         'number of subdirectories
       ReDim dirList(numDirs) As String   'make space for each directory entry
       For i = 0 To numDirs - 1           'go through each subdirectory
	    dirList(i) = dirBox.List(i)   'and add it to the list
       Next i
       For j = 0 To numDirs - 1           'now for each subdirectory in list,
	    DoEvents                      'give the system some time to breathe,
	    dirBox.Path = dirList(j)      'recurse down the tree!
       Next j
     End If
End Sub

Sub drivBox_Change ()
     dirBox.Path = CurDir$(drivBox.Drive) 'set new path
End Sub

Sub Form_Load ()
     dataMain.DatabaseName = appPath() & "\chp27.mdb"
     dataMain.RecordSource = "SELECT * FROM Icon"      'get all records

     dataLog.DatabaseName = appPath() & "\chp27.mdb"

End Sub

Sub Form_Resize ()
     dataMain.Recordset.LockEdits = False 'optimistic locking
End Sub

Sub menuFile_Click (Index As Integer)
     End                                  'end program
End Sub

Sub menuUtils_Click (Index As Integer)
     Select Case Index
       Case 0 '*********************************** Compact
	    MousePointer = HOURGLASS      'we're going to be a while
	    dataMain.Database.Close       'close all databases
	    dataLog.Database.Close
	    formDatasheet!dataGrid.Database.Close
	    
	    CompactDatabase appPath() & "\chp27.mdb", appPath() & "\~chp27.tmp"
	    Kill appPath() & "\chp27.mdb"   'get rid of old file
	    Name appPath() & "\~chp27.tmp" As appPath() & "\chp27.mdb"
	    dataMain.Refresh              'open up all databases again
	    dataLog.Refresh
	    formDatasheet!dataGrid.Refresh
	    MousePointer = NORMAL
       Case 1 '*********************************** Statistics
	    PrintStats                    'prints out the statistics
       Case 2 '*********************************** Log
	    Dim f As form                 'this just shortens subsequent code
	    Set f = formStats
	    f.Show                        'show the stats form
	    f.Cls                         'clear off any existing text
	    dataLog.Recordset.MoveLast    'go to latest log entry
	    Do Until dataLog.Recordset.BOF        'at beginning of time yet?
	      f.Print
	      f.Print dataLog.Recordset!Date      'print date and time of entry
	      totalSize = dataLog.Recordset.Fields("Memo").FieldSize()
	      biteSize = 2500                     'size of our chunks
	      numChunks = totalSize \ biteSize    'number of whole chunks
	      remChunk = totalSize Mod biteSize   'number of bites remaining
	      For i = 0 To numChunks              'go through each chunk
		If i = numChunks Then biteSize = remChunk    'on last chunk?
		chunk = dataLog.Recordset.Fields("Memo").GetChunk(i * biteSize, biteSize) 'eat the chunk
		f.Print chunk                     'and print it out
	      Next i
	      f.Print
	      dataLog.Recordset.MovePrevious      'step backwards in time
	    Loop
       Case 3 '*********************************** Find
	    formFind.Show                         'show the find form
       Case 4 '*********************************** Datasheet
	    MousePointer = HOURGLASS              'we're going to be a while
	    formDatasheet.Show                    'show the data sheet
	    MousePointer = DEFAULT                'ready to roll!
     End Select
End Sub

Sub PrintStats ()
     Dim f As formStats                   'This just shortens up subsequent code
     Set f = formStats
     f.Show                               'show the stats form
     f.Cls                                'clear off any existing text
     f.Print
     f.Print , "Database Filename: "; Tab(40); dataMain.DatabaseName
     f.Print , "Database Size:"; Tab(40); FileLen(dataMain.DatabaseName)
     f.Print , "Database User Status:"; Tab(40); Choose(dataMain.Exclusive + 2, "Single User", "Multi-User")
     f.Print , "Database Access Mode:"; Tab(40); Choose(dataMain.ReadOnly + 2, "Read Only", "Read and Write")
     f.Print , "Database Collating Order"; Tab(40); dataMain.Database.CollatingOrder
     f.Print
     For j = 0 To dataMain.Database.TableDefs.Count - 1
       If (dataMain.Database.TableDefs(j).Attributes And DB_SYSTEMOBJECT) = 0 Then
	    f.Print
	    f.Print "  _______________________________________________________"
	    f.Print
	    f.Print , "Table: "; dataMain.Database.TableDefs(j).Name
	    f.Print
	    f.Print , "Updatable: "; Choose(dataMain.Database.TableDefs(j).Updatable + 2, "All fields updatable", "Some fields not updatable")
	    For i = 0 To dataMain.Database.TableDefs(j).Indexes.Count - 1
	      f.Print , "Index: "; dataMain.Database.TableDefs(j).Indexes(i).Name,
	      f.Print "Fields: "; dataMain.Database.TableDefs(j).Indexes(i).Fields
	    Next i
	    f.Print
	    f.Print , "Field #", "Name", "Type"
	    f.Print , "----------------------------------------------------"
	    For i = 0 To dataMain.Database.TableDefs(j).Fields.Count - 1
	      f.Print , dataMain.Database.TableDefs(j).Fields(i).OrdinalPosition,
	      f.Print dataMain.Database.TableDefs(j).Fields(i).Name,
	      Select Case dataMain.Database.TableDefs(j).Fields(i).Type
		Case 1: f.Print "Boolean",
		Case 2: f.Print "Byte",
		Case 3: f.Print "Integer",
		Case 4: f.Print "Long",
		Case 5: f.Print "Currency",
		Case 6: f.Print "Single",
		Case 7: f.Print "Double",
		Case 8: f.Print "Date",
		Case 9: f.Print "Binary",
		Case 10: f.Print "Text",
		Case 11: f.Print "LongBinary",
		Case 12: f.Print "Memo",
	      End Select
	      f.Print
	    Next i
       End If
     Next j
End Sub

Sub textFileName_KeyPress (KeyAscii As Integer)
     cmndSave.Enabled = True              'user changed entry; enable save button
End Sub

Sub textPathName_KeyPress (KeyAscii As Integer)
     cmndSave.Enabled = True              'user changed entry; enable save button
End Sub

