VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Updating with the data control"
   ClientHeight    =   5340
   ClientLeft      =   2415
   ClientTop       =   2385
   ClientWidth     =   7365
   Height          =   5745
   Left            =   2355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7365
   Top             =   2040
   Width           =   7485
   Begin PictureBox Picture1 
      Height          =   4995
      Left            =   120
      ScaleHeight     =   4965
      ScaleWidth      =   7050
      TabIndex        =   0
      Top             =   195
      Width           =   7080
      Begin CommandButton cmdCancel 
         Caption         =   "&Cancel changes"
         Height          =   450
         Left            =   5385
         TabIndex        =   15
         Top             =   4365
         Width           =   1545
      End
      Begin CommandButton cmdDelete 
         Caption         =   "&Delete Record"
         Height          =   450
         Left            =   2835
         TabIndex        =   14
         Top             =   4380
         Width           =   1545
      End
      Begin CommandButton cmdAdd 
         Caption         =   "&Add new record"
         Height          =   450
         Left            =   120
         TabIndex        =   13
         Top             =   4380
         Width           =   1545
      End
      Begin Data datPublishers 
         Caption         =   "Publishers"
         Connect         =   ""
         DatabaseName    =   "biblio"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   4755
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "Publishers"
         Top             =   4005
         Visible         =   0   'False
         Width           =   2205
      End
      Begin Data datAuthors 
         Caption         =   "Authors"
         Connect         =   ""
         DatabaseName    =   "biblio.mdb"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "Authors"
         Top             =   4005
         Visible         =   0   'False
         Width           =   2070
      End
      Begin Data datTitles 
         Caption         =   "Use the arrows to scroll through the titles in the database"
         Connect         =   ""
         DatabaseName    =   "biblio.mdb"
         Exclusive       =   -1  'True
         Height          =   285
         Left            =   135
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "Titles"
         Top             =   3630
         Width           =   6795
      End
      Begin ComboBox cboPublisher 
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2820
         Width           =   4080
      End
      Begin ComboBox cboAuthor 
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   4080
      End
      Begin TextBox txtISBN 
         BackColor       =   &H0000FFFF&
         DataField       =   "ISBN"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2070
         TabIndex        =   10
         Top             =   1515
         Width           =   2385
      End
      Begin TextBox txtYear 
         BackColor       =   &H0000FFFF&
         DataField       =   "Year Published"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Top             =   870
         Width           =   900
      End
      Begin TextBox txtAuthor 
         DataField       =   "Au_ID"
         DataSource      =   "datTitles"
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Text            =   "These are not visible at run time"
         Top             =   2490
         Visible         =   0   'False
         Width           =   4080
      End
      Begin TextBox txtPublisher 
         DataField       =   "PubID"
         DataSource      =   "datTitles"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Text            =   "These are not visible at run time"
         Top             =   3150
         Visible         =   0   'False
         Width           =   4080
      End
      Begin TextBox txtTitle 
         BackColor       =   &H0000FFFF&
         DataField       =   "Title"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2070
         TabIndex        =   6
         Top             =   225
         Width           =   4080
      End
      Begin Label Label5 
         Caption         =   "Publisher"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   255
         TabIndex        =   5
         Top             =   2775
         Width           =   1845
      End
      Begin Label Label4 
         Caption         =   "Author"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   255
         TabIndex        =   4
         Top             =   2145
         Width           =   1845
      End
      Begin Label Label3 
         Caption         =   "ISBN"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   255
         TabIndex        =   3
         Top             =   1515
         Width           =   1845
      End
      Begin Label Label2 
         Caption         =   "Year Published"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   255
         TabIndex        =   2
         Top             =   885
         Width           =   1845
      End
      Begin Label Label1 
         Caption         =   "Title"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   255
         TabIndex        =   1
         Top             =   255
         Width           =   1845
      End
   End
End
Option Explicit


' Declare a variable to hold True/False values. This will be used to indicate at various stages
' of the program whether or not the user is adding a new record, or simply updating an existing
' one.
Dim miAdding As Integer

' Declare a True/False variable to tell the Validate event when to cancel updates to the database
Dim miCancel As Integer

' Declare a string variable to hold record bookmarks.
Dim sPreviousRecord As String

Sub cboAuthor_Click ()

    ' If the user selects a new author from the combo box, then the invisible bound control
    ' must also be updated
    txtAuthor.Text = cboAuthor.ItemData(cboAuthor.ListIndex)

End Sub

Sub cboPublisher_click ()

    ' If the user selects a different publisher, then the invisible txtPublisher bound control
    ' should be updated so that the changes are stored in the database
    txtPublisher.Text = cboPublisher.ItemData(cboPublisher.ListIndex)

End Sub

Sub cmdAdd_Click ()

    ' If we are currently adding a new record then do an Exit sub straight away
    If miAdding Then Exit Sub

    ' Save the bookmark of the current record into a variable. We will use this later
    ' when cancel is pressed to get back to the this record.
    sPreviousRecord = datTitles.Recordset.Bookmark

    ' Add a new record to the recordset
    datTitles.Recordset.AddNew
    
    ' Set the miAdding variable to true to show the delete and cancel buttons that we are
    ' adding a new record
    miAdding = True

End Sub

Sub cmdCancel_Click ()
    
    If miAdding Then
	
	' If we are adding a new record then set the miCancel variable to True. This is
	' checked in the validate evetn to see if it is OK to write new infroamtion to the
	' database.
	miCancel = True

	' User has selected cancel - jump back to the record they were on before they hit Add
	datTitles.Recordset.Bookmark = sPreviousRecord

	' Set the miAdding variable to False since we are no longer adding a new record
	miAdding = False

    Else
	
	' If the user hits cancel whilst not adding a new record, simply reload the data for
	' the current record from the data control. This will overwrite any changes the user
	' has made.
	datTitles.UpdateControls
    End If

End Sub

Sub cmdDelete_Click ()

    Dim nResponse
    
    If miAdding Then
	
	' If we are adding a new record then the delete button does exactly the same as the
	' cancel button.
	cmdCancel_Click

    Else
	
	' Before we attempt a delete, the program needs to be certain that there is a current
	' record available to delete.
	If datTitles.Recordset.EOF Or datTitles.Recordset.BOF Then Exit Sub
    
	nResponse = MsgBox("Do you really want to delete this record?", 20, "Delete Record")
	If nResponse = 6 Then
	    ' Delete the current record
	    datTitles.Recordset.Delete

	    ' After the current record has been deleted, there is no current record. We need to
	    ' use a move method to jump to a valid record
	    datTitles.Recordset.MoveFirst

	End If

    End If

End Sub

Sub datTitles_Reposition ()

    ' Whenever a new record is found, we need to set miAdding to False to show all the other
    ' buttons that a new record is no longer being added
    miAdding = False

End Sub

Sub datTitles_Validate (action As Integer, Save As Integer)

    ' When validate tries to save something to the database we need to check that the
    ' author and publisher fields are valid. When a new record is created for instance
    ' these are loaded in as blank. If the user then decides not to select a valid
    ' publisher or author with the combo boxes then they get saved out as blank as well.

    If Save And miCancel Then Save = False

End Sub

Sub Display_Author ()
    
    ' This procedure displays the authors name based on the author's id in txtAuthor

    ' Declare a variable to be used to loop through the entries in the combo box.
    Dim nListIndex As Integer

    ' If the bound control holding the authors ID is blank, that means we are dealing with
    ' a new record. Select the first item in the combo box and this will automatically update
    ' the bound control - because of our code in cboAuthor_Click
    If txtAuthor.Text = "" Then
	cboAuthor.ListIndex = 0
	Exit Sub
    End If

    ' Loop through the items in the combo box list until one is found which matches the
    ' value in the bound control.
    nListIndex = 0
    Do While cboAuthor.ItemData(nListIndex) <> Val(txtAuthor.Text)

	nListIndex = nListIndex + 1

    Loop
    cboAuthor.ListIndex = nListIndex

End Sub

Sub Display_Publisher ()

    ' This procedure displays the publisher's name based on the publisher id in txtPublisher

    ' Declare a variable to be used to loop through the entries in the combo box.
    Dim nListIndex As Integer

    ' If txtPublisher (the bound control holding the publisher's ID) is blank then we are dealing
    ' with a new record and need to set the field to something. Selecting the first item in the
    ' combo box should do the trick since this will trigger our cboPublisher_click event, which
    ' will in turn update the text box.
    If txtPublisher.Text = "" Then
	cboPublisher.ListIndex = 0
	Exit Sub
    End If

    ' Loop through the publishers in the list until one is found matching the one displayed
    ' in the text box. When that one is found select it, which will cause the publishers
    ' name to appear in the edit portion of the combo box.
7
    nListIndex = 0
    Do While cboPublisher.ItemData(nListIndex) <> Val(txtPublisher.Text)

	nListIndex = nListIndex + 1

    Loop
    cboPublisher.ListIndex = nListIndex

End Sub

Sub Form_Load ()

    ' At form load we need to make sure that the databasename properties of the three combo boxes
    ' are all valid. We can assume that the database is in the same directory as the program
    ' itself so the databasename property can be set using App.Path

    Dim sDatabasename As String


    ' Change the mousepointer into an hourglass
    screen.MousePointer = 11

    ' Build the databasename
    sDatabasename = app.Path & "\" & "biblio.mdb"

    ' If the database cannot be found at the above location then stop the program now.
    If Dir$(sDatabasename) = "" Then
	MsgBox "Database could not be found, please make sure that Biblio.MDB " & Chr$(10) & "is in the same location as Biblio.Mak", , "Database Error"
	End
    End If

    
    ' Set the database name property of all the data controls on the form.
    datTitles.DatabaseName = sDatabasename
    datPublishers.DatabaseName = sDatabasename
    datTitles.DatabaseName = sDatabasename
    
    ' Now that we have set the databasename properties of all the data controls we need to
    ' refresh them . Normally a refresh will take place at the end of the form load event but we
    ' need access to the data controls now, before the end of the form load
    datPublishers.Refresh
    datAuthors.Refresh
    ' The titles data control is refreshed last, at the end of the event

    frmMain.Show
    DoEvents

    ' Now we need to add items to the two combo boxes for the authors and publishers
    
    ' Find the first author record
    datAuthors.Recordset.MoveFirst
    ' Loop through the authors recordset
    Do While Not datAuthors.Recordset.EOF

	' Add the current author name to the combo box
	cboAuthor.AddItem datAuthors.Recordset.Fields("Author")

	' Store the author ID in the combo box ItemData array
	cboAuthor.ItemData(cboAuthor.ListCount - 1) = datAuthors.Recordset.Fields("Au_ID")

	' Move to the next author
	datAuthors.Recordset.MoveNext

    Loop

    ' Find the first publisher record
    datPublishers.Recordset.MoveFirst
    ' Loop through the publishers recordset
    Do While Not datPublishers.Recordset.EOF

	' Add the current publisher name to the combo box
	cboPublisher.AddItem datPublishers.Recordset.Fields("Name")
	
	' Store the publisherid in the combo box Itemdata array
	cboPublisher.ItemData(cboPublisher.ListCount - 1) = datPublishers.Recordset.Fields("PubID")

	' Move to the next publisher
	datPublishers.Recordset.MoveNext

    Loop

    ' Now we can update the Titles data control. This will change the txtAuthor
    ' and txtPublisher controls causing them to display the appropriate entry from the two
    ' combo boxes. This can only be done after the combos have been built which is why the
    ' datTitles.Refresh is put last
    datTitles.Refresh

    ' Change the mousepointer back to normal.
    screen.MousePointer = 0
    

    miAdding = False
    miCancel = False

End Sub

Sub txtAuthor_Change ()

    ' If the value in the bound txtAuthor control changes, we need to reflect the change in the
    ' the combo box
    Call Display_Author

End Sub

Sub txtPublisher_Change ()

    ' If the value in the bound txtPublisher control changes, we need to reflect the change in
    ' the combo box
    Call Display_Publisher

End Sub

