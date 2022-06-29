VERSION 2.00
Begin Form formMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OLE Project"
   ClientHeight    =   4668
   ClientLeft      =   1212
   ClientTop       =   2088
   ClientWidth     =   5940
   Height          =   5436
   Icon            =   CHP26A.FRX:0000
   Left            =   1152
   LinkTopic       =   "Form1"
   ScaleHeight     =   4668
   ScaleWidth      =   5940
   Top             =   1380
   Width           =   6060
   Begin CommandButton cmndGraph 
      Caption         =   "Create &Graph"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1332
   End
   Begin CommandButton cmndReport 
      Caption         =   "Create &Report"
      Height          =   372
      Left            =   1620
      TabIndex        =   3
      Top             =   4200
      Width           =   1332
   End
   Begin OLE OLE 
      BorderStyle     =   0  'None
      fFFHk           =   -1  'True
      Height          =   2532
      Index           =   0
      Left            =   3240
      SizeMode        =   2  'AutoSize
      TabIndex        =   2
      Top             =   1560
      Width           =   2532
   End
   Begin OLE OLE 
      fFFHk           =   -1  'True
      Height          =   2532
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2772
   End
   Begin OLE OLE 
      AutoActivate    =   1  'GetFocus
      AutoVerbMenu    =   0   'False
      Class           =   "Excel.Sheet.5"
      fFFHk           =   -1  'True
      Height          =   1332
      Index           =   2
      Left            =   120
      OleObjectBlob   =   CHP26A.FRX:0302
      SourceDoc       =   "c:\waite\vbsb\chp26obj.xls"
      TabIndex        =   0
      Top             =   120
      Width           =   5652
   End
   Begin Menu menuBar 
      Caption         =   "&File"
      Index           =   0
      Begin Menu menuFile 
         Caption         =   "&Open Objects"
         Index           =   0
      End
      Begin Menu menuFile 
         Caption         =   "&Save Objects"
         Index           =   1
      End
      Begin Menu menuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin Menu menuFile 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin Menu menuBar 
      Caption         =   "&Edit"
      Index           =   1
      Begin Menu menuEdit 
         Caption         =   "Cu&t"
         Index           =   0
      End
      Begin Menu menuEdit 
         Caption         =   "&Copy"
         Index           =   1
      End
      Begin Menu menuEdit 
         Caption         =   "&Paste"
         Index           =   2
      End
      Begin Menu menuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin Menu menuEdit 
         Caption         =   "Paste &Special"
         Index           =   4
      End
      Begin Menu menuEdit 
         Caption         =   "Insert &Object"
         Index           =   5
      End
      Begin Menu menuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin Menu menuEdit 
         Caption         =   "&Info"
         Index           =   7
      End
   End
   Begin Menu menuDOC 
      Caption         =   "(POPUP)"
      Visible         =   0   'False
      Begin Menu menuDOCPop 
         Caption         =   "&Generate"
         Index           =   0
      End
      Begin Menu menuDOCPop 
         Caption         =   "&Copy"
         Index           =   1
      End
      Begin Menu menuDOCPop 
         Caption         =   "-"
         Index           =   2
      End
   End
End
Const DOC = 0
Const GRAPH = 1
Const SHEET = 2

Const OLE_CREATE_EMBED = 0
Const OLE_COPY = 4
Const OLE_PASTE = 5
Const OLE_UPDATE = 6
Const OLE_ACTIVATE = 7
Const OLE_CLOSE = 9
Const OLE_DELETE = 10
Const OLE_SAVE_TO_FILE = 11
Const OLE_READ_FROM_FILE = 12
Const OLE_INSERT_OBJ_DLG = 14
Const OLE_PASTE_SPECIAL_DLG = 15
Const OLE_FETCH_VERBS = 17
Const OLE_SAVE_TO_OLE1FILE = 18

Const OLE_DISPLAY_CONTENT = 0

Const OLE_EMBEDDED = 1

Const VERB_PRIMARY = 0
Const VERB_SHOW = -1
Const VERB_OPEN = -2
Const VERB_HIDE = -3
Const VERB_INPLACEUIACTIVATE = -4
Const VERB_INPLACEACTIVATE = -5

Const VF_ENABLED = &H0
Const VF_GRAYED = &H1
Const VF_DISABLED = &H2
Const VF_CHECKED = &H8
Const VF_SEPARATOR = &H800

Const DEFAULT = 0
Const HOURGLASS = 11

Dim oleChanged(3) As Integer

Sub ActivateObject (Object As Control)

'fancy way of activating an object; tries several times to activate it.

'most applications respond fast enough that you don't need this routine; you could
'just do OLE.ACTION = OLE_ACTIVATE and be fine. The beta copy of Excel 5.0 was WAY slow
'and returned error 31027 (object not responding). The production version will probably
'be tuned better for speed, and may not need this.

    Static attempts As Integer                  'remember how many times we've tried this before
    Const MAXTRIES = 20                         'maximum number of attempts
    Const RETRYWAIT = 1000                      'how long to let doevents run
    
    If TypeOf Object Is OLE Then                'make sure we're on an OLE object!
	On Error GoTo failedToActivate          'in case it times out
	Object.Action = OLE_ACTIVATE            'try to activate
	attempts = 0                            'success!
    End If
Exit Sub

failedToActivate:
    Debug.Print Err                             'double check the error!
    If Err <> 31027 Then                        'if it's other than expected,
	    MsgBox Error$(Err)                  'print out error message
	Exit Sub
    Else                                        'ah, the expected "object not responding" error
	attempts = attempts + 1                 'remember theat we've tried this
	If attempts < MAXTRIES Then             'and if it's still less than the max,
	    For i = 0 To RETRYWAIT
		DoEvents                        'give the system some time to catch up
	    Next i
	    Resume                              'and try it again...
	Else
	    MsgBox "Failed to open OLE host"    'too many times, so admit defeat
	    attempts = 0                        'reset for next time
	    Exit Sub
	End If
    End If

End Sub

Sub cmndGraph_Click ()
     
'This routine relies on having Excel 5.0 or greater and MSGraph 1.0 or greater.
'It will bomb unless you have both applications. The general techniques will
'work for ole automation, but any line commented with ( >> ) is host-application specific

     OLE(SHEET).Verb = VERB_HIDE                        'keep app hidden
     ActivateObject OLE(SHEET)                          'activate host
     OLE(SHEET).Object.Range("QuarterlyFigures").Copy   ' >> copy to clipboard
     OLE(SHEET).Action = OLE_CLOSE                      'close host
     AppActivate "OLE Project"                          'back to the project
     
     quarterlyFigures = Clipboard.GetText()             'contents of clipboard
     
     OLE(GRAPH).Verb = VERB_HIDE                        'keep app hidden
     OLE(GRAPH).Format = "CF_TEXT"                      'we're sending text
     OLE(GRAPH).Class = "MSGraph"                       'graph
     OLE(GRAPH).Action = OLE_CREATE_EMBED               'embedded object
     ActivateObject OLE(GRAPH)                          'activate host
     OLE(GRAPH).DataText = quarterlyFigures             'send it the figures
     OLE(GRAPH).Action = OLE_UPDATE                     'and update the display

End Sub

Sub cmndReport_Click ()
     
'This routine relies on having Excel 5.0 or greater and Winword 6.0 or greater.
'It will bomb unless you have both applications. The general techniques will
'work for ole automation, but any line commented with ( >> ) is host-application 'specific

     MousePointer = HOURGLASS                       'going to be awhile
     ReDim q(4) As Long                             'four quarter's figures
     Dim wordProc As Object                         'for word processor
     Dim avg1 As Long, avg2 As Long                 'quarterly averages
     Static createOrGet As Integer                  'two ways of automation

     OLE(SHEET).Verb = VERB_HIDE                    'keep app hidden
     ActivateObject OLE(SHEET)                          'activate host
     For i = 1 To 4                                 'for each of the 4 quarters,
	 q(i) = OLE(SHEET).Object.Range("SalesFigures").Cells(1, i).Value
						    ' >> get the sales total
     Next i
     OLE(SHEET).Action = OLE_CLOSE                  'close the host
     AppActivate "OLE Project"                      'the project

     avg1 = (q(1) + q(2)) / 2                       'first half of year
     avg2 = (q(3) + q(4)) / 2                       'second half of year
     textToSet = CreateSpiel(avg1, avg2)            'make up the sales talk

'you can use either CreateObject or GetObject for OLE automation. This block
'alternates between the two methods. CreateObject creates an automation object
'out of "thin air", while GetObject uses an existing word processor document
'as a "template". Note the commented-out line 'createOrGet = True'; Word6.0 doesn't
'support "getting" an object.
     
     If createOrGet = False Then                    'use CreateObject?
	Set wordProc = CreateObject("Word.Basic")   'create the object
	wordProc.filenew                            ' >> create a new document
	'NOTE: Word 6.0 doesn't support "getting" an object, it only supports
	'creating a WordBasic object.  Remove the comment on this next line if
	'you want to experiment with another object as an OLE server
	'createOrGet = True                         'flip the flag for next time
     Else
	Set wordProc = GetObject(App.Path & "\CHP26BLK.DOC")  'blank template
	createOrGet = False                         'flip the flag for next time
     End If
     DoEvents                                       'give the system some time
     wordProc.FilePageSetup 0, 0, .2, .2, .2, .2, 0, "1.6 in", "10 in"
						    ' >> set document margins
     wordProc.Insert textToSet                      ' >> sales talk
     wordProc.EditSelectAll                         ' >> select entire document
     wordProc.FormatParagraph .1, .1, .1, 8, 0, 0, 3  ' >> paragraph spacing
     wordProc.FormatFont 8                          ' >> set font to 8 point
     wordProc.StartOfDocument                       ' >> goto beginning
     wordProc.FormatDropCap 1, , 2, 1               ' >> format first letter
     wordProc.ParaDown 1                            ' >> goto next paragraph
     wordProc.FormatDropCap 1, , 2, 1               ' >> format first letter
     wordProc.FileSaveAs App.Path & "\oledoc.dat"   ' >> save document
     wordProc.FileCloseAll 2                        ' >> close documents
     
     OLE(DOC).Class = "Word.Document"               'set class to Winword 6.0
     OLE(DOC).OLETypeAllowed = OLE_EITHER           'either linked or embedded
     OLE(DOC).SourceDoc = App.Path & "\oledoc.dat"  'the document we just saved
     OLE(DOC).Action = OLE_CREATE_EMBED             'an embedded document
     MousePointer = DEFAULT                         'all done!

End Sub

Function CreateSpiel (a1 As Long, a2 As Long) As String
     
'create a sales talk, with different verbage dependent on last year's results
     Select Case Sgn(a2 - a1)       'down, even, or upward trend?
	  Case -1
	       results = "Although sales declined overall for the year, "
	       results = results & "we anticipate phenomonal growth next year!"
	  Case 0
	       results = "Sales were relatively flat throughout the year. "
	       results = results & "The strengthening economy will help us set "
	       results = results & "new productivity records next year!"
	  Case 1
	       results = "Sales were up, up, up! We had a fantastic year, with "
	       results = results & "even better prospects for next year!"
     End Select
     spiel = "Our five year goals have been met. "
     spiel = spiel & "PSA has experienced tremendous growth over the past "
     spiel = spiel & "five years, and our higher level of partnership with "
     spiel = spiel & "Puryear Enterprises sets us up for even higher levels "
     spiel = spiel & "of growth in the future!" & Chr$(13)
     spiel = spiel & "Set your sights high! The past year's sales results "
     spiel = spiel & "are quite encouraging. "
     spiel = spiel & results

     CreateSpiel = spiel

End Function

Sub Form_Load ()
     
     For i = DOC To SHEET                               'for each OLE container,
	  OLE(i).HostName = "OLE Project"               'set the host name
	  OLE(i).DisplayType = OLE_DISPLAY_CONTENT      'and display type
     Next i
     OLE(GRAPH).AutoVerbMenu = True
     OLE(SHEET).AutoVerbMenu = True
     OLE(DOC).AutoVerbMenu = False

End Sub

Sub menuBar_Click (Index As Integer)
     
     Select Case Index
	  Case 0 '*********** File
	       menuFile(1).Enabled = False              'turn off save
	       For i = DOC To SHEET                     'go through containers
		   If oleChanged(i) = True Then         'if it has changed,
			menuFile(1).Enabled = True      'then enable File Save
		   End If
	       Next i
	  Case 1 '*********** Edit
	       For i = 0 To 7                           'iterate through Edit
							'menu commands (except
							'separator bars)
		    If (i <> 3 And i <> 6) Then menuEdit(i).Enabled = False
							'and turn them off
	       Next i
	       Dim s As Control
	       Set s = Screen.ActiveControl             'control that has focus
	       If TypeOf s Is OLE Then                  'if it's OLE control,
		    menuEdit(5).Enabled = True          'enable Insert Object
		    If s.OLEType <> 3 Then              'if it's got an object
			 menuEdit(0).Enabled = True     'enable Cut
			 menuEdit(1).Enabled = True     'enable Copy
			 menuEdit(7).Enabled = True     'enable Info
		    End If
		    If s.PasteOK Then                   'if the clipboard has a
							'pastable object
			 menuEdit(2).Enabled = True     'enable Paste
			 menuEdit(4).Enabled = True     'enable Paste Special
		    End If
	       End If
     End Select

End Sub

Sub menuDOCPop_Click (Index As Integer)

    On Error GoTo noMoreItems               'no more menu items (verbs)
    i = 3                                   'custom menu items begin after
					    'separator bar
    Do
	Unload menuDOCPop(i)                'get rid of custom menu item
	i = i + 1                           'and go to next one
    Loop

noMoreItems:                                'ok, we've zapped all the verbs!
    On Error GoTo 0
	
    Select Case Index                       'choose the correct menu item
	Case 0 ' ------- Generate
	    cmndReport_Click                'generate the sales talk again
	Case 1 ' ------- Copy
	    menuEdit_Click 1                'copy the sales talk to the clipboard
	Case 2 ' ------- Separator bar
	Case Is > 2                         'if it was a custom item (that is,
					    'an object verb)
	    OLE(DOC).Verb = Index - 2       'set the verb appropriately
	    ActivateObject OLE(DOC)         'and perform the verb
    End Select
    Exit Sub

End Sub

Sub menuEdit_Click (Index As Integer)
     
     Dim s As Control
     Set s = Screen.ActiveControl                   'control that has focus
     formMain.MousePointer = HOURGLASS              'we're going to be awhile
     Select Case Index
	  Case 0 '************ Cut
	       If Not s.AppIsRunning Then           'if host app isn't running,
		    s.Verb = VERB_HIDE              'make sure it stays hidden
		    ActivateObject s                'and start it up
	       End If
	       s.Action = OLE_COPY                  'copy object to clipboard
	       s.Action = OLE_DELETE                'and get rid of it
	  Case 1 '************ Copy
	       If Not s.AppIsRunning Then           'if host app isn't running,
		    s.Verb = VERB_HIDE              'make sure it stays hidden
		    ActivateObject s                'and start it up
	       End If
	       s.Action = OLE_COPY                  'copy object to clipboard
	  Case 2 '************ Paste
	       s.Class = "temp"                     'if ole container has never
						    'been used, class needs to be
						    'set to something (but
						    'doesn't really matter what)
	       s.Action = OLE_PASTE                 'paste the contents of
						    'clipboard into container
	  Case 3 '************ (separator bar)
	  Case 4 '************ Paste Special
	       s.Action = OLE_PASTE_SPECIAL_DLG     'open paste special dialog
	  Case 5 '************ Insert Object
	       s.Action = OLE_INSERT_OBJ_DLG        'open insert object dialog
	  Case 6 ' *********** Separator Bar
	  Case 7 ' *********** Info
	       printInfo s                          'print info about object
     End Select

     If s.OLEType <> 3 Then                         'object in the ole container?
	  s.Action = OLE_CLOSE                      'close host
	  s.Refresh                                 'and update display
     End If
     formMain.MousePointer = DEFAULT                'all done!
     
End Sub

Sub menuFile_Click (Index As Integer)
     
     Select Case Index
	  Case 0 ' *********** Open objects
	       fileNum = FreeFile                   'get a file handle
	       Open App.Path & "\oleobject.dat" For Binary As fileNum   'open
	       For i = DOC To SHEET                 'iterate through containers
		    OLE(i).FileNumber = fileNum     'set container's file handle
		    If OLE(i).OLEType <> 3 Then     'if it's got an object in it,
			 OLE(i).Action = OLE_READ_FROM_FILE 'get object from file
		    End If
		    oleChanged(i) = False           'disable saving
	       Next i
	  Close fileNum                             'close the file
	  Case 1 ' *********** Save objects
	       fileNum = FreeFile                    'get a file handle
	       Open App.Path & "\oleobject.dat" For Binary As fileNum   'open
	       For i = DOC To SHEET                  'iterate through containers
		    OLE(i).FileNumber = fileNum      'set the container's file
		    If OLE(i).OLEType <> 3 Then      'if it's got an object
			 OLE(i).Action = OLE_SAVE_TO_FILE   'save the object
		    End If
		    oleChanged(i) = False            'disable saving
	       Next i
	       Close fileNum                         'close the file
	       Case 2 ' *********** (separator bar)
	  Case 3
	       End
     End Select

End Sub

Sub OLE_MouseDown (Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And Index = DOC And OLE(DOC).OLEType <> 3 Then
				   'right button, on doc, with an object in it:
	OLE(DOC).Action = OLE_FETCH_VERBS           'update the available verbs
	For i = 1 To OLE(DOC).ObjectVerbsCount - 1  'iterate through each verb
	    Load menuDOCPop(2 + i)                  'add a new menu item for verb
	    menuDOCPop(2 + i).Caption = OLE(DOC).ObjectVerbs(i) 'and add the verb
	Next i
	PopupMenu menuDOC                           'pop up the menu
    End If

End Sub

Sub OLE_Resize (Index As Integer, HeightNew As Single, WidthNew As Single)

    If Index = DOC Then                             'just for the sales talk,
	If HeightNew > 3100 Then HeightNew = 3100   'don't exceed these
	If WidthNew > 2532 Then WidthNew = 2532     'maximum dimensions
    End If

End Sub

Sub OLE_Updated (Index As Integer, Code As Integer)

    If Code = 0 Then oleChanged(Code) = True        'if object has changed, then
						    'flag it to enable saving

End Sub

Sub printInfo (s As Control)
    
    MousePointer = HOURGLASS
    formInfo.Show                                    'show info form
    formInfo!pictInfo.Cls                                     'wipe it clean
    If Not s.AppIsRunning Then                       'if the host isn't running
	s.Verb = VERB_HIDE                           'make sure it stays hidden
	ActivateObject s                             'and start it up
    End If
    formInfo!pictInfo.Print
    formInfo!pictInfo.Print "  Class:";
    formInfo!pictInfo.Print Tab(20); s.Class                  'host class
    formInfo!pictInfo.Print
    formInfo!pictInfo.Print "  Accept Formats:";              'host accepted formats,
    For i = 0 To s.ObjectAcceptFormatsCount - 1      'iterating through each one
	formInfo!pictInfo.Print Tab(20); s.ObjectAcceptFormats(i)
    Next i
    formInfo!pictInfo.Print
    formInfo!pictInfo.Print "  Return Formats:";              'host returned formats,
    For i = 0 To s.ObjectGetFormatsCount - 1         'iterating through each one
	formInfo!pictInfo.Print Tab(20); s.ObjectGetFormats(i)
    Next i
    formInfo!pictInfo.Print
    formInfo!pictInfo.Print "  Verbs:";                       'host verbs,
    For i = 1 To s.ObjectVerbsCount - 1              'iterating through each one
	formInfo!pictInfo.Print Tab(20); s.ObjectVerbs(i); Tab(40);
	Select Case s.ObjectVerbFlags(i)             'and print the menu state
	    Case VF_CHECKED: formInfo!pictInfo.Print "Checked"
	    Case VF_DISABLED: formInfo!pictInfo.Print "Disabled"
	    Case VF_ENABLED: formInfo!pictInfo.Print "Enabled"
	    Case VF_GRAYED: formInfo!pictInfo.Print "Grayed"
	    Case VF_SEPARATOR: formInfo!pictInfo.Print "Separator"
	End Select
    Next i
    
    formInfo!pictInfo.Print
    formInfo!pictInfo.Print "  Long Pointer: ";               'memory position of object
    formInfo!pictInfo.Print Tab(20); Hex$(s.LpOleObject)      '(could use with API calls)
    
    formInfo!pictInfo.Print
    s.Format = "Native"
    formInfo!pictInfo.Print "  Data:";                        'print the GDI handle
    formInfo!pictInfo.Print Tab(20); Hex$(s.Data)
    
End Sub

