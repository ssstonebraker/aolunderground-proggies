VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DDE Project"
   ClientHeight    =   5112
   ClientLeft      =   1200
   ClientTop       =   1512
   ClientWidth     =   6912
   Height          =   5880
   Icon            =   CHP25.FRX:0000
   Left            =   1140
   LinkMode        =   1  'Source
   LinkTopic       =   "formMain"
   ScaleHeight     =   5112
   ScaleWidth      =   6912
   Top             =   804
   Width           =   7032
   Begin TextBox textCogs 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   3
      Left            =   5520
      TabIndex        =   17
      Text            =   "Text9"
      Top             =   4320
      Width           =   1212
   End
   Begin TextBox textCogs 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   2
      Left            =   4200
      TabIndex        =   16
      Text            =   "Text8"
      Top             =   4320
      Width           =   1212
   End
   Begin TextBox textCogs 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   1
      Left            =   2880
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   4320
      Width           =   1212
   End
   Begin TextBox textCogs 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   0
      Left            =   1560
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   4320
      Width           =   1212
   End
   Begin TextBox textSales 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   3
      Left            =   5520
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   3960
      Width           =   1212
   End
   Begin TextBox textSales 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   2
      Left            =   4200
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   3960
      Width           =   1212
   End
   Begin TextBox textSales 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   1
      Left            =   2880
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   3960
      Width           =   1212
   End
   Begin TextBox textSales 
      Alignment       =   1  'Right Justify
      Height          =   288
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3960
      Width           =   1212
   End
   Begin TextBox textBox 
      Height          =   972
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   240
      Width           =   5172
   End
   Begin PictureBox pictGraph 
      Height          =   2532
      Left            =   1560
      ScaleHeight     =   2508
      ScaleWidth      =   5148
      TabIndex        =   4
      Top             =   1320
      Width           =   5172
   End
   Begin Label lablTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label11"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   5400
      TabIndex        =   20
      Top             =   4680
      Width           =   1332
   End
   Begin Label lablTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label10"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   19
      Top             =   4680
      Width           =   1332
   End
   Begin Label lablTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   2760
      TabIndex        =   18
      Top             =   4680
      Width           =   1332
   End
   Begin Label lablTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label8"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   4680
      Width           =   1212
   End
   Begin Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   1332
   End
   Begin Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Of Goods"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   1332
   End
   Begin Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1332
   End
   Begin Label lablGraphLinkStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1332
   End
   Begin Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Graph"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1332
   End
   Begin Label lablTextLinkStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1332
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Miscellaneous Text"
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1332
      WordWrap        =   -1  'True
   End
   Begin Menu menuBar 
      Caption         =   "&File"
      Index           =   0
      Begin Menu menuFile 
         Caption         =   "E&xit"
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
         Caption         =   "Paste &Link"
         Index           =   3
      End
   End
   Begin Menu menuBar 
      Caption         =   "&Link"
      Index           =   2
      Begin Menu menuLink 
         Caption         =   "&Automatic"
         Index           =   0
      End
      Begin Menu menuLink 
         Caption         =   "&Manual"
         Index           =   1
      End
      Begin Menu menuLink 
         Caption         =   "&Notify"
         Index           =   2
      End
      Begin Menu menuLink 
         Caption         =   "-"
         Index           =   3
      End
      Begin Menu menuLink 
         Caption         =   "&Update Links"
         Index           =   4
      End
   End
End

' Clipboard formats
Const CF_LINK = &HBF00
Const CF_TEXT = 1
Const CF_BITMAP = 2
Const CF_METAFILE = 3
Const CF_DIB = 8
Const CF_PALETTE = 9
' Link constants
Const NONE = 0
Const LINK_AUTOMATIC = 1
Const LINK_MANUAL = 2
Const LINK_NOTIFY = 3

' link error
Const WRONG_FORMAT = 1
Const DDE_SOURCE_CLOSED = 6
Const TOO_MANY_LINKS = 7
Const DATA_TRANSFER_FAILED = 8

' MousePointer
Const HOURGLASS = 11
Const DEFAULT = 0

Const THISAPPNAME = "CHP25"             'used when establishing links - application name
Const TOPICNAME = "formMain"            'used when establishing links - topic name

Dim MenuLinkModeCheck As Integer        'contains current linkmode setting on the link menu

Sub describeLinkError (linkerror, where)
     Select Case linkerror                              'describe the link error
       Case WRONG_FORMAT
            msg$ = "The linked application asked for data in an incompatible format when using the " & where
       Case DDE_SOURCE_CLOSED
            msg$ = "An application attempted a DDE operation after this application disabled its source abilities."
       Case TOO_MANY_LINKS
            msg$ = "Attempt to open too many DDE conversations - 128 maximum!"
       Case DATA_TRANSFER_FAILED
            msg$ = "Failure to update control " & where
       Case 11
            msg$ = "Not enough remaining memory to complete DDE link. Close applications before trying again"
     End Select
     MsgBox msg$
End Sub

Sub Form_LinkError (LinkErr As Integer)
     describeLinkError LinkErr, " the form."            'display error message
End Sub

Sub Form_LinkExecute (CmdStr As String, Cancel As Integer)
     Cancel = False                                  'accept the command
     Select Case CmdStr
       Case "Update Graph"                           'other app wants the graph again
            pictGraph.LinkSend                       'give them the graph
       Case "Sever Graph Link"                       'other app doesn't want graph at all
            pictGraph.LinkMode = NONE                'cut the link
       Case "End"                                    'other app wants this program to end
            End                                      'end
       Case Else
              Cancel = True                          'oops! didn't recognize command
     End Select
End Sub

Sub Form_Load ()
     Screen.MousePointer = HOURGLASS
     TextBox.Text = "Type some text here to link to the other app"
     menuLink_Click 0                                   'Set default to Automatic link
     TextBox.Tag = "textBox"                            'tags used in links
     pictGraph.Tag = "pictGraph"
     lablTextLinkStatus = "(unlinked)"
     
     TB = Chr$(9)                                       'tab
     CR = Chr$(13) & Chr$(10)                           'carriage return/linefeed

                                   'quick little spreadsheet for excel;
                                   'tab and carriage return delimited for each cell and row
     s = " " & TB & "Q1" & TB & "Q2" & TB & "Q3" & TB & "Q4" & CR
     s = s & "Gross Sales" & CR
     s = s & "COGS" & CR
     s = s & "Net Sales" & TB & "= SUM(B2 - B3)" & TB & "= SUM(C2 - C3)" & TB & "= SUM(D2 - D3)" & TB & "= SUM(E2 - E3)"
     Clipboard.SetText s                               'put spreadsheet data on clipboard
   
Startup:
     On Error GoTo OpenExcel                           'set trap
     textSales(0).LinkTopic = "Excel|Sheet1"           'blank worksheet
     textSales(0).LinkItem = "R1C1"                    'arbitrary; doesn't matter what cell
     textSales(0).LinkMode = LINK_MANUAL               'set up link
     textSales(0).LinkExecute "[ACTIVATE(""Sheet1"")]"     'make sure  on the worksheet
     textSales(0).LinkExecute "[SELECT(""R1C1"")]"         'make Excel select upper left
     textSales(0).LinkExecute "[PASTE]"                    'make Excel paste in clipboard
     textSales(0).LinkExecute "[SELECT(""R1C1:R4C5"")]"    'make Excel select worksheet
     textSales(0).LinkExecute "[COLUMN.WIDTH(15)]"         'make Excel set column widths
     textSales(0).LinkExecute "[SELECT(""R2C2:R4C5"")]"    'make Excel select worksheet
     textSales(0).LinkExecute "[FORMAT.NUMBER(""$#,##0_);($#,##0)"")]"  'currency
     textSales(0).LinkExecute "[SELECT(""R1C1:R4C5"")]"    'make Excel select worksheet
     textSales(0).LinkExecute "[NEW(2,1)]"             'make Excel create a new chart item
                                                   'make Excel format chart as a line chart
     textSales(0).LinkExecute "[FORMAT.MAIN(4,1,,,FALSE,FALSE,FALSE)]"
     textSales(0).LinkExecute "[LEGEND(TRUE)]"         'make Excel give it a legend
     textSales(0).LinkExecute "[SELECT(""Chart"")]"    'make Excel select entire chart
                                       'make Excel format all text on chart as Arial 8 point
     textSales(0).LinkExecute "[FORMAT.FONT(0,1,FALSE,""Arial"",8,FALSE,FALSE,FALSE,FALSE)]"
     For i = 0 To 3                                 'iterate through all members
       textSales(i).LinkTopic = "Excel|Sheet1"      'set their link topics
       textCogs(i).LinkTopic = "Excel|Sheet1"
       lablTotal(i).LinkTopic = "Excel|Sheet1"
       textSales(i).LinkItem = "R2C" & Trim$(Str$(i + 2))          'R2C2, R2C3, R2C4, R2C5
       textCogs(i).LinkItem = "R3C" & Trim$(Str$(i + 2))           'R3C2, R3C3, R3C4, R3C5
       lablTotal(i).LinkItem = "R4C" & Trim$(Str$(i + 2))          'R4C2, R4C3, R4C4, R4C5
       textSales(i).LinkMode = LINK_MANUAL          'establish links
       textCogs(i).LinkMode = LINK_MANUAL
       lablTotal(i).LinkMode = LINK_MANUAL
       textSales(i).Text = Str$((i + 4) * 10000)    'dummy up some data
       textCogs(i).Text = Str$((i + 2) * 8345)      'more dummy data
     Next i
     pictGraph.LinkTimeout = 10000                  'graphics could take a while
     pictGraph.LinkTopic = "Excel|Chart1"           'topic is the new chart
     pictGraph.LinkMode = LINK_NOTIFY               'set up notify link
     pictGraph.LinkRequest                          'immediately update the chart
     lablGraphLinkStatus = "Notify Link"            'set label correctly
     Screen.MousePointer = DEFAULT                  'and we're done
Exit Sub

OpenExcel:                                          'error trap
     If Err = 282 Then                              'if it's not loaded,
       If AlreadyStarted = False Then               'and not waiting for it to finish loading
            On Error GoTo NoExcel                   'if it still doesn't work, then there's no excel
            X = Shell("EXCEL.EXE", 2)               'then load it!
            AlreadyStarted = True                   'make sure we don't try loading again
       End If
       DoEvents                                     'give it some time to load
       Resume Startup                               'and try the link again
     Else
       Error Err                                    'woops! something else went wrong....
       Exit Sub
     End If

NoExcel:
    MsgBox "Cannot find Excel on your system - failed to start."
    Exit Sub

End Sub

Sub Form_Unload (Cancel As Integer)
     For i = 0 To 3                                    'iterate through each member
       textSales(i).LinkMode = NONE                    'sever the links
       textCogs(i).LinkMode = NONE
       lablTotal(i).LinkMode = NONE
     Next i
     pictGraph.LinkMode = NONE                         'deactivate any existing links
     pictGraph.LinkTopic = "Excel|System"              'talk to excel
     pictGraph.LinkMode = LINK_MANUAL                  'set up the link
     pictGraph.LinkExecute "[EXIT]"                    'and tell excel to exit
     DoEvents                                          'give it some time to respond
     TextBox.LinkMode = NONE                           'sever link
End Sub

Sub labelLinkStatus ()
     Select Case pictGraph.LinkMode                    'update the graph label
       Case LINK_AUTOMATIC
            lablGraphLinkStatus.Caption = "Auto Link"
       Case LINK_MANUAL
            lablGraphLinkStatus.Caption = "Manual Link"
       Case LINK_NOTIFY
            lablGraphLinkStatus.Caption = "Notify Link"
       Case Else
            lablGraphLinkStatus.Caption = "(unlinked)"
     End Select
     Select Case TextBox.LinkMode                      'update the text box label
       Case LINK_AUTOMATIC
            lablTextLinkStatus.Caption = "Auto Link"
       Case LINK_MANUAL
            lablTextLinkStatus.Caption = "Manual Link"
       Case LINK_NOTIFY
            lablTextLinkStatus.Caption = "Notify Link"
       Case Else
            lablTextLinkStatus.Caption = "(unlinked)"
     End Select
End Sub

Sub menuBar_Click (Index As Integer)
     Select Case Index
       Case 0 ' ------------ File
       Case 1 ' ------------ Edit
            menuEdit(0).Enabled = False                 'Turn off all menu choices
            menuEdit(1).Enabled = False
            menuEdit(2).Enabled = False
            menuEdit(3).Enabled = False
            If Clipboard.GetFormat(CF_LINK) Then        'if there is link info on clipboard
              linkID = Clipboard.GetText(CF_LINK)       'find out what the link info is
              itemSeparator = InStr(linkID, "!")        'this separates topic from item
              If Left$(linkID, itemSeparator - 1) <> THISAPPNAME & "|" & TOPICNAME Then
                link = True                             'another app's link; OK for linking
              End If
            End If
            If Clipboard.GetFormat(CF_TEXT) Then Text = True    'there's text on the clipboard
            If Clipboard.GetFormat(CF_BITMAP) Then pict = True  'there's a bitmap on clipboard
            If TypeOf Screen.ActiveControl Is PictureBox Then   'are we on a picture box?
              menuEdit(0).Enabled = True                'OK to cut picture
              menuEdit(1).Enabled = True                'OK to copy picture
              If pict Then                              'if there's a picture on the clipboard
                menuEdit(2).Enabled = True              'OK to paste clipboard's picture
                If link Then                            'if the picture is linked,
                     menuEdit(3).Enabled = True         'OK to offer Paste Link for picture
                End If
              End If
            End If
            If TypeOf Screen.ActiveControl Is TextBox Then  'if we're on a text box,
              If Screen.ActiveControl.SelText <> "" Then    'and there is some text selected
                menuEdit(0).Enabled = True              'OK to cut text
                menuEdit(1).Enabled = True              'OK to copy text
              End If
              If Text Then                              'if there's text on the clipboard
                menuEdit(2).Enabled = True              'OK to paste text into text box
                If link And Screen.ActiveControl.Tag = "textBox" Then  'if there's a link and
                                                        'we're in the top text box,
                     menuEdit(3).Enabled = True         'OK to offer Paste Link for the text
                End If
              End If
            End If
       Case 2 ' ----------- Links
     End Select
End Sub

Sub menuEdit_Click (Index As Integer)
     Select Case Index
       Case 0 '*********************************************** Cut
            If TypeOf Screen.ActiveControl Is TextBox Then        'on a text box?
              Clipboard.SetText Screen.ActiveControl.Text         'put text on clipboard
              Screen.ActiveControl.SelText = ""                   'cut out selected text
            ElseIf TypeOf Screen.ActiveControl Is PictureBox Then 'on a picture box?
              Clipboard.SetData Screen.ActiveControl.Picture      'put picture on clipboard
              Screen.ActiveControl.Picture = LoadPicture("")      'blank out picture
            End If
            itemName = Screen.ActiveControl.Tag                   'Tag has the control's Name
            Clipboard.SetText THISAPPNAME & "|" & TOPICNAME & "!" & itemName, CF_LINK
       Case 1 '*********************************************** Copy
            If TypeOf Screen.ActiveControl Is TextBox Then        'on a textbox?
              Clipboard.SetText Screen.ActiveControl.Text         'put text on clipboard
            ElseIf TypeOf Screen.ActiveControl Is PictureBox Then 'on a picturebox?
              Clipboard.Clear                                     'get rid of any existing junk on clipboard
              Clipboard.SetData Screen.ActiveControl.Picture      'put picture on clipboard
            End If
            itemName = Screen.ActiveControl.Tag                   'Tag has the control's Name
            Clipboard.SetText THISAPPNAME & "|" & TOPICNAME & "!" & itemName, CF_LINK
       Case 3 '*********************************************** Paste
            Screen.ActiveControl.LinkMode = NONE                  'no link status
            If TypeOf Screen.ActiveControl Is TextBox Then        'on a textbox?
              Screen.ActiveControl.SelText = Clipboard.GetText()  'put text from clipboard
            ElseIf TypeOf Screen.ActiveControl Is PictureBox Then 'on a picture box?
              Screen.ActiveControl.Picture = Clipboard.GetData()  'put picture from clipboard
            End If
       Case 4 '*********************************************** Paste Link
            linkID = Clipboard.GetText(CF_LINK)        'get link information from clipboard
            itemSeparator = InStr(linkID, "!")         'this separates topicname from itemname
            Screen.ActiveControl.LinkMode = NONE       'turn off any existing links
            Screen.ActiveControl.LinkTopic = Left$(linkID, itemSeparator - 1)
            Screen.ActiveControl.LinkItem = Mid$(linkID, itemSeparator + 1)
            Screen.ActiveControl.LinkMode = MenuLinkModeCheck  'link mode to what’s on link menu
     End Select
     labelLinkStatus                                   'update labels with correct link status
End Sub

Sub menuFile_Click ()
     End                                               'end the program
End Sub

Sub menuLink_Click (Index As Integer)
     Select Case Index
       Case 0, 1, 2   'link mode
            menuLink(0).Checked = False             'turn off the all checkmarks
            menuLink(1).Checked = False
            menuLink(2).Checked = False
            menuLink(Index).Checked = True          'check the one the user clicked
            MenuLinkModeCheck = Index + 1           'and set the correct linkmode
       Case 4         'update links
            If TextBox.LinkMode <> NONE Then        'if there is a link for text box
              TextBox.LinkRequest                   'update it
            End If
            If pictGraph.LinkMode <> NONE Then      'if there is a link for picture box
              pictGraph.LinkRequest                 'update it
            End If
            labelLinkStatus                         'update display of link status labels
     End Select
End Sub

Sub pictGraph_LinkClose ()
     lablGraphLinkStatus.Caption = "(unlinked)"        'no more link
End Sub

Sub pictGraph_LinkError (LinkErr As Integer)
     describeLinkError LinkErr, "picture box."         'show the error message
End Sub

Sub pictGraph_LinkNotify ()
     lablGraphLinkStatus.Caption = "(Data Available)"  'notify user new data available
End Sub

Sub pictGraph_LinkOpen (Cancel As Integer)
     lablGraphLinkStatus.Caption = "...linking..."     'could take a while, eh?
End Sub

Sub textBox_LinkError (LinkErr As Integer)
     describeLinkError LinkErr, "text box."            'display the error message
End Sub

Sub textBox_LinkNotify ()
     lablTextLinkStatus.Caption = "(Data Available)"   'notify user new data available
End Sub

Sub textCogs_Change (Index As Integer)
     textCogs(Index).LinkPoke                          'give excel the new data
     lablTotal(Index).LinkRequest                      'and update the totals
End Sub

Sub textCogs_GotFocus (Index As Integer)
     textCogs(Index).SelLength = 100                   'select all text in the box
End Sub

Sub textSales_Change (Index As Integer)
     textSales(Index).LinkPoke                         'give excel the new data
     lablTotal(Index).LinkRequest                      'and update the totals
End Sub

Sub textSales_GotFocus (Index As Integer)
     textSales(Index).SelLength = 100                  'select all text in the box
End Sub

