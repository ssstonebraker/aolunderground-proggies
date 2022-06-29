Global Const MAPI_ORIG = 0            '// Recipient is message originator
Global Const MAPI_TO = 1              '// Recipient is a primary recipient
Global Const MAPI_CC = 2              '// Recipient is a copy recipient
Global Const MAPI_BCC = 3             '// Recipient is blind copy recipient
Global Const MAIL_LONGDATE = 0
Global Const MAIL_LISTVIEW = 1

Global Const Option_General = 1       '// Constant for Option Dialog Type - General Options
Global Const Option_Messaage = 2      '// Constant for Option Dialog Type - Message Options

Global Const MAPI_ATT_File = 0        'Attachment Type: data File
Global Const MAPI_ATT_EOle = 1        'Attachment Type: embedded OLE Object
Global Const MAPI_ATT_SOle = 2        'Attachment Type: static OLE Object

Type ListDisplay
    Name As String * 20
    Subject As String * 40
    Date As String * 20
End Type
Global currentRCIndex As Integer
Global UnRead As Integer
Global SendWithMapi As Integer
Global ReturnRequest As Integer
Global OptionType As Integer

'----------------------------------------------------------------------------
' ACTION PROPERTY CONSTANTS
'----------------------------------------------------------------------------
Global Const MESSAGE_FETCH = 1             ' Load all messages from message store
Global Const MESSAGE_SENDDLG = 2           ' Send mail bring up default mapi dialog
Global Const MESSAGE_SEND = 3              ' Send mail without default mapi dialog
Global Const MESSAGE_SAVEMSG = 4           ' Save message in the compose buffer
Global Const MESSAGE_COPY = 5              ' Copy current message to compose buffer
Global Const MESSAGE_COMPOSE = 6           ' Initialize compose buffer (previous
                                    ' data is lost
Global Const MESSAGE_Reply = 7             ' Fill Compose buffer as REPLY
Global Const MESSAGE_ReplyAll = 8          ' Fill Compose buffer as REPLY ALL
Global Const Message_Forward = 9           ' Fill Compose buffer as FORWARD
Global Const MESSAGE_DELETE = 10           ' Delete current message
Global Const MESSAGE_SHOWADBOOK = 11       ' Show Address book
Global Const MESSAGE_SHOWDETAILS = 12      ' Show details of the current recipient
Global Const MESSAGE_RESOLVENAME = 13      ' Resolve the display name of the recipient
Global Const RECIPENT_DELETE = 14          ' Delete the current Reciptent
'----------------------------------------------------------------------------
' Windows API Routines
'----------------------------------------------------------------------------
Declare Function GetProfileString% Lib "Kernel" (ByVal lpSection$, ByVal lpEntry$, ByVal lpDefault$, ByVal Buffer$, ByVal cbBuffer%)

Sub Attachments (Msg As Form)
'---- Clear the current attachment list
    Msg.aList.Clear

'---- If there are attachments, load them into the listbox
    If VBMail.MapiMess.AttachmentCount Then
        Msg.NumAtt = VBMail.MapiMess.AttachmentCount + " Files"
        For i% = 0 To VBMail.MapiMess.AttachmentCount - 1
            VBMail.MapiMess.AttachmentIndex = i%
            a$ = VBMail.MapiMess.AttachmentName
            Select Case VBMail.MapiMess.AttachmentType
                Case MAPI_ATT_File
                    a$ = a$ + " (Data File)"
                Case MAPI_ATT_EOle
                    a$ = a$ + " (Embedded OLE Object)"
                Case MAPI_ATT_SOle
                    a$ = a$ + " (Static OLE Object)"
                Case Else
                    a$ = a$ + " (Unknown attachment type)"
            End Select
            Msg.aList.AddItem a$
        Next i%
        
        If Not Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = True
            Call SizeMessageWindow(Msg)
            'If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height + Msg.AttachWin.Height
            'End If
        End If
    
    Else
        If Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = False
            Call SizeMessageWindow(Msg)
            'If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height - Msg.AttachWin.Height
            'End If
        End If
    End If
    Msg.Refresh
End Sub

Sub CopyNamestoMsgBuffer (Msg As Form, fResolveNames As Integer)
    Call KillRecips(VBMail.MapiMess)
    Call SetRCList(Msg.txtTo, VBMail.MapiMess, MAPI_TO, fResolveNames)
    Call SetRCList(Msg.txtcc, VBMail.MapiMess, MAPI_CC, fResolveNames)
End Sub

Function DateFromMapiDate$ (ByVal S$, wFormat%)
'----------------------------------------------------
'   This routine formats a MAPI Date in one of
'   two formats for use in viewing the message
'----------------------------------------------------
    Y$ = Left$(S$, 4)
    M$ = Mid$(S$, 6, 2)
    D$ = Mid$(S$, 9, 2)
    T$ = Mid$(S$, 12)
    Ds# = DateValue(M$ + "/" + D$ + "/" + Y$) + TimeValue(T$)
    Select Case wFormat
        Case MAIL_LONGDATE
            f$ = "dddd, mmmm d, yyyy, h:mmAM/PM"
        Case MAIL_LISTVIEW
            f$ = "mm/dd/yy hh:mm"
    End Select
    DateFromMapiDate = Format$(Ds#, f$)
End Function

Sub DeleteMessage ()
  '------------------------------------------------------------------
  '  If the currently active form is a message, set the MListIndex to
  '  the correct value
  '------------------------------------------------------------------
    If TypeOf VBMail.ActiveForm Is MsgView Then
        MailLst.MList.ListIndex = Val(VBMail.ActiveForm.Tag)
        ViewingMsg = True
    End If

   '------------------------------------------------------------------
   ' Delete the mail message
   '------------------------------------------------------------------
    If MailLst.MList.ListIndex <> -1 Then
        VBMail.MapiMess.MsgIndex = MailLst.MList.ListIndex
        VBMail.MapiMess.Action = 10  'Delete Mail
        x% = MailLst.MList.ListIndex
        MailLst.MList.RemoveItem x%
        If x% < MailLst.MList.ListCount - 1 Then
            MailLst.MList.ListIndex = x%
        Else
            MailLst.MList.ListIndex = MailLst.MList.ListCount - 1
        End If
        VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " Messages"

     '--------------------------------------------------------------------------
     ' Go through and adjust the index values for currently viewed messages
     '--------------------------------------------------------------------------
        If ViewingMsg Then
            VBMail.ActiveForm.Tag = Str$(-1)
        End If

        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) > x% Then
                    Forms(i).Tag = Val(Forms(i).Tag) - 1
                End If
            End If
        Next i
     '--------------------------------------------------------------------------
     ' If we were viewing a message, load the next message into the MsgView form
     ' if the message isn't currently displayed...
     '--------------------------------------------------------------------------
        If ViewingMsg Then
            '--------------------------------------------------------------
            ' First Check to see if the message is currently being viewed
            '--------------------------------------------------------------
            WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
            If WindowNum% > 0 Then
                If Forms(WindowNum%).Caption <> VBMail.ActiveForm.Caption Then
                    Unload VBMail.ActiveForm
                    '--- find the correct window again and show it.  Index isn't valid after the unload
                     Forms(FindMsgWindow((MailLst.MList.ListIndex))).Show
                Else
                     Forms(WindowNum%).Show
                End If
            Else
                Call LoadMessage(MailLst.MList.ListIndex, VBMail.ActiveForm)
            End If
        Else
            '---- Check to see if there was a window viewing the message and unload it
            WindowNum% = FindMsgWindow(x%)
            If WindowNum% > 0 Then
                Unload Forms(x%)
            End If
        End If
     End If
End Sub

Sub DisplayAttachedFile (ByVal FileName As String)
On Error Resume Next
    '----- Determine the extension
        ext$ = FileName
        junk$ = Token$(ext$, ".")
    '----- Get the application from the WIN.INI file to run
        Buffer$ = String$(256, " ")
        errCode% = GetProfileString("Extensions", ext$, "NOTFOUND", Buffer$, 255)
        If errCode% Then
            Buffer$ = Mid$(Buffer$, 1, errCode% - 1)
            If Buffer$ <> "NOTFOUND" Then
                '---- Strip off the ^.EXT information from the string
                ExeName$ = Token$(Buffer$, " ")
                errCode% = Shell(ExeName$ + " " + FileName, 1)
                If Err Then
                    MsgBox "Error occured during the shell: " + Error$
                End If
            Else
                MsgBox "Application that uses: <" + ext$ + "> not found in WIN.INI"
            End If
        End If
End Sub

Function FindMsgWindow (Index As Integer) As Integer
'------------------------------------------------------
'  This function searchs through the active windows
'  and locates those w/ the MsgView type and then
'  checked to see if the tag contains the index we
'  are searching for
'------------------------------------------------------
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) = Index Then
                    FindMsgWindow = i
                    Exit Function
                End If
            End If
        Next i
        FindMsgWindow = -1
End Function

Function GetHeader (Msg As Control) As String
Dim CR As String
CR = Chr$(13) + Chr$(10)
      Header$ = String$(25, "-") + CR
      Header$ = Header$ + "Form: " + Msg.MsgOrigDisplayName + CR
      Header$ = Header$ + "To: " + GetRCList(Msg, MAPI_TO) + CR
      Header$ = Header$ + "Cc: " + GetRCList(Msg, MAPI_CC) + CR
      Header$ = Header$ + "Subject: " + Msg.MsgSubject + CR
      Header$ = Header$ + "Date: " + DateFromMapiDate$(Msg.MsgDateReceived, MAIL_LONGDATE) + CR + CR
      GetHeader = Header$
End Function

Sub GetMessageCount ()
'--------------------------------------------------
'  Reads all mail messages and displays the count
'--------------------------------------------------
    Screen.MousePointer = 11
    VBMail.MapiMess.FetchUnreadOnly = 0
    VBMail.MapiMess.Action = MESSAGE_FETCH
    VBMail.MsgCountLbl = Format$(VBMail.MapiMess.MsgCount) + " Messages"
    Screen.MousePointer = 0
End Sub

Function GetRCList (Msg As Control, RCType As Integer) As String
'--------------------------------------------------
'  Given a list of Recips, this function returns
'  a list of Recips of the specified type in the
'  following format:
'
'       Person 1;Person 2;Person 3
'
'--------------------------------------------------
    For i = 0 To Msg.RecipCount - 1
        Msg.RecipIndex = i
        If RCType = Msg.RecipType Then
                a$ = a$ + ";" + Msg.RecipDisplayName
        End If
    Next i
    If a$ <> "" Then
       a$ = Mid$(a$, 2)  'Strip-off leading ";"
    End If
    GetRCList = a$
End Function

Sub KillRecips (MsgControl As Control)
'---- Delete each reciptent.  Loop until no more exist
    While MsgControl.RecipCount
        MsgControl.Action = RECIPENT_DELETE
    Wend
End Sub

Sub LoadList (mailctl As Control)
'------------------------------------------------------
'   This routines loads the mail message headers
'   into the MailLst.MList.  Unread messages have
'   a chr$(187) placed at the beginning of the string
'------------------------------------------------------
    MailLst.MList.Clear
    UnRead = 0
    StartIndex = 0
    For i = 0 To mailctl.MsgCount - 1
        mailctl.MsgIndex = i
        If Not mailctl.MsgRead Then
            a$ = Chr$(187) + " "
            If UnRead = 0 Then
                StartIndex = i  'Position to start in the mail list
            End If
            UnRead = UnRead + 1
        Else
            a$ = "  "
        End If
        a$ = a$ + Mid$(Format$(mailctl.MsgOrigDisplayName, "!" + String$(10, "@")), 1, 10)
        If mailctl.MsgSubject <> "" Then
            b$ = Mid$(Format$(mailctl.MsgSubject, "!" + String$(35, "@")), 1, 35)
        Else
            b$ = String$(30, " ")
        End If
        C$ = Mid$(Format$(DateFromMapiDate(mailctl.MsgDateReceived, MAIL_LISTVIEW), "!" + String$(15, "@")), 1, 15)
        MailLst.MList.AddItem a$ + Chr$(9) + b$ + Chr$(9) + C$
        MailLst.MList.Refresh
    Next i

    MailLst.MList.ListIndex = StartIndex
    
    '----- Enable the correct buttons
    VBMail.Next.Enabled = True
    VBMail.Previous.Enabled = True
    VBMail![Delete].Enabled = True

    '----- Adjust the value of the labels displaying message counts
    If UnRead Then
        VBMail.UnreadLbl = " - " + Format$(UnRead) + " Unread"
        MailLst.Icon = MailLst.NewMail.Picture
    Else
        VBMail.UnreadLbl = ""
        MailLst.Icon = MailLst.nonew.Picture
    End If
End Sub
    

Sub LoadMessage (ByVal Index As Integer, Msg As Form)
'------------------------------------------------------
'   This routine loads the specified mail message into
'   a form to either view or edit a message
'------------------------------------------------------
    If TypeOf Msg Is MsgView Then
        a$ = MailLst.MList.List(Index)
        '---- Message is unread; reset the text
        If Mid$(a$, 1, 1) = Chr$(187) Then
            Mid$(a$, 1, 1) = " "
            MailLst.MList.List(Index) = a$
            UnRead = UnRead - 1
            If UnRead Then
                VBMail.UnreadLbl = Format$(UnRead) + " Unread"
            Else
                VBMail.UnreadLbl = ""
                '---- Change the icon on the list window
                MailLst.Icon = MailLst.nonew.Picture
            End If
        End If
    End If

    '----- These fields only apply to viewing
    If TypeOf Msg Is MsgView Then
        VBMail.MapiMess.MsgIndex = Index
        Msg.txtDate = DateFromMapiDate$(VBMail.MapiMess.MsgDateReceived, MAIL_LONGDATE)
        Msg.txtFrom = VBMail.MapiMess.MsgOrigDisplayName
        MailLst.MList.ItemData(Index) = True
    End If
    '----- These fields apply to both form types
    Call Attachments(Msg)
    Msg.txtNoteText = VBMail.MapiMess.MsgNoteText
    Msg.txtsubject = VBMail.MapiMess.MsgSubject
    Msg.Caption = VBMail.MapiMess.MsgSubject
    Msg.Tag = Index
    Call UpdateRecips(Msg)
    Msg.Refresh
    Msg.Show
End Sub

Sub LogOffUser ()
    On Error Resume Next
    VBMail.MapiSess.Action = 2
    If Err <> 0 Then
        MsgBox "Logoff Failure: " + ErrorR
    Else
        VBMail.MapiMess.SessionID = 0
        '----- Adjust the menu items
        VBMail.LogOff.Enabled = 0
        VBMail.Logon.Enabled = -1
        '---- Unload all the windows; but the MDIForm
        Do Until Forms.Count = 1
            i = Forms.Count - 1
            If TypeOf Forms(i) Is MDIForm Then
                'do nothing
            Else
                Unload Forms(i)
            End If
        Loop
        '---- Disable the toolbar buttons
        VBMail.Next.Enabled = False
        VBMail.Previous.Enabled = False
        VBMail![Delete].Enabled = False
        VBMail.SendCtl(MESSAGE_COMPOSE).Enabled = False
        VBMail.SendCtl(MESSAGE_ReplyAll).Enabled = False
        VBMail.SendCtl(MESSAGE_Reply).Enabled = False
        VBMail.SendCtl(Message_Forward).Enabled = False
        VBMail.rMsgList.Enabled = False
        VBMail.PrintMessage.Enabled = False
        VBMail.DispTools.Enabled = False
        VBMail.EditDelete.Enabled = False
                          
        '---- Reset the caption for the status bar labels
        VBMail.MsgCountLbl = "Off Line"
        VBMail.UnreadLbl = ""
    End If

End Sub

Sub PrintLongText (ByVal LongText As String)
'------------------------------------------------------
'   This routine prints a text stream to a printer and
'   ensures that words are not split between lines and
'   wrap as needed
'------------------------------------------------------
    Do Until LongText = ""
        Word$ = Token$(LongText, " ")
        If Printer.TextWidth(Word$) + Printer.CurrentX > Printer.Width - Printer.TextWidth("ZZZZZZZZ") Then
            Printer.Print
        End If
        Printer.Print " " + Word$;
    Loop
End Sub

Sub PrintMail ()
    '---- In List view all selected messages are printed
    '---- In Message view, the selected message is printed

    If TypeOf VBMail.ActiveForm Is MsgView Then
        Call PrintMessage(VBMail.MapiMess, False)
        Printer.EndDoc
    ElseIf TypeOf VBMail.ActiveForm Is MailLst Then
        For i = 0 To MailLst.MList.ListCount - 1
            If MailLst.MList.Selected(i) Then
                VBMail.MapiMess.MsgIndex = i
                Call PrintMessage(VBMail.MapiMess, False)
            End If
        Next i
        Printer.EndDoc
    End If
End Sub

Sub PrintMessage (Msg As Control, fNewPage As Integer)
    '-------------------------------------------
    '   Print a mail message
    '-------------------------------------------
    Screen.MousePointer = 11
    '----- Start a new page if needed
    If fNewPage Then
        Printer.NewPage
    End If
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.DrawWidth = 10
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.Print
    Printer.FontSize = 9.75
    Printer.Print "From:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgOrigDisplayName
    Printer.Print "To:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, MAPI_TO)
    Printer.Print "Cc:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, MAPI_CC)
    Printer.Print "Subject:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgSubject
    Printer.Print "Date:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print DateFromMapiDate$(Msg.MsgDateReceived, MAIL_LONGDATE)
    Printer.Print
    Printer.DrawWidth = 5
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.FontSize = 9.75
    Printer.FontBold = False
    Call PrintLongText(Msg.MsgNoteText)
    Printer.Print
    Screen.MousePointer = 0
End Sub

Sub SaveMessage (Msg As Form)
    '--- Save the current subject and note text
    '    copy the message to the compose buffer
    '    reset the subject and message text
    '    save the message
    svSub = Msg.txtsubject
    svNote = Msg.txtNoteText
    VBMail.MapiMess.Action = MESSAGE_COPY
    VBMail.MapiMess.MsgSubject = svSub
    VBMail.MapiMess.MsgNoteText = svNote
    VBMail.MapiMess.Action = MESSAGE_SAVE
End Sub

Sub SetRCList (ByVal NameList As String, Msg As Control, RCType As Integer, fResolveNames As Integer)
'--------------------------------------------------
'  Given a list of Recips in the form:
'
'       Person 1;Person 2;Person 3
'
'   This sub places then names into the Msg.Recip
'   structures.
'
'--------------------------------------------------
    If NameList = "" Then
        Exit Sub
    End If

    i = Msg.RecipCount
    Do
        Msg.RecipIndex = i
        Msg.RecipDisplayName = Trim$(Token(NameList, ";"))
        If fResolveNames Then
            Msg.Action = MESSAGE_RESOLVENAME
        End If
        Msg.RecipType = RCType
        i = i + 1
    Loop Until (NameList = "")
End Sub

Sub SizeMessageWindow (MsgWindow As Form)
    If MsgWindow.WindowState <> 1 Then
        '--- Determine smalled allowed window size based
        '    on the visiblity of AttachWin (Attachment window)
        If MsgWindow.AttachWin.Visible Then    'Attachment Window
            MinSize = 3700
        Else
            MinSize = 3700 - MsgWindow.AttachWin.Height
        End If

        '---- Maintain minimum form size
        If MsgWindow.Height < MinSize And (MsgWindow.WindowState = 0) Then
            MsgWindow.Height = MinSize
            Exit Sub

        End If
        '---- Adjust the size of the textbox
        If MsgWindow.ScaleHeight > MsgWindow.txtNoteText.Top Then
            If MsgWindow.AttachWin.Visible Then
                x% = MsgWindow.AttachWin.Height
            Else
                x% = 0
            End If
            MsgWindow.txtNoteText.Height = MsgWindow.ScaleHeight - MsgWindow.txtNoteText.Top - x%
            MsgWindow.txtNoteText.Width = MsgWindow.ScaleWidth
        End If
    End If

End Sub

Function Token$ (tmp$, search$)
    x = InStr(1, tmp$, search$)
    If x Then
       Token$ = Mid$(tmp$, 1, x - 1)
       tmp$ = Mid$(tmp$, x + 1)
    Else
       Token$ = tmp$
       tmp$ = ""
    End If
End Function

Sub UpdateRecips (Msg As Form)
'---- This routine updates the correct editfields the
'---- the Recip information.
    Msg.txtTo.Text = GetRCList(VBMail.MapiMess, MAPI_TO)
    Msg.txtcc.Text = GetRCList(VBMail.MapiMess, MAPI_CC)
End Sub

Sub ViewNextMsg ()
    '--------------------------------------------------
    ' Check to see if the message is currently loaded.
    '    If YES -> Show that form
    '    If NO  -> Load the message
    '--------------------------------------------------
    WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
    If WindowNum% > 0 Then
        Forms(WindowNum%).Show
    Else
        If TypeOf VBMail.ActiveForm Is MsgView Then
            Call LoadMessage(MailLst.MList.ListIndex, VBMail.ActiveForm)
        Else
            Dim Msg As New MsgView
            Call LoadMessage(MailLst.MList.ListIndex, Msg)
        End If
    End If
End Sub

