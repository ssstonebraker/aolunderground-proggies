VERSION 5.00
Begin VB.UserControl CringeOCX 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   ScaleHeight     =   735
   ScaleWidth      =   1065
End
Attribute VB_Name = "CringeOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub PostIcon(Icon As Long)
'Clicks any icon
On Error Resume Next
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(Icon&, WM_LBUTTONUP, 0, 0&)
End Sub
Private Function FindChatRoom() As Long
Dim AOL As Long, MDI As Long, Child As Long
Dim Rich As Long, List As Long, Combo As Long
Dim Icon As Long, ChatStatic As Long
AOL& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
List& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
Icon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
ChatStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
Combo& = FindWindowEx(Child&, 0&, "_AOL_ComboBox", vbNullString)
If AOL& <> 0 And MDI& <> 0 And Rich& <> 0& And List& <> 0& And Icon& <> 0& And ChatStatic& <> 0& And Combo& <> 0& Then
    FindChatRoom& = Child&
    Exit Function
Else
    Do
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
        Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
        Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
        List& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
        Icon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
        ChatStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
        Combo& = FindWindowEx(Child&, 0&, "_AOL_ComboBox", vbNullString)
If AOL& <> 0 And MDI& <> 0 And Rich& <> 0& And List& <> 0& And Icon& <> 0& And ChatStatic& <> 0& And Combo& <> 0& Then
    FindChatRoom& = Child&
    Exit Function
        End If
    Loop Until Child& = 0&
End If
FindChatRoom& = Child&
End Function
Private Sub Window_Close(Window As Long)
Call PostMessage(Window&, WM_CLOSE, 0, 0)
End Sub
Public Sub XMailSend(Personz As String, Subjectz As String, Messagez As String)
Dim AOL As Long, MDI As Long, Email As Long, EmailEdit As Long
Dim EmailRich As Long, EMailIcon As Long, X
Dim TheError As Long, Tool As Long, MailIcon As Long
Dim GetEmailIcon As Long, SubjectBox As Long
AOL& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL ToolBar", vbNullString)
Tool& = FindWindowEx(Tool&, 0&, "_AOL_ToolBar", vbNullString)
MailIcon& = FindWindowEx(Tool&, 0&, "_AOL_Icon", vbNullString)
MailIcon& = FindWindowEx(Tool&, MailIcon&, "_AOL_Icon", vbNullString)
PostIcon (MailIcon&)

Do: DoEvents
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Email& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
EmailEdit& = FindWindowEx(Email&, 0&, "_AOL_Edit", vbNullString)
EmailRich& = FindWindowEx(Email&, 0&, "RICHCNTL", vbNullString)
EMailIcon& = FindWindowEx(Email&, 0&, "_AOL_Icon", vbNullString)
If Email& <> 0 And EmailEdit& <> 0 And EmailRich& <> 0 Then Exit Do
Loop

SubjectBox& = FindWindowEx(Email&, EmailEdit&, "_AOL_EDIT", vbNullString)
SubjectBox& = FindWindowEx(Email&, SubjectBox&, "_AOL_EDIT", vbNullString)

Call SendMessageByString(EmailEdit&, WM_SETTEXT, 0, Person$)
Call SendMessageByString(SubjectBox&, WM_SETTEXT, 0, Subject$)
Call SendMessageByString(EmailRich&, WM_SETTEXT, 0, Message$)

For GetEmailIcon& = 1 To 13
EMailIcon& = FindWindowEx(Email&, EMailIcon&, "_AOL_Icon", vbNullString)
Next GetEmailIcon&

PostIcon (EMailIcon&)

Do: DoEvents
Email& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
If Email& = 0& Then Exit Do
TheError& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
If TheError& <> 0 Then
Window_Close (TheError&)
Window_Close (Email&)
Exit Do
End If
Loop Until Email& = 0
Exit Sub
End Sub
Public Sub XSendChat(text As String)
Dim Room As Long, Rich As Long, OldText As String, Icon As Long
Room& = FindChatRoom&
If Room& = 0 Then Exit Sub
Rich& = FindWindowEx(Room&, 0&, "RichCntl", vbNullString)
Rich& = FindWindowEx(Room&, Rich&, "RichCntl", vbNullString)
OldText$ = GetText(Rich&)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "")
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, text$)
Call SendMessageLong(Rich&, WM_CHAR, 13, 0&)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, OldText$)
End Sub
Public Sub XKeyword(Keyword As String)
Dim AOL As Long, Tool1 As Long, Tool2 As Long, Tool3 As Long
Dim KWText As Long
AOL& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
Tool1& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
Tool2& = FindWindowEx(Tool1&, 0, "_AOL_Toolbar", vbNullString)
Tool3& = FindWindowEx(Tool2&, 0, "_AOL_ComboBox", vbNullString)
KWText& = FindWindowEx(Tool3&, 0, "Edit", vbNullString)
Call SendMessageByString(KWText&, WM_SETTEXT, 0&, Keyword$)
Call SendMessageLong(KWText&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(KWText&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub XSendIM(Person As String, Message As String)
Dim AOL As Long, MDI As Long, IM As Long, IMRich As Long
Dim IMEdit As Long, SendIcon As Long, TheError As Long
Dim GetSendIcon As Long

AOL& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call XKeyword("aol://9293:" & Person$)

Do: DoEvents
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
IMRich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
IMEdit& = FindWindowEx(IM&, 0&, "_AOL_Edit", vbNullString)
Loop Until IM& <> 0 And IMRich& <> 0 And IMEdit& <> 0

Call SendMessageByString(IMEdit&, WM_SETTEXT, 0, Person$)
Call SendMessageByString(IMRich&, WM_SETTEXT, 0, Message$)

SendIcon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)

For GetSendIcon& = 1 To 8
SendIcon& = FindWindowEx(IM&, SendIcon&, "_AOL_Icon", vbNullString)
Next GetSendIcon&

ClickIcon (SendIcon&)

Do: DoEvents

If FindWindowEx(0, 0&, "#32770", "America Online") <> 0 Then
Window_Close (FindWindow("#32770", "America Online"))
Window_Close (IM&)
Exit Do
End If

If IM& = 0 Then Exit Do

Loop
End Sub

