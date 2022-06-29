VERSION 2.00
Begin Form frmOptions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailing Options"
   ClientHeight    =   2565
   ClientLeft      =   1620
   ClientTop       =   1890
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   2970
   Left            =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5835
   Top             =   1545
   Width           =   5955
   Begin CommandButton Command1 
      Caption         =   "next option"
      Enabled         =   0   'False
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   255
      TabIndex        =   15
      Top             =   1725
      Width           =   2655
   End
   Begin SSOption optAlpha 
      Caption         =   "Sorted Alpha"
      Enabled         =   0   'False
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1455
      Width           =   1335
   End
   Begin SSOption optSelected 
      Caption         =   "Selected Mails"
      Enabled         =   0   'False
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin CommandButton cmdOK 
      Caption         =   "Looks Good"
      FontBold        =   0   'False
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   4650
      TabIndex        =   12
      Top             =   60
      Width           =   1065
   End
   Begin CommandButton cmdMMFlashsession 
      Caption         =   "MM Flashsessions"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   180
      TabIndex        =   11
      Top             =   2040
      Width           =   2790
   End
   Begin SSCheck chkNotify 
      Caption         =   "Notify Additions"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   165
      TabIndex        =   10
      Top             =   1455
      Width           =   1380
   End
   Begin TextBox txtMessage 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1935
      Left            =   3300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   405
      Width           =   2295
   End
   Begin SSFrame fraMessage 
      Caption         =   "Mailing Message:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   2340
      Left            =   3165
      TabIndex        =   8
      Top             =   90
      Width           =   2550
   End
   Begin SSOption optRange 
      Caption         =   "Select Range"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   930
      Width           =   1335
   End
   Begin SSOption optAll 
      Caption         =   "Send All"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   675
      Width           =   1335
   End
   Begin SSOption optCurrent 
      Caption         =   "Send Current"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   405
      Value           =   -1  'True
      Width           =   1335
   End
   Begin SSCheck chkSignOffAfter 
      Caption         =   "Sign-Off After"
      Enabled         =   0   'False
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   165
      TabIndex        =   4
      Top             =   1185
      Width           =   1335
   End
   Begin SSCheck chkIMFull 
      Caption         =   "IM Full Boxes"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   165
      TabIndex        =   3
      Top             =   930
      Width           =   1335
   End
   Begin SSCheck chkStartatBottom 
      Caption         =   "Start at Bottom"
      Enabled         =   0   'False
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   165
      TabIndex        =   2
      Top             =   660
      Width           =   1335
   End
   Begin SSCheck chkKeepasNew 
      Caption         =   "Keep as New"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   165
      TabIndex        =   1
      Top             =   390
      Value           =   -1  'True
      Width           =   1335
   End
   Begin SSFrame fraMailOptions 
      Caption         =   "Mailing Options:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   2340
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   3000
   End
   Begin Menu mnuMsgOptions 
      Caption         =   "MsgOptions"
      Visible         =   0   'False
      Begin Menu mnuItalics 
         Caption         =   "Italics"
      End
      Begin Menu mnuBold 
         Caption         =   "Bold"
      End
      Begin Menu mnuLink 
         Caption         =   "Link"
      End
      Begin Menu mnuMailRef 
         Caption         =   "Mail Ref."
      End
   End
End

Sub cmdMMFlashsession_Click ()
MsgBox "This feature is finished but buggy, will be available in the next beta release"
End Sub

Sub cmdOK_Click ()
For x = 0 To 5
Me.Hide
Me.Hide
Me.Hide
Open app.Path & "\SoLuTioN.INI" For Output As #1
    Print #1, "Message=" & txtmessage.Text
    Close #1
Next x
frmmmer.Enabled = True
frm% = frmmmer.hWnd
x = setfocusapi(frm%)
End Sub

Sub fadeform (TheForm As Form)
Dim hBrush%
    Dim FormHeight%, red%, StepInterval%, x%, retVal%, OldMode%
    Dim FillArea As RECT
    OldMode = TheForm.ScaleMode
    TheForm.ScaleMode = 3  'Pixel
    FormHeight = TheForm.ScaleHeight
' Divide the form into 63 regions
    StepInterval = FormHeight \ 63
    red = 255
    FillArea.Left = 0
    FillArea.Right = TheForm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = StepInterval
    For x = 1 To 63
	hBrush% = CreateSolidBrush(RGB(0, 0, red))
	retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
	retVal% = DeleteObject(hBrush)
	red = red - 4
	FillArea.Top = FillArea.Bottom
	FillArea.Bottom = FillArea.Bottom + StepInterval
    Next
' Fill the remainder of the form with black
    FillArea.Bottom = FillArea.Bottom + 63
    hBrush% = CreateSolidBrush(RGB(0, 0, 0))
    retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
    retVal% = DeleteObject(hBrush)
    TheForm.ScaleMode = OldMode
End Sub

Sub Form_GotFocus ()
stayontop Me

End Sub

Sub Form_Load ()
If ifileexists(app.Path & "\SoLuTioN.ini") = False Then GoTo skipini
 Close #1
Open app.Path & "\SoLuTioN.INI" For Input As #1
    Do While Not EOF(1)
    Line Input #1, f$
    If InStr(f$, "Message=") Then txtmessage.Text = Mid$(f$, 9)
    Loop
    Close #1

skipini:
stayontop Me
frm% = Me.hWnd: x = setfocusapi(frm%)
End Sub

Sub Form_LostFocus ()
stayontop Me
End Sub

Sub Form_Resize ()
stayontop Me
End Sub

Sub make3d (frm As Form, Ctl As Control)
frm.ScaleMode = 3
frm.CurrentX = Ctl.Left - 1
frm.CurrentY = Ctl.Top + Ctl.Height
frm.Line -Step(0, -(Ctl.Height + 1)), RGB(92, 92, 92)
frm.Line -Step(Ctl.Width + 1, 0), RGB(92, 92, 92)
frm.Line -Step(0, Ctl.Height + 1), RGB(255, 255, 255)
frm.Line -Step(-(Ctl.Width + 1), 0), RGB(255, 255, 255)
End Sub

Sub mnuBold_Click ()
MsgBox "This can be done with the HTML code <B>, Example:" & Chr(13) & Chr(13) & "<B>Glen is your god</B>", 32
End Sub

Sub mnuItalics_Click ()
MsgBox "This can be done with the HTML code <I>, Example:" & Chr(13) & Chr(13) & "<I>Glen is your god</I>", 32
End Sub

Sub mnuLink_Click ()
MsgBox "This can be done with the HTML code <A HREF>, Example:" & Chr(13) & Chr(13) & "<A HREF=""http://glen.is.your.god"">Glen is your god</A>", 32
End Sub

Sub mnuMailRef_Click ()
MsgBox "This can be done with the HTML code <mailto:>, Example:" & Chr(13) & Chr(13) & "<A HREF=mailto:glenisme@hotmail.com>E-Mail Glen</A>", 32
End Sub

Sub optAll_Click (Value As Integer)
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
tree% = findchildbyclass(mdi%, "_AOL_TREE")
mbox% = getparent(tree%): DoEvents
mstr$ = grabtext(mbox%): DoEvents
If InStr(mstr$, "Mail") = False Then MsgBox "Your New Mail or Incoming FlashMail box must be open", 32: optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
tree% = findchildbyclass(mbox%, "_AOL_TREE")
nCount = sendmessage(tree%, LB_GETCOUNT, 0, 0&): DoEvents
If nCount = 0 Then MsgBox "You have no mail!", 32: optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
HighMail% = nCount: LowMail% = 0
End Sub

Sub optAlpha_Click (Value As Integer)
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
tree% = findchildbyclass(mdi%, "_AOL_TREE")
mbox% = getparent(tree%): DoEvents
mstr$ = grabtext(mbox%): DoEvents
If InStr(mstr$, "Mail") = False Then MsgBox "Your New Mail or Incoming FlashMail box must be open", 32: optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
MsgBox "Sorted Alpha not implemented in rude beta copy", 16: Exit Sub
optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
tree% = findchildbyclass(mbox%, "_AOL_TREE")
nCount = sendmessage(tree%, LB_GETCOUNT, 0, 0&): DoEvents
HighMail% = nCount: LowMail% = 0
End Sub

Sub optImage_Click (Value As Integer)
End Sub

Sub optRange_Click (Value As Integer)
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
tree% = findchildbyclass(mdi%, "_AOL_TREE")
mbox% = getparent(tree%): DoEvents
mstr$ = grabtext(mbox%): DoEvents
If InStr(mstr$, "Mail") = False Then MsgBox "Your New Mail or Incoming FlashMail box must be open", 32: optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
tree% = findchildbyclass(mbox%, "_AOL_TREE")
nCount = sendmessage(tree%, LB_GETCOUNT, 0, 0&)
If nCount = 0 Then MsgBox "You have no mail!", 32: optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
MsgBox "Select the first mail to be sent out in your real mailbox", 32
FirstMail% = sendmessage(tree%, LB_GETCURSEL, 0&, 0&): DoEvents
MsgBox "Select the last mail to be sent out in your real mailbox", 32
LastMail% = sendmessage(tree%, LB_GETCURSEL, 0&, 0&): DoEvents
End Sub

Sub optSelected_Click (Value As Integer)
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
tree% = findchildbyclass(mdi%, "_AOL_TREE")
mbox% = getparent(tree%): DoEvents
mstr$ = grabtext(mbox%): DoEvents
If InStr(mstr$, "Mail") = False Then MsgBox "Your New Mail or Incoming FlashMail box must be open", 32: optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
MsgBox "Selected Mails not implemented in rude beta copy", 16: Exit Sub
optcurrent.Value = True: optcurrent.SetFocus : DoEvents: Exit Sub
tree% = findchildbyclass(mbox%, "_AOL_TREE")
nCount = sendmessage(tree%, LB_GETCOUNT, 0, 0&): DoEvents
HighMail% = nCount: LowMail% = 0
End Sub

Sub txtMessage_DblClick ()
txtmessage.SelStart = Len(txtmessage)   ' Set selection start.
txtmessage.SelLength = Len(txtmessage)
 
 PopupMenu mnuMsgOptions
End Sub

Sub txtMessage_KeyPress (KeyAscii As Integer)
Open app.Path & "\SoLuTioN.INI" For Output As #1
    Print #1, "Message=" & txtmessage.Text
    Close #1

End Sub

