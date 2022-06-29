VERSION 2.00
Begin Form frmMailTools 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailbox Tools"
   ClientHeight    =   2205
   ClientLeft      =   960
   ClientTop       =   2085
   ClientWidth     =   6255
   Height          =   2610
   Left            =   900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6255
   Top             =   1740
   Width           =   6375
   Begin SSCommand cmdSave 
      BevelWidth      =   1
      Caption         =   "Save Mail List"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      HelpContextID   =   50
      Left            =   3945
      TabIndex        =   8
      Top             =   1665
      Width           =   1215
   End
   Begin ListBox lstNames 
      Height          =   420
      Left            =   1335
      TabIndex        =   7
      Top             =   165
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ListBox List1 
      Height          =   420
      Left            =   105
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SSCommand cmdViewNames 
      BevelWidth      =   1
      Caption         =   "View Names"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   2745
      TabIndex        =   5
      Top             =   1665
      Width           =   1215
   End
   Begin SSCommand cmdUpdateMail 
      BevelWidth      =   1
      Caption         =   "Update Mail List"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   1545
      TabIndex        =   4
      Top             =   1665
      Width           =   1215
   End
   Begin SSCommand cmdDLLater 
      BevelWidth      =   1
      Caption         =   "Quick DL Later"
      Enabled         =   0   'False
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      Left            =   345
      TabIndex        =   3
      Top             =   1665
      Width           =   1215
   End
   Begin SSFrame Frame3D2 
      Caption         =   "Mail Commands:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   210
      TabIndex        =   2
      Top             =   1395
      Width           =   5760
   End
   Begin ListBox lstMail 
      BackColor       =   &H00FFFFFF&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   210
      TabIndex        =   0
      Top             =   285
      Width           =   5760
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Mail List:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   2085
      Left            =   60
      TabIndex        =   1
      Top             =   15
      Width           =   6075
   End
   Begin Menu mnuMnu 
      Caption         =   "Mnu"
      Visible         =   0   'False
      Begin Menu mnuCountMail 
         Caption         =   "Count Mail"
      End
   End
End

Sub cmdSave_Click ()
If lstMail.ListCount = 0 Then MsgBox "load yer mail first niggah!": Exit Sub
Open app.Path & "\MMLIST.TXT" For Output As #1
For X = 0 To lstMail.ListCount - 1
Print #1, lstMail.List(X): DoEvents
Next X
Close #1
MsgBox "The file is called " & app.Path & "\MMLIST.TXT .. thats all folks!"

End Sub

Sub cmdUpdateMail_Click ()
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
List% = findchildbyclass(mdi%, "_AOL_TREE")
If List% = 0 Then MsgBox "Unable to locate any mailbox", 32: Exit Sub
n% = sendmessage(List%, LB_GETCOUNT, 0, 0&)
If n% = 0 Then MsgBox "You have no mail in your open box", 32: Exit Sub
        
lstMail.Clear
For index% = 0 To n% - 1
    DoEvents
    StringSpace$ = String(255, 0)
    g% = SendMessageByString(List%, LB_GETTEXT, index%, StringSpace$)
    StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
    StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
    StringSpace$ = Left(StringSpace$, InStr(1, StringSpace$, Chr$(0)) - 1)
    'List1.AddItem "(" & Trim(Str$(index%)) & ") " & StringSpace$
    'p$ = "(" & Trim(Str$(index%)) & ") " & StringSpace$ & Chr$(13) & Chr$(10)
    lstMail.AddItem StringSpace$
    'ListWait.Label1.Caption = "Building Mail List... (Mail " & Trim(Str$(index%)) & " of " & Trim(Str$(n%)) & ")"
Next index%
End Sub

Sub cmdViewNames_Click ()
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
List% = findchildbyclass(mdi%, "_AOL_TREE")
If List% = 0 Then MsgBox "Unable to locate any mailbox", 32: Exit Sub
n% = sendmessage(List%, LB_GETCOUNT, 0, 0&)
If n% = 0 Then MsgBox "You have no mail in your open box", 32: Exit Sub
If lstMail.ListCount = 0 Then MsgBox "Please update your current mail list before continuing", 32: Exit Sub
mousepointer = 11

For index% = 0 To n% - 1
    DoEvents
    StringSpace$ = String(255, 0)
    g% = SendMessageByString(List%, LB_GETTEXT, index%, StringSpace$)
    List1.AddItem StringSpace$
 Next index%

Rem    StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
Rem    StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
Rem    StringSpace$ = Left(StringSpace$, InStr(1, StringSpace$, Chr$(0)) - 1)

For X = 0 To List1.ListCount - 1
    firchar = Left(lstMail.List(X), 1)
    lofir = Left$(List1.List(X), InStr(List1.List(X), firchar) - 1)
    afterdate = Mid$(lofir, InStr(lofir, "/") + 3)
    StringSpace$ = Mid$(lofir, InStr(1, lofir, Chr$(9)) + 1)
    StringSpace$ = Trim$(StringSpace$)
    Call addlist(lstNames, StringSpace$)
  Next X
nme$ = ""
For X = 0 To lstNames.ListCount - 1
If X <> lstNames.ListCount - 1 Then nme$ = nme$ & lstNames.List(X) & Chr(13)
If X = lstNames.ListCount - 1 Then nme$ = nme$ & lstNames.List(X)
Next X
mousepointer = 1
    MsgBox "You have mail from: " & Chr(13) & Chr(13) & nme$ & Chr(13) & "In your MailBox", 32

End Sub

Sub Command1_Click ()
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
'mb% = findchildbytitle(mdi%, "Incoming FlashMail")
'If mb% = 0 Then MsgBox "wheres yer mailbox?": Exit Sub
List% = findchildbyclass(mdi%, "_AOL_TREE")
If List% = 0 Then MsgBox "where's yer mail?": Exit Sub
n% = sendmessage(List%, LB_GETCOUNT, 0, 0&)
If n% = 0 Then MsgBox "no mail!": Exit Sub
        
        Rem lentos = sendmessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
        Rem Buffer$ = Space$(lentos)
        Rem X = sendmessagebystring(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
        Rem GetWinText = Buffer$

For index% = 0 To n% - 1
    DoEvents
    StringSpace$ = String(255, 0)
    g% = SendMessageByString(List%, LB_GETTEXT, index%, StringSpace$)
    StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
    StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
    StringSpace$ = Left(StringSpace$, InStr(1, StringSpace$, Chr$(0)) - 1)
    'List1.AddItem "(" & Trim(Str$(index%)) & ") " & StringSpace$
    p$ = "(" & Trim(Str$(index%)) & ") " & StringSpace$ & Chr$(13) & Chr$(10)
    lstMail.AddItem StringSpace$
    'ListWait.Label1.Caption = "Building Mail List... (Mail " & Trim(Str$(index%)) & " of " & Trim(Str$(n%)) & ")"
Next index%
End Sub

Sub Command2_Click ()
If lstMail.ListCount = 0 Then MsgBox "load yer mail first niggah!": Exit Sub
Open app.Path & "\MM.LST" For Output As #1
For X = 0 To lstMail.ListCount - 1
Print #1, lstMail.List(X): DoEvents
Next X
Close #1
MsgBox "The file is called " & app.Path & "\MM.LST .. thats all folks!"

End Sub

Sub Form_GotFocus ()
stayontop Me
End Sub

Sub Form_Load ()
lstMail.Height = 1080
stayontop Me
frm% = Me.hWnd: X = setfocusapi(frm%): DoEvents
End Sub

Sub Form_LostFocus ()
stayontop Me
End Sub

Sub Form_Resize ()
stayontop Me
End Sub

Sub lstMail_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuMnu

End Sub

Sub mnuCountMail_Click ()
If lstMail.ListCount = 0 Then MsgBox "Please update your mail", 32: Exit Sub
MsgBox "You have " & Trim$(Str(lstMail.ListCount)) & " mail[s] in your box", 32
End Sub

