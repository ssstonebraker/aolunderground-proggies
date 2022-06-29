VERSION 2.00
Begin Form frmBuster 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Room Buster"
   ClientHeight    =   1140
   ClientLeft      =   3015
   ClientTop       =   3660
   ClientWidth     =   4170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   1545
   HelpContextID   =   30
   Left            =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4170
   Top             =   3315
   Width           =   4290
   Begin ComboBox cmbRoom 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   330
      HelpContextID   =   30
      Left            =   165
      TabIndex        =   6
      Top             =   330
      Width           =   2175
   End
   Begin VBMsg VBMsg1 
      Height          =   420
      Left            =   3705
      MessageCount    =   FRMBUSTE.FRX:0000
      MessageList     =   FRMBUSTE.FRX:0002
      MessageTypes    =   0  'Selected Messages
      PostDefault     =   0   'False
      Top             =   690
      Width           =   420
   End
   Begin SSCommand cmdRemove 
      BevelWidth      =   1
      Caption         =   "Remove"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      HelpContextID   =   30
      Left            =   165
      TabIndex        =   5
      Top             =   690
      Width           =   975
   End
   Begin SSCommand cmdBust 
      BevelWidth      =   1
      Caption         =   "Bust-In"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      HelpContextID   =   30
      Left            =   1125
      TabIndex        =   4
      Top             =   690
      Width           =   1215
   End
   Begin SSCommand cmdExit 
      BevelWidth      =   1
      Caption         =   "Exit"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      HelpContextID   =   30
      Left            =   2655
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin SSCommand cmdSaveAll 
      BevelWidth      =   1
      Caption         =   "Save All"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   270
      HelpContextID   =   30
      Left            =   2655
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin SSFrame fraCommands 
      Caption         =   "Commands:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   990
      Left            =   2490
      TabIndex        =   0
      Top             =   45
      Width           =   1530
   End
   Begin SSFrame fraBuster 
      Caption         =   "Buster:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   990
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   2340
   End
End

Sub cmbRoom_KeyPress (keyascii As Integer)
If keyascii = 13 Then
    Call addcombo(cmbRoom, (cmbRoom.Text))
End If
End Sub

Sub cmdBust_Click ()
If cmdBust.Caption = "Bust-In" Then
aol% = FindWindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
toolbar% = findchildbyclass(aol%, "AOL Toolbar")
ico% = findchildbyclass(toolbar%, "_AOL_ICON")
    ico% = getnextwindow(ico%, 2): DoEvents
    ico% = getnextwindow(ico%, 2): DoEvents
    ico% = getnextwindow(ico%, 2): DoEvents
  If iswindowenabled(ico%) = False Then
    MsgBox "You must be signed on to bust into a room", 32: Exit Sub
  End If
  cmdBust.Caption = "End Early"
    Call runmenubystring(aol%, "Keyword..."): DoEvents
    Do: DoEvents
    KeyWnd% = findchildbytitle(mdi%, "Keyword")
    Loop Until KeyWnd% <> 0
 VBMsg1.SubClasshWnd = KeyWnd%
EdithWnd% = findchildbyclass(KeyWnd%, "_AOL_EDIT")
GoIco% = findchildbyclass(KeyWnd%, "_AOL_ICON")
    Call sendtext(EdithWnd%, "aol://2719:2-2-" & (cmbRoom.Text)): DoEvents
nxtBust:
    Call runmenubystring(aol%, "Keyword..."): DoEvents
    Do: DoEvents
    KeyWnd% = findchildbytitle(mdi%, "Keyword")
    Loop Until KeyWnd% <> 0
 VBMsg1.SubClasshWnd = KeyWnd%
EdithWnd% = findchildbyclass(KeyWnd%, "_AOL_EDIT")
GoIco% = findchildbyclass(KeyWnd%, "_AOL_ICON")
    Call sendtext(EdithWnd%, "aol://2719:2-2-" & (cmbRoom.Text)): DoEvents
    
    If cmdBust.Caption = "Bust-In" Then
	VBMsg1.SubClasshWnd = 0
	Call CloseWin(KeyWnd%): DoEvents
	Unload Me
	Exit Sub
    End If
 Call Click(GoIco%, 1): DoEvents
    Do: DoEvents
	Room% = FindChathWnd()
	Full% = FindWindow("#32770", 0&)
    Loop Until Room% <> 0 Or Full% <> 0
  If Room% <> 0 Then
    VBMsg1.SubClasshWnd = 0
    CloseWin KeyWnd%: DoEvents
    Call ChatSay("SoLuTioN HoST", "Room Bust was successful"): DoEvents
    Unload Me
    Exit Sub
  End If
  If Full% <> 0 Then
    sFull$ = MsgBoxText()
	If InStr(sFull$, "not available") = True Then
	    VBMsg1.SubClasshWnd = 0
	    CloseWin KeyWnd%: DoEvents
	    MsgBox (sFull$): DoEvents
	    Unload Me
	    Exit Sub
	Else
	    waitforok
	    GoTo nxtBust
	    Exit Sub
	End If
  End If
Else
cmdBust.Caption = "Bust-In"
End If
End Sub

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdRemove_Click ()
f$ = cmbRoom.Text
For x = 0 To cmbRoom.ListCount - 1
    If UCase$(cmbRoom.List(x)) = UCase(f$) Then
    cmbRoom.Text = ""
    cmbRoom.RemoveItem x
    cmbRoom.SetFocus
    End If
Next x
cmbRoom.SetFocus
End Sub

Sub cmdSaveAll_Click ()
Open app.Path & "\Buster.Rms" For Output As #1
    For x = 0 To cmbRoom.ListCount - 1
	Print #1, cmbRoom.List(x)
    Next x
Close #1
Open app.Path & "\PopUps.INI" For Output As #1
If cmbRoom.ListCount = 0 Then Print #1, "[PopUps]": Print #1, "[End of PopUps]": Close #1: Exit Sub
If cmbRoom.ListCount = 1 Then Print #1, txt$ & Chr(13) & Chr(10) & "[PopUps]" & Chr(13) & Chr(10) & "PopUp0=" & cmbRoom.List(0): Print #1, "[End of PopUps]": Close #1: Exit Sub
If cmbRoom.ListCount = 2 Then Print #1, txt$ & Chr(13) & Chr(10) & "[PopUps]" & Chr(13) & Chr(10) & "PopUp0=" & cmbRoom.List(0) & Chr(13) & Chr(10) & "PopUp1=" & cmbRoom.List(1): Print #1, "[End of PopUps]": Close #1:  Exit Sub
If cmbRoom.ListCount > 2 Then Print #1, txt$ & Chr(13) & Chr(10) & "[PopUps]" & Chr(13) & Chr(10) & "PopUp0=" & cmbRoom.List(0) & Chr(13) & Chr(10) & "PopUp1=" & cmbRoom.List(1) & Chr(13) & Chr(10) & "PopUp2=" & cmbRoom.List(2): Print #1, "[End of PopUps]": Close #1:  Exit Sub
Close #1
End Sub

Sub Command1_Click ()
Room% = FindChathWnd%()
MsgBox Str(Room%)
End Sub

Sub Form_GotFocus ()
stayontop Me
End Sub

Sub Form_Load ()
Rem Unload frmavishow

If ifileexists(app.Path & "\Buster.Rms") = False Then GoTo b

frm% = Me.hWnd: x = setfocusapi(frm%): DoEvents
    Open app.Path & "\Buster.Rms" For Input As #1
      Do While Not EOF(1)
	Line Input #1, rm$
	Call addcombo(cmbRoom, rm$)
      Loop
     Close #1

If PopBuster = True Then
cmdBust_Click
End If

b:
stayontop Me
End Sub

Sub Form_LostFocus ()
stayontop Me
End Sub

Sub Form_Resize ()
stayontop Me
End Sub

Sub VBMsg1_WindowDestroyed (hWindow As Integer)
calldefproc = 0
End Sub

Sub VBMsg1_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, calldefproc As Integer)
calldefproc = 0
End Sub

