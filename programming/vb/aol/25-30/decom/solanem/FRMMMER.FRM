VERSION 2.00
Begin Form frmMMer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Mailer"
   ClientHeight    =   2685
   ClientLeft      =   1695
   ClientTop       =   2355
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   3375
   Left            =   1635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4350
   Top             =   1725
   Width           =   4470
   Begin CommonDialog CMDialog1 
      Left            =   3810
      Top             =   2625
   End
   Begin SSCommand cmdNextBtn 
      BevelWidth      =   1
      Caption         =   "Next Btn"
      Enabled         =   0   'False
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   3375
      TabIndex        =   23
      Top             =   2175
      Width           =   830
   End
   Begin SSCommand cmdUnload 
      BevelWidth      =   1
      Caption         =   "End All/X"
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   3375
      TabIndex        =   22
      Top             =   1905
      Width           =   830
   End
   Begin SSCommand cmdOptions 
      BevelWidth      =   1
      Caption         =   "Options"
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   2565
      TabIndex        =   21
      Top             =   2175
      Width           =   830
   End
   Begin SSCommand cmdStartMM 
      BevelWidth      =   1
      Caption         =   "Start MM"
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   2565
      TabIndex        =   20
      Top             =   1905
      Width           =   830
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
      ForeColor       =   &H00000000&
      Height          =   1020
      Left            =   2490
      TabIndex        =   19
      Top             =   1590
      Width           =   1770
   End
   Begin VBMsg VBMsg1 
      Height          =   420
      Left            =   3525
      MessageCount    =   FRMMMER.FRX:0000
      MessageList     =   FRMMMER.FRX:0002
      MessageTypes    =   0  'Selected Messages
      PostDefault     =   0   'False
      Top             =   2715
      Width           =   420
   End
   Begin Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3060
      Top             =   2715
   End
   Begin ListBox lstIM 
      Height          =   420
      Left            =   1815
      TabIndex        =   18
      Top             =   2700
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2685
   End
   Begin SSCommand cmdScroll 
      BevelWidth      =   1
      Caption         =   "Scroll"
      Enabled         =   0   'False
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   1635
      TabIndex        =   16
      Top             =   2190
      Width           =   720
   End
   Begin SSCommand cmdBotOff 
      BevelWidth      =   1
      Caption         =   "Bot Off"
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   930
      TabIndex        =   15
      Top             =   2190
      Width           =   720
   End
   Begin SSCommand cmdBotOn 
      BevelWidth      =   1
      Caption         =   "Bot On"
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   225
      TabIndex        =   14
      Top             =   2190
      Width           =   720
   End
   Begin VScrollBar VScroll1 
      Height          =   240
      Left            =   450
      Max             =   1
      Min             =   99
      TabIndex        =   12
      Top             =   1860
      Value           =   60
      Width           =   180
   End
   Begin SSFrame fraBot 
      Caption         =   "Mail Manager"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   1020
      Left            =   90
      TabIndex        =   10
      Top             =   1590
      Width           =   2355
      Begin Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes til shutdown"
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   570
         TabIndex        =   13
         Top             =   285
         Width           =   1710
      End
      Begin Label lblMinutes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "60"
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   285
         Width           =   180
      End
   End
   Begin SSCommand cmdRMVName 
      BevelWidth      =   1
      Caption         =   "Del Name"
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   1110
      Width           =   945
   End
   Begin SSCommand cmdAddName 
      BevelWidth      =   1
      Caption         =   "Add Name"
      Font3D          =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   840
      Width           =   945
   End
   Begin TextBox txtSN 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   3240
      MaxLength       =   10
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Top             =   420
      Width           =   945
   End
   Begin ListBox lstPeople 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   450
      Left            =   1995
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   405
      Width           =   1215
   End
   Begin SSFrame fraPPL 
      Caption         =   "People On: 0"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   2340
   End
   Begin TextBox txtRMV 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   225
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Text            =   "Sheep Sex Rules!"
      Top             =   1155
      Width           =   1485
   End
   Begin TextBox txtAdd 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   225
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Text            =   "SoLuTioN MM"
      Top             =   615
      Width           =   1485
   End
   Begin SSFrame fraTrigger 
      Caption         =   "Trigger Phrase(s):"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   1770
      Begin Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Removal:"
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   810
         Width           =   780
      End
      Begin Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addition:"
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   285
         Width           =   780
      End
   End
   Begin Label lblSeconds 
      Caption         =   "60"
      Height          =   495
      Left            =   585
      TabIndex        =   17
      Top             =   2700
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Menu mnuOptions 
      Caption         =   "&Options"
      Begin Menu mnuPeopleList 
         Caption         =   "&People List"
         Begin Menu mnuLoadList 
            Caption         =   "&Load List"
         End
         Begin Menu mnuSaveList 
            Caption         =   "&Save List"
         End
         Begin Menu mnuOptLstSep 
            Caption         =   "-"
         End
         Begin Menu mnuClearList 
            Caption         =   "&Clear List"
         End
         Begin Menu mnuAddRoom 
            Caption         =   "&Add Room"
         End
      End
   End
End

Sub cmdaddname_click ()
Call addlist(lstPeople, txtsn)
frappl.Caption = "People On: " & lstPeople.ListCount
txtsn.Text = ""
txtsn.SetFocus
End Sub

Sub cmdBotOff_Click ()
If cmdscroll.Enabled = False Then MsgBox "The Mail Manager is not active!", 48: Exit Sub
Room% = findChathWnd%()
If Room% = 0 Then MsgBox "Your chatroom has not been detected, ..exiting", 32: Exit Sub
vscroll1.Enabled = True: txtadd.Enabled = True: txtrmv.Enabled = True: timer1.Enabled = False: timer2.Enabled = False: cmdscroll.Enabled = False: lblseconds.Caption = 60
chattext "···• The SoLuTioN MM Bot has been ended early [User]": DoEvents
VBMsg1.SubClasshWnd = 0
End Sub

Sub cmdBotOn_Click ()
If txtadd.Text = "" Then MsgBox "No addition phrase has been detected!", 48: Exit Sub
If cmdscroll.Enabled = True Then MsgBox "The Mail Manager is already active!", 48: Exit Sub
Room% = findChathWnd%()
If Room% = 0 Then MsgBox "Your chatroom has not been detected, ..exiting", 32: Exit Sub
vscroll1.Enabled = False: txtadd.Enabled = False: txtrmv.Enabled = False: timer1.Enabled = True: timer2.Enabled = True: cmdscroll.Enabled = True: lblseconds.Caption = 60
chattext "···• SoLuTioN Mailing System Mail Manager Has Begun!": DoEvents
chattext "···• Type ''" & txtadd.Text & "'' to be added": DoEvents
If txtrmv.Text <> "" Then chattext "···• Or ''" & txtrmv.Text & "'' to be removed": DoEvents
View% = findchildbyclass(Room%, "_AOL_VIEW")
VBMsg1.SubClasshWnd = View%
End Sub

Sub cmdOptions_Click ()
Me.Enabled = False
frmOptions.Show
End Sub

Sub cmdRMVName_Click ()
On Error GoTo errhand
lstPeople.RemoveItem lstPeople.ListIndex
frappl.Caption = "People On: " & lstPeople.ListCount
txtsn.Text = ""
txtsn.SetFocus
errhand:
Resume Next
End Sub

Sub cmdStartMM_Click ()
If lstPeople.ListCount = 0 Then MsgBox "You have chosen to start the MM, yet you have no people added... are you stupid?", 32: Exit Sub

If frmOptions.optCurrent.Value = True Then Call SendCurrent: DoEvents: Exit Sub
If frmOptions.optAll.Value = True Then Call SendAll: DoEvents: Exit Sub
If frmOptions.optRange.Value = True Then Call SendRange: DoEvents: Exit Sub
End Sub

Sub cmdUnload_Click ()
rEND = MsgBox("Are you sure you want to close the Window and lose everything that isn't saved?", 36)
If rEND = 7 Then Exit Sub
timer1.Enabled = False: timer2.Enabled = False
If cmdscroll.Enabled = True Then chattext "···• SoLuTioN MaiLeR Window has been closed, Bot has been shutdown": DoEvents
Unload Me
End Sub

Sub Form_GotFocus ()
stayontop Me
End Sub

Sub Form_Load ()
lstPeople.Height = 1080
stayontop Me
frm% = Me.hWnd: x = setfocusapi(frm%)
End Sub

Sub Form_LostFocus ()
stayontop Me
End Sub

Sub Form_Resize ()
stayontop Me
End Sub

Sub lstPeople_Click ()
txtsn.Text = lstPeople.Text
End Sub

Sub mnuAddRoom_Click ()
Call AddRoom(lstPeople, False): DoEvents
frappl.Caption = "People On: " & lstPeople.ListCount
End Sub

Sub mnuClearList_Click ()
lstPeople.Clear
frappl.Caption = "People On: 0"
End Sub

Sub mnuLoadList_Click ()
cmdialog1.DialogTitle = "Open a SoLuTioN List File"
cmdialog1.Filter = "Solution List Files|*.SoL"
cmdialog1.Action = 1
If cmdialog1.Filename = "" Then Exit Sub
Open cmdialog1.Filename For Input As #1
Do While Not EOF(1)
Line Input #1, sn$
Call addlist(lstPeople, sn$): DoEvents
Loop
DoEvents
frappl.Caption = "People On: " & lstPeople.ListCount
End Sub

Sub mnuSaveList_Click ()
If lstPeople.ListCount = 0 Then MsgBox "You have nothing to put in the file", 48: Exit Sub
cmdialog1.DialogTitle = "Save a SoLuTioN List File"
cmdialog1.Filter = "Solution List Files|*.SoL"
cmdialog1.Action = 2
If cmdialog1.Filename = "" Then Exit Sub
Close #1
Open cmdialog1.Filename For Output As #1
Close #1
Open cmdialog1.Filename For Output As #1
For x = 0 To lstPeople.ListCount - 1
Print #1, lstPeople.List(x)
Next x
Close #1
DoEvents
End Sub

Sub Timer1_Timer ()
lblseconds.Caption = lblseconds.Caption - 1
If lblseconds.Caption = 1 Then
lblminutes.Caption = lblminutes.Caption - 1
chattext "···• SoLuTioN Mailing System Mailing in: " & lblminutes.Caption & " Minute[s]": DoEvents
chattext "···• Type ''" & txtadd.Text & "'' to be added": DoEvents
chattext "···• Or ''" & txtrmv.Text & "'' to be removed": DoEvents
lblseconds.Caption = 60
End If
End Sub

Sub Timer2_Timer ()
If lstIM.ListCount = 0 Then Exit Sub
Call sendim((lstIM.List(0)), "Added you to the MM i did, MM addition is now yours!"): DoEvents
Call ChatSay("SoLuTioN HoST", lstIM.List(0) & " was added to the MM list"): DoEvents
lstIM.RemoveItem 0
End Sub

Sub txtSN_KeyPress (keyascii As Integer)
If keyascii = 13 Then
cmdaddname_click
End If

End Sub

Sub VBMsg1_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lparam As Long, RetVal As Long, CallDefProc As Integer)
Message$ = agGetStringFromLPSTR$(lparam)
sn$ = Mid$(Message$, 3, InStr(Message$, ":") - 3)
txt$ = Mid$(Message$, InStr(Message$, ":") + 2)

If UCase(txt$) = UCase(txtadd.Text) Then
Call addlist(lstPeople, (sn$)): DoEvents: frappl.Caption = "People On: " & lstPeople.ListCount
If frmOptions.chkNotify.Value = True Then Call addlist(lstIM, (sn$)): DoEvents
Exit Sub
End If

If UCase(txt$) = UCase(txtrmv.Text) Then
Call removeitemlst(lstPeople, (sn$)): DoEvents: frappl.Caption = "People On: " & lstPeople.ListCount
Exit Sub
End If
End Sub

Sub VScroll1_Change ()
lblminutes.Caption = vscroll1.Value
lblminutes.Refresh
End Sub

