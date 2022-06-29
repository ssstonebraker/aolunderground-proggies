VERSION 5.00
Object = "{04E76B5B-D725-11D2-B3D8-4481F5C00000}#11.0#0"; "DOGBARV2.OCX"
Begin VB.Form Form1 
   Caption         =   "Auto Upchat Example By: 2ooo"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3495
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1545
      TabIndex        =   10
      Text            =   "File"
      Top             =   1530
      Width           =   1905
   End
   Begin VB.CommandButton Command4 
      Caption         =   "not on top"
      Height          =   300
      Left            =   2370
      TabIndex        =   8
      Top             =   1965
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   " on top"
      Height          =   285
      Left            =   45
      TabIndex        =   7
      Top             =   1950
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   75
      TabIndex        =   5
      Text            =   "upload time remaining"
      Top             =   1140
      Width           =   3405
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   705
      Top             =   585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ON"
      Height          =   300
      Left            =   270
      TabIndex        =   0
      Top             =   285
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Upload Status"
      Height          =   975
      Left            =   1890
      TabIndex        =   3
      Top             =   60
      Width           =   1560
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   870
         TabIndex        =   6
         Text            =   "0"
         Top             =   660
         Visible         =   0   'False
         Width           =   690
      End
      Begin DogBarProgressBarv2.DogBar DogBar1 
         Height          =   405
         Left            =   105
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BarStyle        =   0
         BarColor1       =   16777215
         BarColor2       =   8388608
         BackColor       =   -2147483644
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   240
         Left            =   495
         TabIndex        =   4
         Top             =   630
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto upchat"
      Height          =   990
      Left            =   105
      TabIndex        =   1
      Top             =   60
      Width           =   1800
      Begin VB.Timer Timer3 
         Interval        =   100
         Left            =   1185
         Top             =   540
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   165
         Top             =   540
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OFF"
         Height          =   300
         Left            =   945
         TabIndex        =   2
         Top             =   210
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      Caption         =   "File Being uploaded:"
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   1545
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub GETFILENAME2()
If find_aolchild() <> 0 Then
  getfilename
 Else
Exit Sub
 End If
End Sub
Public Function find_aolchild() As Long
' If this function finds the window, it will return it's
' handle. If it doesn't find it, it will return 0.
Dim aolframe&
Dim mdiclient&
Dim aolchild&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
Dim Winkid1&, Winkid2&, Winkid3&, Winkid4&, Winkid5&, Winkid6&, Winkid7&, Winkid8&, Winkid9&, FindOtherWin&
FindOtherWin& = GetWindow(aolchild&, GW_HWNDFIRST)
Do While FindOtherWin& <> 0
       DoEvents
       Winkid1& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
       Winkid2& = FindWindowEx(FindOtherWin&, 0&, "_aol_edit", vbNullString)
       Winkid3& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
       Winkid4& = FindWindowEx(FindOtherWin&, 0&, "_aol_edit", vbNullString)
       Winkid5& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
       Winkid6& = FindWindowEx(FindOtherWin&, 0&, "_aol_edit", vbNullString)
       Winkid7& = FindWindowEx(FindOtherWin&, 0&, "_aol_fontcombo", vbNullString)
       Winkid8& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
       Winkid9& = FindWindowEx(FindOtherWin&, 0&, "_aol_combobox", vbNullString)
       If (Winkid1& <> 0) And (Winkid2& <> 0) And (Winkid3& <> 0) And (Winkid4& <> 0) And (Winkid5& <> 0) And (Winkid6& <> 0) And (Winkid7& <> 0) And (Winkid8& <> 0) And (Winkid9& <> 0) Then
              find_aolchild = FindOtherWin&
              Exit Function
       End If
       FindOtherWin& = GetWindow(FindOtherWin&, GW_HWNDNEXT)
Loop
find_aolchild = 0

End Function

Private Sub Command1_Click()
 Command1.Enabled = False
 Command2.Enabled = True
 Timer2.Enabled = True
 Text1.Text = "Status : ON"
 End Sub


Private Sub Command2_Click()
Command1.Enabled = True
 Command2.Enabled = False
 Timer1.Enabled = False
 Timer2.Enabled = False
 Label1.Caption = "0"
 Text2.Text = "0"
 DogBar1.Value = "0"
 Text1.Text = "Status:OFF"
 Text3.Text = "File"
 
End Sub
Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Private Sub Command3_Click()
StayOnTop Me
End Sub
Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Sub NotOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Private Sub Command4_Click()
NotOnTop Me
End Sub

Private Sub Command5_Click()
StayOnTop Me
Me.WindowState = 1
End Sub

Private Sub Form_Click()
 Me.WindowState = 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If find_aolmodal() <> 0 Then
   miniuploadwin
   disableuploadwin
Call Pause(1#) ' to let the upload win minimize after that its useless
  GetWinTxt
 Form1.Text2.Text = Label1.Caption
Form1.Text2.Text = Left(Text2.Text, 14)
Form1.Text2.Text = Left(Text2.Text, 2)
 Form1.DogBar1.Value = Form1.Text2
 GETFILENAME2
 getstatictxt
 Else
 Timer3.Enabled = True
 Timer1.Enabled = False
  End If
 Exit Sub
 End Sub

Private Sub Timer2_Timer()
If find_aolmodal() <> 0 Then
   Timer1.Enabled = True
     Else
  Timer1.Enabled = False
  Timer3.Enabled = False
  Label1.Caption = "0"
  Text2.Text = "0"
  Text3.Text = "File"
   End If
 End Sub
Sub getfilename()
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aolstatic&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolstatic& = FindWindowEx(aolchild&, 0&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_aol_static", vbNullString)
Dim TheText$, TL As Long
TL = SendMessageLong(aolstatic&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(aolstatic&, WM_GETTEXT, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
Form1.Text3.Text = (TheText$)
End Sub

Private Sub Timer3_Timer()
 If Text2.Text = "00" Then
 Form1.DogBar1.Value = "0"
 Label1.Caption = "0"
 Text1.Text = "Status:ON"
 Text3.Text = "File"
End If
End Sub
