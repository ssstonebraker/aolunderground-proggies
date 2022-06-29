VERSION 4.00
Begin VB.Form Form15 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Punter"
   ClientHeight    =   2475
   ClientLeft      =   1575
   ClientTop       =   1935
   ClientWidth     =   3660
   Height          =   2880
   Left            =   1515
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Top             =   1590
   Width           =   3780
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   225
      Left            =   1950
      TabIndex        =   25
      Top             =   720
      Width           =   555
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Punt N"
      Height          =   195
      Left            =   2370
      TabIndex        =   24
      Top             =   30
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   1380
      Top             =   3090
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Story"
      Height          =   195
      Left            =   3000
      TabIndex        =   15
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Rndm"
      Height          =   195
      Left            =   3000
      TabIndex        =   14
      Top             =   30
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Punt B"
      Height          =   195
      Left            =   2370
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Blank"
      Height          =   195
      Left            =   2370
      TabIndex        =   11
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   450
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1710
      TabIndex        =   2
      Top             =   330
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1620
      TabIndex        =   23
      Top             =   720
      Width           =   1965
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   780
      TabIndex        =   22
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "Length"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3150
      Width           =   945
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   225
      Left            =   2070
      TabIndex        =   19
      Top             =   3120
      Width           =   1425
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Secs"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3210
      TabIndex        =   18
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2490
      TabIndex        =   17
      Top             =   990
      Width           =   585
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "In"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2250
      TabIndex        =   16
      Top             =   990
      Width           =   195
   End
   Begin RichtextLib.RichTextBox RichTextBox1 
      Height          =   1125
      Left            =   30
      TabIndex        =   12
      Top             =   1260
      Width           =   3585
      _Version        =   65536
      _ExtentX        =   6324
      _ExtentY        =   1984
      _StockProps     =   69
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScrollBars      =   3
      TextRTF         =   $"prog16.frx":0000
      RightMargin     =   3150
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1620
      TabIndex        =   10
      Top             =   990
      Width           =   525
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Out of:"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   9
      Top             =   990
      Width           =   555
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   540
      TabIndex        =   8
      Top             =   990
      Width           =   465
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Sent:"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2610
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   2610
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Number"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1710
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Who to punt?"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1425
   End
End
Attribute VB_Name = "Form15"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
RichTextBox1.text = "‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡‡"
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Sends A Blank Im"
End Sub


Private Sub Command2_Click()
Form15.Hide
Form4.Show
End Sub


Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Hmmm, Gee I wonder..."
End Sub


Private Sub Command3_Click()
RichTextBox1.text = String(1023, 13)
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Use this if your modem sux"
End Sub


Private Sub Command4_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindWindowEx(AOL%, 0, "MDIClient", vbNullString)
ass = 1
'K this is opening the first series # of IM boxes
Counter = 0
Dim Series As Integer
Series = Text2.text
Do: DoEvents
Counter = Counter + 1
Call RunMenuByString(AOL%, "Send an instant message")
bobo = 0
Do: DoEvents
bobo = bobo + 1
Loop Until bobo = 10
Do: DoEvents
IM% = FindWindowEx(MDI%, 0, "AOL Child", "Send Instant Message")
Loop Until IM% <> 0
Label9.Caption = Series

If Counter = 1 Then
aaa% = IM%
gen = aaa%
End If
If Counter = 2 Then
aab% = IM%
gen = aab%
End If
If Counter = 3 Then
aac% = IM%
gen = aac%
End If
If Counter = 4 Then
aad% = IM%
gen = aad%
End If
If Counter = 5 Then
aae% = IM%
gen = aae%
End If
If Counter = 6 Then
aaf% = IM%
gen = aaf%
End If
If Counter = 7 Then
aag% = IM%
gen = aag%
End If
If Counter = 8 Then
aah% = IM%
gen = aah%
End If
If Counter = 9 Then
aai% = IM%
gen = aai%
End If
If Counter = 10 Then
aaj% = IM%
gen = aaj%
End If
If Counter = 11 Then
aak% = IM%
gen = aak%
End If
If Counter = 12 Then
aal% = IM%
gen = aal%
End If
If Counter = 13 Then
aam% = IM%
gen = aam%
End If
If Counter = 14 Then
aan% = IM%
gen = aan%
End If
If Counter = 15 Then
aao% = IM%
gen = aao%
End If
If Counter = 16 Then
aap% = IM%
gen = aap%
End If
If Counter = 17 Then
aaq% = IM%
gen = aaq%
End If
If Counter = 18 Then
aar% = IM%
gen = aar%
End If
If Counter = 19 Then
aas% = IM%
gen = aas%
End If
If Counter = 20 Then
aat% = IM%
gen = aat%
End If
If Counter = 21 Then
aau% = IM%
gen = aau%
End If
If Counter = 22 Then
aav% = IM%
gen = aav%
End If
If Counter = 23 Then
aaw% = IM%
gen = aaw%
End If
If Counter = 24 Then
aax% = IM%
gen = aax%
End If

'Finding to who box and text box, and icon for all windows
'To who?
Do: DoEvents
Edit% = FindWindowEx(IM%, 0, "_AOL_EDIT", vbNullString)
Loop Until Edit% <> 0

Do: DoEvents
Call AOLSetText(Edit%, Text1.text)
EText = AOLGetText(Edit%)
Loop Until EText = Text1.text
Call sendtext(IM%, "IM" & ass)
ass = ass + 1
'Text box
Do: DoEvents
RICH% = FindWindowEx(IM%, 0, "RICHCNTL", vbNullString)
Loop Until RICH% <> 0

Do: DoEvents
Call AOLSetText(RICH%, RichTextBox1.text)
nText = AOLGetText(RICH%)
Loop Until nText = RichTextBox1.text

Do: DoEvents
hChild = GetWindow(RICH%, GW_HWNDNEXT)
Loop Until hChild <> 0
but% = hChild

If Counter = 1 Then
but1% = but%
End If
If Counter = 2 Then
but2% = but%
End If
If Counter = 3 Then
but3% = but%
End If
If Counter = 4 Then
but4% = but%
End If
If Counter = 5 Then
but5% = but%
End If
If Counter = 6 Then
but6% = but%
End If
If Counter = 7 Then
but7% = but%
End If
If Counter = 8 Then
but8% = but%
End If
If Counter = 9 Then
but9% = but%
End If
If Counter = 10 Then
but10% = but%
End If
If Counter = 11 Then
but11% = but%
End If
If Counter = 12 Then
but12% = but%
End If
If Counter = 13 Then
but13% = but%
End If
If Counter = 14 Then
but14% = but%
End If
If Counter = 15 Then
but15% = but%
End If
If Counter = 16 Then
but16% = but%
End If
If Counter = 17 Then
but17% = but%
End If
If Counter = 18 Then
but18% = but%
End If
If Counter = 19 Then
but19% = but%
End If
If Counter = 20 Then
but20% = but%
End If
If Counter = 21 Then
but21% = but%
End If
If Counter = 22 Then
but22% = but%
End If
If Counter = 23 Then
but23% = but%
End If
Loop Until Counter = Series

X = 0
Label7.Caption = X


If Counter > 0 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM1")
If IMs <> 0 Then
Call Sendclick(but1%)
End If
IMs = FindChildByTitle(MDI%, "IM1")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If



If Counter > 1 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM2")
If IMs <> 0 Then
Call Sendclick(but2%)
End If
IMs = FindChildByTitle(MDI%, "IM2")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 2 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM3")
If IMs <> 0 Then
Call Sendclick(but3%)
End If
IMs = FindChildByTitle(MDI%, "IM3")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 3 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM4")
If IMs <> 0 Then
Call Sendclick(but4%)
End If
IMs = FindChildByTitle(MDI%, "IM4")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 4 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM5")
If IMs <> 0 Then
Call Sendclick(but5%)
End If
IMs = FindChildByTitle(MDI%, "IM5")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 5 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM6")
If IMs <> 0 Then
Call Sendclick(but6%)
End If
IMs = FindChildByTitle(MDI%, "IM6")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 6 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM7")
If IMs <> 0 Then
Call Sendclick(but7%)
End If
IMs = FindChildByTitle(MDI%, "IM7")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 7 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM8")
If IMs <> 0 Then
Call Sendclick(but8%)
End If
IMs = FindChildByTitle(MDI%, "IM8")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 8 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM9")
If IMs <> 0 Then
Call Sendclick(but9%)
End If
IMs = FindChildByTitle(MDI%, "IM9")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 9 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM10")
If IMs <> 0 Then
Call Sendclick(but10%)
End If
IMs = FindChildByTitle(MDI%, "IM10")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 10 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM11")
If IMs <> 0 Then
Call Sendclick(but11%)
End If
IMs = FindChildByTitle(MDI%, "IM11")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 11 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM12")
If IMs <> 0 Then
Call Sendclick(but12%)
End If
IMs = FindChildByTitle(MDI%, "IM12")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 12 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM13")
If IMs <> 0 Then
Call Sendclick(but13%)
End If
IMs = FindChildByTitle(MDI%, "IM13")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 13 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM14")
If IMs <> 0 Then
Call Sendclick(but14%)
End If
IMs = FindChildByTitle(MDI%, "IM14")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 14 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM15")
If IMs <> 0 Then
Call Sendclick(but15%)
End If
IMs = FindChildByTitle(MDI%, "IM15")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 15 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM16")
If IMs <> 0 Then
Call Sendclick(but16%)
End If
IMs = FindChildByTitle(MDI%, "IM16")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 16 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM17")
If IMs <> 0 Then
Call Sendclick(but17%)
End If
IMs = FindChildByTitle(MDI%, "IM17")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 17 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM18")
If IMs <> 0 Then
Call Sendclick(but18%)
End If
IMs = FindChildByTitle(MDI%, "IM18")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 18 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM19")
If IMs <> 0 Then
Call Sendclick(but19%)
End If
IMs = FindChildByTitle(MDI%, "IM19")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 19 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM19")
If IMs <> 0 Then
Call Sendclick(but20%)
End If
IMs = FindChildByTitle(MDI%, "IM20")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 20 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM20")
If IMs <> 0 Then
Call Sendclick(but21%)
End If
IMs = FindChildByTitle(MDI%, "IM21")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 21 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM22")
If IMs <> 0 Then
Call Sendclick(but22%)
End If
IMs = FindChildByTitle(MDI%, "IM22")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If


If Counter > 22 Then
Do: DoEvents
IMs = FindChildByTitle(MDI%, "IM23")
If IMs <> 0 Then
Call Sendclick(but23%)
End If
IMs = FindChildByTitle(MDI%, "IM23")
Pass% = FindWindow("#32770", "America Online")
If Pass% <> 0 Then
Exit Sub
End If
Loop Until IMs = 0
X = X + 1
Label7.Caption = X
End If
Timer1.Enabled = False
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Goes fast w/ good modem"
End Sub


Private Sub Command5_Click()
Counter = 0
Do: DoEvents
MyValue$ = Int((9 * Rnd) + 1)

Mytotal$ = Mytotal$ + MyValue$
Counter = Counter + 1
Loop Until Counter = 500
RichTextBox1.text = Mytotal$
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Sends 500 Random #'s"
End Sub


Private Sub Command6_Click()
RichTextBox1.text = "Hey what's up, Wow haven't seen you in forever how are you these days anyway? Still hanging around with Rhonda and All the rest of them? I saw her a little while ago, She asked how you were and everything, it was just like old times...So you and your band still together or what? I remember how we used to party before everything got messed up, Man I have a question are these IM's doing anything weird?!? Someone told me a little while ago that when I Imed them it sent 10 in a row!!!!! Pretty weird"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Good Tos Protection"
End Sub


Private Sub Command7_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
ass = 1
'K this is opening the first series # of IM boxes
Counter = 0
Dim Series As Integer
Series = Text2.text

Do: DoEvents
    Counter = Counter + 1
    Call RunMenuByString(AOL%, "Send an instant message")
    bobo = 0

    Do: DoEvents
        bobo = bobo + 1
    Loop Until bobo = 10

    Do: DoEvents
        IM% = FindChildByTitle(MDI%, "Send Instant Message")
    Loop Until IM% <> 0

    Label9.Caption = Series

    If Counter = 1 Then
        aaa% = IM%
        gen = aaa%
    End If

    If Counter = 2 Then
        aab% = IM%
        gen = aab%
    End If

    If Counter = 3 Then
        aac% = IM%
        gen = aac%
    End If

    If Counter = 4 Then
        aad% = IM%
        gen = aad%
    End If

    If Counter = 5 Then
        aae% = IM%
        gen = aae%
    End If

    If Counter = 6 Then
        aaf% = IM%
        gen = aaf%
    End If

    If Counter = 7 Then
        aag% = IM%
        gen = aag%
    End If

    If Counter = 8 Then
        aah% = IM%
        gen = aah%
    End If

    If Counter = 9 Then
        aai% = IM%
        gen = aai%
    End If

    If Counter = 10 Then
        aaj% = IM%
        gen = aaj%
    End If

    If Counter = 11 Then
        aak% = IM%
        gen = aak%
    End If

    If Counter = 12 Then
        aal% = IM%
        gen = aal%
    End If

    If Counter = 13 Then
        aam% = IM%
        gen = aam%
    End If

    If Counter = 14 Then
        aan% = IM%
        gen = aan%
    End If

    If Counter = 15 Then
        aao% = IM%
        gen = aao%
    End If

    If Counter = 16 Then
        aap% = IM%
        gen = aap%
    End If

    If Counter = 17 Then
        aaq% = IM%
        gen = aaq%
    End If

    If Counter = 18 Then
        aar% = IM%
        gen = aar%
    End If

    If Counter = 19 Then
        aas% = IM%
        gen = aas%
    End If

    If Counter = 20 Then
        aat% = IM%
        gen = aat%
    End If

    If Counter = 21 Then
        aau% = IM%
        gen = aau%
    End If

    If Counter = 22 Then
        aav% = IM%
        gen = aav%
    End If

    If Counter = 23 Then
        aaw% = IM%
        gen = aaw%
    End If

    If Counter = 24 Then
        aax% = IM%
        gen = aax%
    End If

'Finding to who box and text box, and icon for all windows
'To who?
    hChild = GetWindow(gen, GW_CHILD)
    rooms = GetClass(hChild)
    Label4.Caption = rooms

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Edit"

    Do: DoEvents
        Call AOLSetText(hChild, Text1.text)
        EText = AOLGetText(hChild)
    Loop Until EText = Text1.text

    Call sendtext(IM%, "IM" & ass)
    ass = ass + 1
    'Text box

    hChild = GetWindow(gen, GW_CHILD)
    rooms = GetClass(hChild)
    Label4.Caption = rooms

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "RICHCNTL"

    Do: DoEvents
        Call AOLSetText(hChild, RichTextBox1.text)
        nText = AOLGetText(hChild)
    Loop Until nText = RichTextBox1.text

    'Send button
    hChild = GetWindow(gen, GW_CHILD)
    rooms = GetClass(hChild)
    Label4.Caption = rooms

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Bob = GetClass(hChild)
        Label4.Caption = Bob
        Y = aolhwnd = hChild
    Loop Until Label4.Caption = "_AOL_Icon"

    but% = hChild

    If Counter = 1 Then
        but1% = but%
    End If

    If Counter = 2 Then
        but2% = but%
    End If

    If Counter = 3 Then
        but3% = but%
    End If

    If Counter = 4 Then
        but4% = but%
    End If
    
    If Counter = 5 Then
        but5% = but%
    End If
    
    If Counter = 6 Then
        but6% = but%
    End If

    If Counter = 7 Then
        but7% = but%
    End If

    If Counter = 8 Then
        but8% = but%
    End If
    
    If Counter = 9 Then
        but9% = but%
    End If

    If Counter = 10 Then
        but10% = but%
    End If
    
    If Counter = 11 Then
        but11% = but%
    End If

    If Counter = 12 Then
        but12% = but%
    End If

    If Counter = 13 Then
        but13% = but%
    End If
    
    If Counter = 14 Then
        but14% = but%
    End If
    
    If Counter = 15 Then
        but15% = but%
    End If
    
    If Counter = 16 Then
        but16% = but%
    End If
    
    If Counter = 17 Then
        but17% = but%
    End If
    
    If Counter = 18 Then
        but18% = but%
    End If
    
    If Counter = 19 Then
        but19% = but%
    End If

    If Counter = 20 Then
        but20% = but%
    End If
    
    If Counter = 21 Then
        but21% = but%
    End If
    
    If Counter = 22 Then
        but22% = but%
    End If

    If Counter = 23 Then
        but23% = but%
    End If

Loop Until Counter = Series

X = 0
Label7.Caption = X


Do: DoEvents

If Counter > 0 Then
IMs = FindChildByTitle(MDI%, "IM1")
    If IMs <> 0 Then
    Call Sendclick(but1%)
    Else
        If aa = 0 Then
            X = X + 1
            Label7.Caption = X
            aa = 1
        End If
    End If
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If



If Counter > 1 Then
    IMs = FindChildByTitle(MDI%, "IM2")
    If IMs <> 0 Then
        Call Sendclick(but2%)
        Else
        If ab = 0 Then
            X = X + 1
            Label7.Caption = X
            ab = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM2")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 2 Then
    IMs = FindChildByTitle(MDI%, "IM3")
    If IMs <> 0 Then
        Call Sendclick(but3%)
        Else
        If ac = 0 Then
            X = X + 1
            Label7.Caption = X
            ac = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM3")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 3 Then
    IMs = FindChildByTitle(MDI%, "IM4")
    If IMs <> 0 Then
        Call Sendclick(but4%)
        Else
        If ad = 0 Then
            X = X + 1
            Label7.Caption = X
            ad = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM4")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 4 Then
    IMs = FindChildByTitle(MDI%, "IM5")
    If IMs <> 0 Then
        Call Sendclick(but5%)
        Else
        If ae = 0 Then
            X = X + 1
            Label7.Caption = X
            ae = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM5")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 5 Then
    IMs = FindChildByTitle(MDI%, "IM6")
    If IMs <> 0 Then
        Call Sendclick(but6%)
        Else
        If af = 0 Then
            X = X + 1
            Label7.Caption = X
            af = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM6")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 6 Then
    IMs = FindChildByTitle(MDI%, "IM7")
    If IMs <> 0 Then
        Call Sendclick(but7%)
        Else
        If ag = 0 Then
            X = X + 1
            Label7.Caption = X
            ag = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM7")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 7 Then
    IMs = FindChildByTitle(MDI%, "IM8")
    If IMs <> 0 Then
        Call Sendclick(but8%)
        Else
        If ah = 0 Then
            X = X + 1
            Label7.Caption = X
            ah = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM8")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 8 Then
    IMs = FindChildByTitle(MDI%, "IM9")
    If IMs <> 0 Then
        Call Sendclick(but9%)
        Else
        If ai = 0 Then
            X = X + 1
            Label7.Caption = X
            ai = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM9")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 9 Then
    IMs = FindChildByTitle(MDI%, "IM10")
    If IMs <> 0 Then
        Call Sendclick(but10%)
        Else
        If aj = 0 Then
            X = X + 1
            Label7.Caption = X
            aj = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM10")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 10 Then
    IMs = FindChildByTitle(MDI%, "IM11")
    If IMs <> 0 Then
        Call Sendclick(but11%)
        Else
        If ak = 0 Then
            X = X + 1
            Label7.Caption = X
            ak = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM11")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 11 Then
    IMs = FindChildByTitle(MDI%, "IM12")
    If IMs <> 0 Then
        Call Sendclick(but12%)
        Else
        If al = 0 Then
            X = X + 1
            Label7.Caption = X
            al = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM12")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 12 Then
    IMs = FindChildByTitle(MDI%, "IM13")
    If IMs <> 0 Then
        Call Sendclick(but13%)
        Else
        If am = 0 Then
            X = X + 1
            Label7.Caption = X
            am = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM13")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 13 Then
    IMs = FindChildByTitle(MDI%, "IM14")
    If IMs <> 0 Then
        Call Sendclick(but14%)
        Else
        If an = 0 Then
            X = X + 1
            Label7.Caption = X
            an = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM14")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 14 Then
    IMs = FindChildByTitle(MDI%, "IM15")
    If IMs <> 0 Then
        Call Sendclick(but15%)
        Else
        If ao = 0 Then
            X = X + 1
            Label7.Caption = X
            ao = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM15")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 15 Then
    IMs = FindChildByTitle(MDI%, "IM16")
    If IMs <> 0 Then
        Call Sendclick(but16%)
        Else
        If ap = 0 Then
            X = X + 1
            Label7.Caption = X
            ap = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM16")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 16 Then
    IMs = FindChildByTitle(MDI%, "IM17")
    If IMs <> 0 Then
        Call Sendclick(but17%)
        Else
        If aq = 0 Then
            X = X + 1
            Label7.Caption = X
            aq = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM17")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 17 Then
    IMs = FindChildByTitle(MDI%, "IM18")
    If IMs <> 0 Then
        Call Sendclick(but18%)
        Else
        If ar = 0 Then
            X = X + 1
            Label7.Caption = X
            ar = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM18")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 18 Then
    IMs = FindChildByTitle(MDI%, "IM19")
    If IMs <> 0 Then
        Call Sendclick(but19%)
        Else
        If asr = 0 Then
            X = X + 1
            Label7.Caption = X
            asr = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM19")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 19 Then
    IMs = FindChildByTitle(MDI%, "IM19")
    If IMs <> 0 Then
        Call Sendclick(but20%)
        Else
        If at = 0 Then
            X = X + 1
            Label7.Caption = X
            at = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM20")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 20 Then
    IMs = FindChildByTitle(MDI%, "IM20")
    If IMs <> 0 Then
        Call Sendclick(but21%)
        Else
        If au = 0 Then
            X = X + 1
            Label7.Caption = X
            au = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM21")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 21 Then
    IMs = FindChildByTitle(MDI%, "IM22")
    If IMs <> 0 Then
        Call Sendclick(but22%)
        Else
        If av = 0 Then
            X = X + 1
            Label7.Caption = X
            av = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM22")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If


If Counter > 22 Then
    IMs = FindChildByTitle(MDI%, "IM23")
    If IMs <> 0 Then
        Call Sendclick(but23%)
        Else
        If aw = 0 Then
            X = X + 1
            Label7.Caption = X
            aw = 1
        End If
    End If
    IMs = FindChildByTitle(MDI%, "IM23")
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
        Exit Sub
    End If
End If
Loop Until aa + ab + ac + ad + ae + af + ag + ah + ai + aj + ak + al + am + an + ao + ap + aq + ar + asr + at + au + av + aw = 23
End Sub

Private Sub Form_DblClick()
Msg = "K, first of all AOL just let you have something like 25 windows open at once...yes including that damn Welcome window, so it help if you close all window you can if you want to send 23, Also before you punt somebody, go to locate member, enter there SN and see the little box that say's {Send IM}? Well if you can press that it means they aren't ghosting, and there IM's are on, if not well your screwed I haven't bothered to make a pinter to much Frigging work :)...well that's about it, ohh yeah Punt N kills wait before it sends, it kinda sux, I recomend using Punt B...unless you are on like a 2400, then you might as well give up using the punting idea, but Punt N will probably work better...(like it will help...heh heh) Well I'm done writing this book so L8r"
Style = vbOKOnly
Title = "Help for all you wennies :)"
Response = MsgBox(Msg, Style, Title)
End Sub

Private Sub Form_Load()
Call StayOnTop(Form15)
Timer1.Enabled = False
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Dblclick for hints"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Pretty, text..."
End Sub


Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "How much time it's taken"
End Sub


Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Secs = Seconds...duh :)"
End Sub


Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Length of Text in box"
End Sub


Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Length of Text in box"
End Sub


Private Sub Label16_DblClick()
Label16.Caption = "LOL you lamo,no secrets"
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Dblclick for hidden area!!!!!"
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "How many to send?"
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "How much time it's taken"
End Sub


Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "How many have sent"
End Sub


Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "How many have sent"
End Sub


Private Sub label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Total number to send"
End Sub


Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Total number to send"
End Sub


Private Sub RichTextBox1_Change()
Z = Len(RichTextBox1.text)
Label15.Caption = Z
If Z = 500 Then
MsgBox "This is the limit on how much text a im will hold, if you add any more it won't work"
End If
If Z > 500 Then
w = Z - 500
MsgBox "Delete " & w & " Characters or else this won't work"
End If
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Text you send the victim"
End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Enter SN of the victim"
End Sub


Private Sub text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Send 23 or less"
End Sub


Private Sub Timer1_Timer()
Label10.Caption = Time
End Sub


