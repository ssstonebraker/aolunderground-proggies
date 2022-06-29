VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Target Range"
   ClientHeight    =   5985
   ClientLeft      =   2160
   ClientTop       =   1770
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   6000
   End
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   6645
      Picture         =   "Form1.frx":934DA
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   4070
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   3390
      Picture         =   "Form1.frx":9815C
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   7
      Top             =   4070
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   120
      Picture         =   "Form1.frx":9CDDE
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   4070
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   6645
      Picture         =   "Form1.frx":A1A60
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   2310
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   3390
      Picture         =   "Form1.frx":A66E2
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   2310
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   120
      Picture         =   "Form1.frx":AB364
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   2310
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   6640
      Picture         =   "Form1.frx":AFFE6
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   550
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   3390
      Picture         =   "Form1.frx":B4C68
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   550
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1205
      Left            =   120
      Picture         =   "Form1.frx":B98EA
      ScaleHeight     =   1200
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   550
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "6"
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Time:"
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Misses:"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Hits:"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Shots:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   495
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu newgame 
         Caption         =   "New Game"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu times 
         Caption         =   "Time's"
         Begin VB.Menu beginner 
            Caption         =   "Beginner Time"
            Checked         =   -1  'True
         End
         Begin VB.Menu intermediate 
            Caption         =   "Intermediate Time"
         End
         Begin VB.Menu expert 
            Caption         =   "Expert Time"
         End
         Begin VB.Menu custom 
            Caption         =   "Custom Time"
         End
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub beginner_Click()
Text2.Text = "t"
Label8.Caption = "6"
beginner.Checked = True
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
End Sub

Private Sub custom_Click()
Form2.Show
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = True
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub expert_Click()
Text2.Text = "t"
Label8.Caption = "2"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = True
custom.Checked = False
End Sub

Private Sub Form_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
If Label5.Caption = "" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label7.Caption = "0" Then
Label7.Caption = "1"
Else
Label7.Caption = Label7.Caption - -"1"
End If
End Sub

Private Sub Form_Load()
Text2.Text = "t"
End Sub

Private Sub intermediate_Click()
Text2.Text = "t"
Label8.Caption = "4"
beginner.Checked = False
intermediate.Checked = True
expert.Checked = False
custom.Checked = False
End Sub

Private Sub newgame_Click()
If Text2.Text = "" Then
MsgBox "You have to select a time first.", vbOKOnly, ""
Exit Sub
Else
beginner.Enabled = False
intermediate.Enabled = False
expert.Enabled = False
custom.Enabled = False
Text1.Text = "n"
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
Label4.Caption = "Time Left:"
Label5.Caption = "0"
Label6.Caption = "0"
Label7.Caption = "0"
MsgBox "Press ok when you are ready to shoot!", vbOKOnly, ""
Timer1.Enabled = True
End If
End Sub

Private Sub Picture1_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture1.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture2_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture2.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture3_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture3.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture4_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture4.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture5_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture5.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture6_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture6.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture7_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture7.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture8_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture8.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Picture9_Click()
If Text1.Text = "" Then
Exit Sub
End If
Call Playwav("C:\unzipped\Target Range\Sounds\GUN.WAV")
Picture9.Visible = False
If Label5.Caption = "0" Then
Label5.Caption = "1"
Else
Label5.Caption = Label5.Caption - -"1"
End If
If Label6.Caption = "0" Then
Label6.Caption = "1"
Else
Label6.Caption = Label6.Caption - -"1"
End If
If Picture1.Visible = False And Picture2.Visible = False And Picture3.Visible = False And Picture4.Visible = False And Picture5.Visible = False And Picture6.Visible = False And Picture7.Visible = False And Picture8.Visible = False And Picture9.Visible = False Then
Timer1.Enabled = False
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
Text1.Text = ""
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Label8.Caption = Label8.Caption - "1"
If Label8.Caption = "0" Then
MsgBox "Good job, here are your stats:" + vbNewLine + Label1.Caption + Label5.Caption + vbNewLine + Label2.Caption + Label6.Caption + vbNewLine + Label3.Caption + Label7.Caption + vbNewLine + Label4.Caption + Label8.Caption + vbNewLine + "Please start a new game now", vbOKOnly, "Stats."
Text1.Text = ""
Text2.Text = ""
Label4.Caption = "Time:"
beginner.Checked = False
intermediate.Checked = False
expert.Checked = False
custom.Checked = False
beginner.Enabled = True
intermediate.Enabled = True
expert.Enabled = True
custom.Enabled = True
Timer1.Enabled = False
End If
End Sub
