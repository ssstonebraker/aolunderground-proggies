VERSION 5.00
Begin VB.Form setup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set-Up"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox cominput 
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   210
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2550
      MaxLength       =   6
      TabIndex        =   1
      Top             =   600
      Width           =   765
   End
   Begin VB.ComboBox list1 
      Height          =   315
      ItemData        =   "setup.frx":0000
      Left            =   2550
      List            =   "setup.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "list1"
      ToolTipText     =   "Enter Max Modem Speed"
      Top             =   1020
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Left            =   1620
      TabIndex        =   3
      Top             =   1740
      Width           =   1470
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Com Port Number"
      Height          =   225
      Left            =   870
      TabIndex        =   6
      Top             =   225
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Name"
      Height          =   225
      Left            =   75
      TabIndex        =   5
      Top             =   630
      Width           =   2310
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum Modem Speed"
      Height          =   225
      Left            =   75
      TabIndex        =   4
      Top             =   1110
      Width           =   2310
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "Please Enter Your Name", vbOKOnly
    Exit Sub
ElseIf list1.Text = "" Then
    MsgBox "Please Enter Your Maximum Modem Speed", vbOKOnly
    Exit Sub
ElseIf cominput.Text = "" Then
    MsgBox "Please Select Your Com Port Number", vbOKOnly
    Exit Sub
End If
'inititate settings
profilename = Text1.Text
commnumber = cominput.Text
maxspeed = list1.Text
setupini = False
'saves reg settings
SaveSetting App.Title, "Settings", "setup", setupini
SaveSetting App.Title, "Settings", "profilename", profilename
SaveSetting App.Title, "Settings", "Com Port", commnumber
SaveSetting App.Title, "Settings", "Modem Speed", maxspeed
MainBoard.Show
Unload Me
End Sub
Private Sub Form_Load()
'checks for mainboard open and gets reg settings
If MainBoard.Visible = False Then
    setupini = GetSetting(App.Title, "Settings", "setup", True)
    profilename = GetSetting(App.Title, "Settings", "profilename", "")
    commnumber = GetSetting(App.Title, "Settings", "Com Port", "")
    maxspeed = GetSetting(App.Title, "Settings", "Modem Speed", "")
End If
'checks if needs to run setup
If setupini = True Or profilename = "" Then
    setup.Show
ElseIf MainBoard.Visible = False Then
  Unload Me
  MainBoard.Show
End If
cominput.Text = commnumber
list1.Text = maxspeed
'error checking for entries
If profilename = "" Then
    Label2.Caption = "Please Enter Your Name."
Else
    Label2.Caption = "Your Name"
    Text1.Text = profilename
End If
'populates combo boxes
list1.AddItem "9600"
list1.AddItem "14400"
list1.AddItem "19200"
list1.AddItem "28800"
list1.AddItem "56000"
cominput.AddItem "1"
cominput.AddItem "2"
cominput.AddItem "3"
cominput.AddItem "4"
cominput.AddItem "5"
cominput.AddItem "6"
End Sub
Private Sub list1_Change() 'checks for list fill
If Text1.Text > "" Then
    If cominput.Text > "" Then
        If list1.Text > "" Then
            Command1.Enabled = True
        End If
    End If
End If
End Sub
Private Sub Text1_Change() 'checks for text fill
If Text1.Text > "" Then
    If cominput.Text > "" Then
        If list1.Text > "" Then
            Command1.Enabled = True
        End If
    End If
End If
End Sub
