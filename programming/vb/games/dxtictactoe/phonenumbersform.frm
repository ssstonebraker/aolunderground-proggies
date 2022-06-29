VERSION 5.00
Begin VB.Form phonenumbersform 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PhoneNumbers"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7110
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save and Dial"
      Default         =   -1  'True
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton new_opponent 
      Caption         =   "Add New Opponent"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   45
   End
End
Attribute VB_Name = "phonenumbersform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_click()
Label1.Caption = thenumber(Combo1.ListIndex)
Label2.Visible = False
End Sub
Private Sub Command2_Click()
MainBoard.Enabled = True
Label2.Visible = False
If numberofnumbers >= 1 Then
    SaveSetting App.Title, "Options", "numberofnumbers", numberofnumbers
        For T = 0 To (numberofnumbers - 1)
            SaveSetting App.Title, "phonenumbers", T, thenumber(T)
            SaveSetting App.Title, "phonenames", T, Combo1.List(T)
            SaveSetting App.Title, "Settings", "lastnamecalled", Combo1.Text
        Next T
    SaveSetting App.Title, "Settings", "lastnamecalled", Combo1.Text
    SaveSetting App.Title, "Settings", "lastnamecalled", Label1.Caption
End If
Unload Me
End Sub
Private Sub Form_Deactivate()
If alwaysshow = 0 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
MainBoard.Enabled = False
Dim label1temp As String
For m = 0 To 20
    thenumber(m) = 0
Next m
numberofnumbers = GetSetting(App.Title, "Options", "numberofnumbers", 0)
If numberofnumbers > 0 Then
    If GetSetting(App.Title, "Settings", "lastnamecalled", 0) <> "Select From List" Then
        Label2.Visible = True
        Label2.Caption = "Last Number Called Was " & GetSetting(App.Title, "Settings", "lastnamecalled", 0) & " at "
        Label1.Caption = GetSetting(App.Title, "Settings", "lastnumbercalled", 0)
    End If
Dim T As Integer
    For T = 0 To (numberofnumbers - 1)
        Combo1.AddItem GetSetting(App.Title, "phonenames", T, 0)
        thenumber(T) = GetSetting(App.Title, "phonenumbers", T, 0)
    Next T
Combo1.Text = "Select Number From List"
Else
    Combo1.Enabled = False
    Combo1.Text = "Click Add New Opponent"
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If numberofnumbers >= 1 Then
    SaveSetting App.Title, "Options", "numberofnumbers", numberofnumbers
        For T = 0 To (numberofnumbers - 1)
            SaveSetting App.Title, "phonenumbers", T, thenumber(T)
            SaveSetting App.Title, "phonenames", T, Combo1.List(T)
            SaveSetting App.Title, "Settings", "lastnamecalled", Combo1.Text
        Next T
    SaveSetting App.Title, "Settings", "lastnamecalled", Combo1.Text
    SaveSetting App.Title, "Settings", "lastnumbercalled", Label1.Caption
End If
End Sub
Private Sub new_opponent_Click()
alwaysshow = 1
On Error GoTo error
Dim newnumber As String
Dim newname As String
newnumber = ""
newname = ""
nameenter: newname = InputBox$("Enter Name of Opponent", "Enter Name")
If newname = "" Then
    If MsgBox("Please Enter Your Opponents Name", vbOKCancel, "Enter Name") = vbCancel Then
        Exit Sub
    Else
        GoTo nameenter
    End If
End If
numberenter: newnumber = InputBox$("Enter Number For " & newname, "Enter Number")
If newnumber > "" Then
    Combo1.AddItem newname
    thenumber(numberofnumbers) = newnumber
    numberofnumbers = numberofnumbers + 1
    Combo1.Enabled = True
    Combo1.Text = "Select From List"
Else
    If MsgBox("Please Enter Phone Number to Dial.", vbOKCancel, "Enter Number") = vbCancel Then
        Exit Sub
    End If
    GoTo numberenter
End If
error:         Exit Sub
End Sub
