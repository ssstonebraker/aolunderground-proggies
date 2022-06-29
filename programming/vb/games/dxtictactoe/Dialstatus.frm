VERSION 5.00
Begin VB.Form Dialstatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialing"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton dialagain 
      Caption         =   "Dial"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Dialstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If answerline = True Then
    MainBoard.Enabled = True
    MainBoard.MSComm1.PortOpen = False
    Unload Me
    Exit Sub
End If
MainBoard.MSComm1.PortOpen = False
phonenumbersform.Enabled = True
instring = ""
Unload Me
alwaysshow = 0
phonenumbersform.Show
End Sub
Private Sub dialagain_Click()
 Call MainBoard.dialme(phonenumbersform.Label1.Caption)
 Label1.Caption = "Dialing... " & phonenumbersform.Label1.Caption
End Sub

Private Sub Form_Load()
MainBoard.Enabled = False
multiplayermode = True
If answerline = True Then
    Me.Caption = "Waiting For a Call."
    dialagain.Visible = False
    Label1.Caption = "Waiting For Ring........"
        If MainBoard.MSComm1.PortOpen = False Then
            MainBoard.MSComm1.CommPort = commnumber
            MainBoard.MSComm1.settings = maxspeed & ",n,8,1"
            MainBoard.MSComm1.InputLen = 0
            MainBoard.MSComm1.EOFEnable = False
            MainBoard.MSComm1.Handshaking = comNone
            MainBoard.MSComm1.InputMode = comInputModeBinary
            MainBoard.MSComm1.RTSEnable = True
            MainBoard.MSComm1.SThreshold = "2"
            MainBoard.MSComm1.RThreshold = "2"
            MainBoard.MSComm1.DTREnable = True
            MainBoard.MSComm1.PortOpen = True
            Debug.Print "Modem Speed is set to " & maxspeed
        End If
    MainBoard.MSComm1.Output = "AT" & vbCr & vbCr & vbCr
    Exit Sub
End If
dialagain.Enabled = False
alwaysshow = 1
Label1.Caption = "Dialing... " & phonenumbersform.Label1.Caption
End Sub

