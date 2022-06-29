VERSION 5.00
Begin VB.Form frmports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Setting"
   ClientHeight    =   1965
   ClientLeft      =   7830
   ClientTop       =   465
   ClientWidth     =   2730
   Icon            =   "hostports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2730
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton frmports 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "4000"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"hostports.frx":0442
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port for This chat:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1260
   End
End
Attribute VB_Name = "frmports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub cmdConnect_Click()
'Someone clicked the Connect button to connect to someone acting as a server.

End Sub

Private Sub cmdHost_Click()

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'OK was clicked.
MsgBox "Please be patient while we register this chat...", 12, "Please wait.."
'Redefine the settings.
glPort = txtPort.text
Call WriteToINI("lastport", "port", txtPort.text, "c:\windows\happychat.ini")
Form1.Timer5.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Hell
txtPort = GetFromINI("lastport", "port", "c:\windows\happychat.ini")
Hell:
Exit Sub
txtPort.text = glPort
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

