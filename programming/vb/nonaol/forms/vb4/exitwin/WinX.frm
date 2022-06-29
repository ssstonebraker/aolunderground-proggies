VERSION 5.00
Begin VB.Form NiVXiT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NiVeK's EXiT WiNDoWs example- NiVeK"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Log Off"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Force Out"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "REBooT WiN"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "SHuT Down WiN"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lable99 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   $"WinX.frx":0000
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "NiVXiT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call ExitWindowsEx(EWX_SHUTDOWN, 0)

End Sub

Private Sub Command2_Click()
Call ExitWindowsEx(EWX_REBOOT, 0)
End Sub

Private Sub Command3_Click()
Call ExitWindowsEx(EWX_FORCE, 0)

End Sub

Private Sub Command4_Click()
Call ExitWindowsEx(EWX_LOGOFF, 0)

End Sub

Private Sub Form_Load()
' please have fun with these codes
' and edit if you please, thanx
' NiVeK

'EWX_FORCE      Closes proggs that do not respond

'EWX_LOGOFF     Logs of current user

'EWX_SHUTDOWN   Shuts Down Windows

'EWX_REBOOT     Reboots Windows

End Sub
