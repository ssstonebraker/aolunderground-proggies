VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "License Agreement."
   ClientHeight    =   3945
   ClientLeft      =   2355
   ClientTop       =   2340
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox accept 
      Caption         =   " I Accept These Terms."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "I Accept The Terms."
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "Do Not Accept"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   105
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNextTip_Click()
End
End Sub

Private Sub cmdOK_Click()
    If accept.Value = 1 Then
    SaveSetting App.Title, "Options", "license", accept.Value
        Unload Me
        Splash.Show
        Exit Sub
    End If
    MsgBox "Please Check The I Agree Box Below, or press disagree to exit", vbOKOnly
    Exit Sub
End Sub
Private Sub Form_Load()
Label1.Caption = "Welcome To Vegas Video KENO" _
    & vbCrLf & "© 1999 TimeLine Studios Software." & vbCrLf & "All Rights Reserved." _
    & vbCrLf & vbCrLf & "This Program is Freeware." & vbCrLf & "This program is not crippled" _
    & vbCrLf & "or function limited in any way." & vbCrLf & vbCrLf & "By clicking below and pressing accept" & vbCrLf _
    & " You agree to the terms below and understand" & vbCrLf & "that this program offers no warranties express or implied" & vbCrLf & _
    "and will not be held liable for any loss of data." & vbCrLf & vbCrLf & "DO NOT RE-DISTRIBUTE WITHOUT PERMISSION." _
    & vbCrLf & "See about box for more information."
    


Dim acceptlicense As Long
        ' See if we should be shown at startup
    acceptlicense = GetSetting(App.Title, "Options", "license", 0)
    If acceptlicense = 0 Then
    Exit Sub
    Else
        Unload Me
        Splash.Show
        End If
    
End Sub

Private Sub Picture1_Click()

End Sub

