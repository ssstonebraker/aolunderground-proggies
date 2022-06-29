VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Yahtzee Deluxe"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":030A
   ScaleHeight     =   2415
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgOK 
      Height          =   330
      Left            =   240
      Picture         =   "frmAbout.frx":1FFC
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblWebsite 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visit our website today!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   525
      TabIndex        =   4
      ToolTipText     =   "Click Now"
      Top             =   2040
      Width           =   2805
   End
   Begin VB.Label lblEmail 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: sharmon@microtechcomputers.com"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   397
      TabIndex        =   3
      Top             =   1800
      Width           =   3060
   End
   Begin VB.Label lblWinVer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Windows(tm) 95 / 98 / NT Version"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   2445
   End
   Begin VB.Label lblCopyright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(C)1999 All Rights Reserved"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1770
      TabIndex        =   2
      Top             =   1320
      Width           =   1995
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
  '//Set version text
  lblVersion = "Version " & App.Major & "." & App.Minor & App.Revision

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  '//User pressed the escape key
  If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub imgOK_Click()
  
  '//Unload Form
  Unload Me

End Sub

Private Sub imgOK_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Change picture on mousedown
  imgOK.Picture = frmMain.ImageList1.ListImages(4).Picture

End Sub

Private Sub imgOK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Reset picture on mouseup
  imgOK.Picture = frmMain.ImageList1.ListImages(3).Picture

End Sub

Private Sub lblWebsite_Click()
  
  '//Website is not valid, waiting on domain number changes...sorry!
  '//You can email for now if you'd like:  sharmon@microtechcomputers.com
  
  '//Goto our website
  ShellExecute 0, "Open", "http://www.vsoftusa.com", "", "", vbNormalFocus
  '//Unload form
  Unload Me

End Sub
