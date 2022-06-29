VERSION 5.00
Begin VB.Form SplashScreen 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13530
   Icon            =   "SplashScreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label statuslbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Me"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   4320
      TabIndex        =   0
      Top             =   4275
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   4995
      Left            =   0
      Picture         =   "SplashScreen.frx":164A
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormOnTop Me              'Calls the sub FormOnTop from the module
Image2.Visible = True     'Makes the SplashScreen Picture Visible Property True So That It Is Viewable
Me.Height = Image2.Height 'Sizes The Form To Fit The Picture
Me.Width = Image2.Width   'Sizes The Form To Fit The Picture
End Sub
Private Sub Image1_Click()
Me.Hide                   'Makes the SplashScreen hidden
FrmMain.Show              'Shows the main form, if this has not been done yet then it loads it
End Sub
Private Sub Image2_Click()
FrmMain.Show              'Shows the main form, if this has not been done yet then it loads it
End Sub
