VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1665
   ClientLeft      =   2265
   ClientTop       =   2250
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"KnK-version.frx":0000
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Unload Me
End Sub
