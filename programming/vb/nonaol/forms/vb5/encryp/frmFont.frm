VERSION 5.00
Begin VB.Form frmFont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Font..."
   ClientHeight    =   1428
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   2628
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1428
   ScaleWidth      =   2628
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   350
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Choose the font to encrypt:"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1932
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FontName As String
FontName$ = Combo1.Text
Unload Me
frmEncrypt.Visible = True
frmEncrypt.Text1.Font = FontName$
End Sub

Private Sub Form_Load()
Dim i%
For i = 0 To Screen.FontCount - 1
Combo1.AddItem Screen.Fonts(i)
Next i
Combo1.Text = "Arial"
End Sub

