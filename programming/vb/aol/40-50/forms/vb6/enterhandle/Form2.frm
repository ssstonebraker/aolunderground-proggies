VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "enter hande example"
   ClientHeight    =   930
   ClientLeft      =   3375
   ClientTop       =   3150
   ClientWidth     =   2925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   2925
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form1 'unloads the enter handle form.
Unload Me 'unloads the current form
End 'unloads entire programs
End Sub

Private Sub Form_Load()
'SetOnTop Me 'lil32.bas (by zb) call to set form on top.



End Sub
