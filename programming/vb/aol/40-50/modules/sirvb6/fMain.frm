VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Form To Load"
   ClientHeight    =   360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Aim"
      Height          =   345
      Left            =   1530
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aol"
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1545
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
fAOL.Show
End Sub

Private Sub Command2_Click()
fAim.Show
End Sub

Private Sub Form_Load()
fMain.Top = 0  '//  Sets this form to the very top of your screen
fMain.Left = (Screen.Width - fMain.Width) \ 2   '//  Sets this form in the middle of your screen, at the top
StayOnTop fMain  '//  Makes this form stay on top of all others
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'//  Sub called from the bas file to remove all opened forms from the memory
Form_UnloadAll
End Sub

Private Sub Form_Resize()
'//  Keeps your window on top even after resized or minimized
StayOnTop fMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
'//  Exits your program
End
End Sub
