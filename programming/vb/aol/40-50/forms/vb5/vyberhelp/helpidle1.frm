VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "/msg"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   ForeColor       =   &H80000004&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ok"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "messages"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   6135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
Form2.Hide ' hides this form
Form1.Show 'shows form1
End Sub

