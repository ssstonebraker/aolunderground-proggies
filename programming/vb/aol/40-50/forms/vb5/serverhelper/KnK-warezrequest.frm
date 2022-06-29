VERSION 5.00
Begin VB.Form Form17 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IM Answering machine"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3090
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Reason Gone"
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
