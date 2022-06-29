VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mass Imer"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Add Room To List"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Send Mass Im"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "What To Say?"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear List"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Users"
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Person"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "Sn To Add"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()

End Sub

Private Sub Command1_Click()
List1.AddItem Text1
End Sub

Private Sub Command2_Click()
List1.Clear
End Sub

Private Sub Command3_Click()
Call AddRoom_ToList(List1)
End Sub

Private Sub Command5_Click()
Call MassIM(List1, Text3)
End Sub

Private Sub Form_Load()
Form1.WindowState = 0
Call AddRoom_ToList(List1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End 'Ternimate
End Sub

Private Sub Label1_Click()
MsgBox "Go to www.angelfire.com/on2/flyman5/cgi.htm,For Cgi Scripts!"
End Sub
