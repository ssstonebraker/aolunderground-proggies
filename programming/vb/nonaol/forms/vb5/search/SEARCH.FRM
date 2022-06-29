VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Search sample by Matt Hart - vbhelp@matthart.com"
   ClientHeight    =   5940
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   Height          =   6345
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1170
   Width           =   6810
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "C"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6495
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   "Text to search for:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Drive To Search:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' Search Sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This shows how to search for specific text in some files.
' Uses recursive techniques to search through sub-directories.

Private Sub Command1_Click()
    Static bInHere As Boolean
    If bInHere Then bCancel = True: Exit Sub
    bInHere = True
    bCancel = False
    Command1.Default = False
    Command1.Cancel = True
    Dim NumFiles As Integer, i As Integer
    Label2.Caption = "Searching"
    Command1.Caption = "Cancel"
    ReDim Files$(0)
    Call FindText(Text2.Text, "*.*", Files$(), NumFiles, UCase$(Text3.Text))
'    For i = 1 To NumFiles
'        List1.AddItem Files$(i)
'    Next
    Label2.Caption = ""
    Command1.Caption = "Begin"
    bInHere = False
    Command1.Cancel = False
    Command1.Default = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_Click()
    Text1.Text = List1.List(List1.ListIndex)
End Sub


