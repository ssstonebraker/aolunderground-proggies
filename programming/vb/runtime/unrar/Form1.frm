VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RAR"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Extract"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim iResult As Integer

    iResult = RARExtract("poo.rar", "", "leigh")
    
    Select Case iResult
    
        Case -1
            MsgBox "Failed to open RAR archive !"
        Case 0
            MsgBox "Nothing to extract or failed !"
        Case Is > 0
            MsgBox iResult & " file(s) extracted !"
        
    End Select
    
End Sub


