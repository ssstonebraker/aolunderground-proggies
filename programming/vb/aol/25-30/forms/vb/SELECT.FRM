VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Multiple Selections With A List Box"
   ClientHeight    =   5130
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7065
   Height          =   5535
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7065
   Top             =   1140
   Width           =   7185
   Begin PictureBox Picture1 
      Height          =   4740
      Left            =   225
      ScaleHeight     =   4710
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   195
      Width           =   6615
      Begin ListBox lstItems 
         Height          =   3735
         Left            =   285
         MultiSelect     =   1  'Simple
         TabIndex        =   2
         Top             =   330
         Width           =   6195
      End
      Begin CommandButton cmdDelete 
         Caption         =   "&Delete selected items from the list"
         Height          =   450
         Left            =   165
         TabIndex        =   1
         Top             =   4065
         Width           =   6225
      End
   End
End
Option Explicit

Sub cmdDelete_Click ()

    Dim nEntryNumber As Integer

    nEntryNumber = lstItems.ListCount

    Do While nEntryNumber > 0

        nEntryNumber = nEntryNumber - 1
        If lstItems.Selected(nEntryNumber) = True Then lstItems.RemoveItem nEntryNumber

    Loop

End Sub

Sub Form_Load ()

    Do While lstItems.ListCount < 500

        lstItems.AddItem "Item " & lstItems.ListCount
    
    Loop

End Sub

