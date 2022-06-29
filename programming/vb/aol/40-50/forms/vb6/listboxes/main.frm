VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
'clears list1 so it can be updated and
'wont over count screen names
List1.Clear
Call AddRoomToListbox(List1, False)
'checks the list count so it will know how many are in it
If List1.ListCount = List2.ListCount Then
Else
'clears the listbox
List2.Clear
'adds people from chat room (lower case)
Call AddRoomToListbox(List2, False)
End If
Label1.Caption = "There are " & (List2.ListCount) & " person(s) in " & LCase(RoomCaption)
End Sub
