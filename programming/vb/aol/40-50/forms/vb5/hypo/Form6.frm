VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attention"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Anyone Have HyPO 2.0"
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Attention "
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ToaST$ = "···÷••(¯`·._HyPO AttenTion _.·´¯)••÷"

a = Len(ToaST$)
For w = 1 To a Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "CD2626" & Chr$(34) & "><sup><B>" & ToaSTR$ & "</sup></B>" & "<FONT COLOR=" & Chr$(34) & "#8B1A1A" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#8B3A3A" & Chr$(34) & "><sub>" & ToaSTs$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#4E2F2F" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$
'···÷••(¯`·._   _.·´¯)••÷
SendChat WavYChaTRBb + "{S Goodbye"

SendChat Text1
 
SendChat WavYChaTRBb + "{S IM"
End Sub

Private Sub Command2_Click()
Unload Form6
Form6.Hide

End Sub

Private Sub Form_Load()
StayOnTop Form6
End Sub
