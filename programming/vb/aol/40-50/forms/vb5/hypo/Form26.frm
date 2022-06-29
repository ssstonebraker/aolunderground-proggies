VERSION 5.00
Begin VB.Form Form26 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Name Decoder"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "IiLlOo0"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decode"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2 = LCase(Text1)
Text2.Visible = True
Tea$ = "HyPO ver.¹·º -=Name Verify=- "

A = Len(Tea$)
For w = 1 To A Step 4
    PimPimp5$ = Mid$(Tea$, w, 1)
    Pimp2$ = Mid$(Tea$, w + 1, 1)
    Pimp3$ = Mid$(Tea$, w + 2, 1)
    Pimp4$ = Mid$(Tea$, w + 3, 1)
    Pimp5$ = Pimp5$ & "<FONT COLOR=" & Chr$(34) & "#FFCL25" & Chr$(34) & "><sup><b>" & PimPimp5$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & Pimp2$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & Pimp3$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#FFD700" & Chr$(34) & ">" & Pimp4$
Next w
WavYChaTRBbPiMp = Pimp5$
SendChat WavYChaTRBbPiMp
TimeOut 1
 

SendChat "DeCoMpIlIng Name -=_.·´¯)••÷"
TimeOut 1
ToaST$ = Text1 + "'s real Screen Name is " + Text2

A = Len(ToaST$)
For w = 1 To A Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><b>" & ToaSTR$ & "</b>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & ">" & ToaSTs$ & "" & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$
'···÷••(¯`·._   _.·´¯)••÷
SendChat WavYChaTRBb

End Sub

Private Sub Command2_Click()
Form26.Hide
Unload Form26
End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

