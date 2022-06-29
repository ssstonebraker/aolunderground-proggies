VERSION 5.00
Begin VB.Form Form23 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wavy Underlined Text"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Wavy Underlined Text"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
G$ = Text1
A = Len(G$)
For w = 1 To A Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<u><s><sup>" & R$ & "</sup>" & u$ & "<sub>" & S$ & "</sub></s></u>" & T$
Next w
Wav = P$
SendChat Wav
End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

