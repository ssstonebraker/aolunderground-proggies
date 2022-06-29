VERSION 5.00
Begin VB.Form Form18 
   BorderStyle     =   0  'None
   Caption         =   "Sn Decoder"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form18.frx":0000
   ScaleHeight     =   2280
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "IloDolanIi"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " _"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sn Decoder"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "  Scroll"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "  Check"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "Form18"
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


End Sub

Private Sub Command2_Click()
Form26.Hide
Unload Form26
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label1_Click()
Text2 = LCase(Text1)
Text2.Visible = True
Tea$ = "Full Moon"

A = Len(Tea$)
For w = 1 To A Step 4
    PimPimp5$ = Mid$(Tea$, w, 1)
    Pimp2$ = Mid$(Tea$, w + 1, 1)
    Pimp3$ = Mid$(Tea$, w + 2, 1)
    Pimp4$ = Mid$(Tea$, w + 3, 1)
    Pimp5$ = Pimp5$ & "<FONT COLOR=" & Chr$(34) & "#FFCL25" & Chr$(34) & "><sup><b>" & PimPimp5$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & Pimp2$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & Pimp3$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#FFD700" & Chr$(34) & ">" & Pimp4$
Next w
WavYChaTRBbPiMp = Pimp5$

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


End Sub

Private Sub Label4_Click()
Unload Form18

End Sub

Private Sub Label5_Click()
Form18.WindowState = 1
End Sub
