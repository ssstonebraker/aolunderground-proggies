VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form9"
   ScaleHeight     =   3000
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Room"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   1065
         ItemData        =   "Form9.frx":0000
         Left            =   240
         List            =   "Form9.frx":0002
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Áñ†ï ÇrîSïS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form9.Top <= -5000
Form9.Top = Trim(str(Int(Form9.Top) - 175))
Loop
Unload Form9
End Sub

Private Sub Form_Load()
Call StayOnTop(Form9.hwnd, True)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
End Sub

Private Sub Label2_Click()
Call AddRoom_ToList(List1)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = "&H00FF00"
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
End Sub
