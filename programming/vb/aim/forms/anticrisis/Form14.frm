VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form14"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form14"
   ScaleHeight     =   3600
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "Form14.frx":0000
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.Frame Frame2 
         BackColor       =   &H80000012&
         Height          =   1455
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   1815
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Clear List"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Add Room"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Width           =   1455
         End
      End
      Begin VB.ListBox List1 
         Height          =   1260
         ItemData        =   "Form14.frx":0015
         Left            =   120
         List            =   "Form14.frx":0017
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Áñ†ï ÇrîSïS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form14.Top <= -5000
Form14.Top = Trim(Str(Int(Form14.Top) - 175))
Loop
Unload Form14
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Form_Load()
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Label2_Click()
Call AddRoom_ToList(List1)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Label3_Click()
List1.Clear
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = "&H00FF00"
End Sub

Private Sub Label4_Click()
Call IM_MassIM(List1, "" + Text2.text + "")
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = "&H00FF00"
End Sub

Private Sub Label5_Click()
If Text1 = "" Then
MsgBox "Add a name dumbass", vbCritical, "Dumbass"
Else
List1.AddItem Text1

End If
Text1 = ""
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = "&H00FF00"
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
End Sub
