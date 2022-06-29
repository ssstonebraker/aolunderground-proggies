VERSION 5.00
Object = "{72134BA9-52CA-11D2-A11E-24AE06C10000}#2.0#0"; "CHATOCX2.OCX"
Object = "{655D25A2-69EB-11D2-A9C1-D94536B35B75}#2.0#0"; "SCOREKEEPER4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Outer Limits Scrambler"
   ClientHeight    =   3600
   ClientLeft      =   3375
   ClientTop       =   3915
   ClientWidth     =   5325
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":000C
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   360
      Width           =   5055
      Begin VB.Frame Frame6 
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   120
         Width           =   975
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Scramble"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   480
         Width           =   975
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hint"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame13 
         Height          =   615
         Left            =   3960
         TabIndex        =   23
         Top             =   840
         Width           =   975
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Use Lists"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "No Lists"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame12 
         Height          =   375
         Left            =   3960
         TabIndex        =   21
         Top             =   1440
         Width           =   975
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Load List"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame14 
         Height          =   375
         Left            =   3960
         TabIndex        =   28
         Top             =   1800
         Width           =   975
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Leaderboard"
         Height          =   2055
         Left            =   2040
         TabIndex        =   26
         Top             =   960
         Width           =   1815
         Begin VB.Frame Frame16 
            Height          =   375
            Left            =   960
            TabIndex        =   32
            Top             =   1560
            Width           =   735
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   33
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.Frame Frame15 
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   1560
            Width           =   735
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Scroll"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   31
               Top             =   120
               Width           =   735
            End
         End
         Begin ScoreKeeper.Score Score1 
            Height          =   1335
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   2355
         End
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3720
         Top             =   2880
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   3960
         Top             =   2760
      End
      Begin VB.Frame Frame8 
         Height          =   375
         Left            =   90
         TabIndex        =   18
         Top             =   645
         Width           =   0
      End
      Begin VB.Frame Frame4 
         Caption         =   "Leaderboard"
         Height          =   2055
         Left            =   90
         TabIndex        =   11
         Top             =   1050
         Width           =   0
         Begin VB.Frame Frame11 
            Height          =   375
            Left            =   960
            TabIndex        =   20
            Top             =   1560
            Width           =   735
         End
         Begin VB.Frame Frame10 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Category"
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1815
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Form1.frx":3E9F6
            Left            =   120
            List            =   "Form1.frx":3EA1B
            TabIndex        =   8
            Text            =   "?"
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Word To Scramble"
         Height          =   735
         Left            =   2040
         TabIndex        =   5
         Top             =   120
         Width           =   1815
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Text            =   "Outer Limits Ownz"
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   30000
         Left            =   3720
         Top             =   3000
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB6Chat2.Chat Chat1 
         Left            =   3480
         Top             =   2880
         _ExtentX        =   3969
         _ExtentY        =   2170
      End
      Begin VB.Frame Frame5 
         Caption         =   "Scrambler List"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
         Begin VB.ListBox List1 
            Enabled         =   0   'False
            Height          =   840
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Points"
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1815
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Form1.frx":3EA8E
            Left            =   120
            List            =   "Form1.frx":3EAB3
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Outer Limits Scrambler"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If LCase$(What_Said) = LCase$(Text1.Text) Then
Score1.AddNameAndScore Screen_Name, Combo2.Text
If Combo2.Text = "1" Then
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Congrats " + Screen_Name + ", you earned " + Combo2.Text + " point!", False)
Else
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Congrats " + Screen_Name + ", you earned " + Combo2.Text + " points!", False)
End If
Chat1.ChatSend FadedText$
Chat1.ScanOff
Timer1.Enabled = False
Timer3.Enabled = False
Label6.Enabled = False
Label9.Caption = "Scramble"
If Label11.FontBold = True And List1.ListCount > 0 Then
Label9_Click
End If
End If
End Sub

Private Sub Combo1_Change()
If Combo1.Text = "" Then Combo1.Text = "?"
End Sub
Private Sub Combo2_GotFocus()
If Combo2.Text = "" Then Combo2.Text = "1"
End Sub
Private Sub Form_Load()
'Score1.AddNameAndScore "(V)agic", 15
'Score1.AddNameAndScore "CooCat", 8
'Score1.AddNameAndScore "TripleKorn", 4
'Score1.AddNameAndScore "DrumBug1", 3
'Score1.AddNameAndScore "BSaDC", 1
FormOnTop Me
Combo2.Text = 1
'Chat1.ChatSend ""
'timeout (0.1)
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Outer Limits Scrambler", False)
Chat1.ChatSend FadedText$
timeout (0.1)
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Created By (V)agic", False)
Chat1.ChatSend FadedText$
timeout (0.1)
If LCase$(GetUser) = LCase$("Caveman83") Then
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Loaded By: (V)agic", False)
Else
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Loaded By: " + GetUser, False)
End If
Chat1.ChatSend FadedText$
'timeout (0.3)
'Chat1.ChatSend ""
End Sub

Private Sub Label1_Click()
Unload Me
End
End Sub
Private Sub Label10_Click()
Call SndClick
CommonDialog1.FileName = ""
CommonDialog1.FLAGS = cdlOFNFileMustExist
CommonDialog1.DialogTitle = "Load Scrambler List"
CommonDialog1.Filter = "Scrambler Lists (*.scr)|*.scr"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
List1.Clear
Call Loadlistbox(CommonDialog1.FileName, List1)
End If

End Sub

Private Sub Label11_Click()
Call SndClick
If Label11.FontBold = False Then
Frame5.Enabled = True
List1.Enabled = True
Label11.FontBold = True
Label12.FontBold = False
If List1.ListCount = 0 Then Label10_Click
Frame5.Enabled = True
List1.Enabled = True
Else
Label11.FontBold = False
Label12.FontBold = True
Frame5.Enabled = False
List1.Enabled = False
End If
End Sub

Private Sub Label12_Click()
Call SndClick
If Label12.FontBold = False Then
Label12.FontBold = True
Label11.FontBold = False
Frame5.Enabled = False
List1.Enabled = False
Else
Label12.FontBold = False
Label11.FontBold = True
Frame5.Enabled = True
List1.Enabled = True
If List1.ListCount = 0 Then Label10_Click
End If
End Sub



Private Sub Label13_Click()
Call SndClick
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Outer Limits Scores", False)
Chat1.ChatSend FadedText$
timeout (0.1)
Score1.SortScore
Score1.SendScore
Call FadeLabel(Label13)
End Sub

Private Sub Label14_Click()
Call SndClick
Score1.ClearScores
Call FadeLabel(Label14)
End Sub

Private Sub Label2_Click()
Form1.WindowState = 1
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub
Private Sub Label4_Click()

End Sub
Private Sub Label5_Click()

End Sub
Private Sub Label6_Click()
Call SndClick
Call FadeLabel(Label6)
If Label9.Caption = "Stop" Then
Timer3.Enabled = True
End If
End Sub
Private Sub Label7_Click()

End Sub
Private Sub Label8_Click()
Call SndClick
Form3.Show
Call FadeLabel(Label8)
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If fade = "y" Then
'Call FadeLabel(Label8)
'Else
'End If
'fade = "n"
End Sub


Private Sub Label9_Click()
Call SndClick
Call FadeLabel(Label9)
If Label9.Caption = "Scramble" Then
    Label9.Caption = "Stop"
    
    If List1.Enabled Then
        Text1.Text = List1.List(0)
        List1.RemoveItem (0) 'list1.listindex=0
    End If

    Text1.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    Chat1.ScanOn
    Text3.Text = ScrambleIt(Text1.Text)
    FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Outer Limits Scrambler", False)
    Chat1.ChatSend FadedText$
    timeout (0.1)
    FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Word: " + Text3.Text, False)
    Chat1.ChatSend FadedText$
    timeout (0.1)
    FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Category: " + Combo1.Text, False)
    Chat1.ChatSend FadedText$

    Timer1.Enabled = True
    Label6.Enabled = True
Else
    FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Outer Limits Scrambler", False)
    Chat1.ChatSend FadedText$
    timeout (0.1)
    FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Stopped", False)
    Chat1.ChatSend FadedText$
    
    Chat1.ScanOff
    Timer1.Enabled = False
    Label9.Caption = "Scramble"
    Text1.Enabled = True
    Combo1.Enabled = True
    Combo2.Enabled = True
    Label6.Enabled = False
    Timer3.Enabled = False
'ascii:  ·•·^v^¤   ¤^v^·•·
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fade = "y"
End Sub


Private Sub Timer1_Timer()
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Outer Limits Scrambler", False)
Chat1.ChatSend FadedText$
timeout (0.1)
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Word: " + Text3.Text, False)
Chat1.ChatSend FadedText$
timeout (0.1)
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Category: " + Combo1.Text, False)
Chat1.ChatSend FadedText$
End Sub

Private Sub Timer2_Timer()
Score1.SortScore
If List1.ListCount = 0 Then
List1.Enabled = False
Frame5.Enabled = False
Else
'List1.Enabled = True
'Frame5.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
For I = 1 To Len(Text1.Text)
hehe = Mid(Text1.Text, 1, I)
FadedText$ = FadeThreeColor(8, 8, 107, 166, 166, 200, 8, 8, 107, "·•·^v^¤ Hint: " + hehe + "...", False)
Chat1.ChatSend FadedText$
timeout (2)
Next I
Timer3.Enabled = False
End Sub
