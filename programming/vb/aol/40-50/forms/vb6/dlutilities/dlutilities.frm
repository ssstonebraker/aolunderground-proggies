VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "koko's download utilities"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   Icon            =   "dlutilities.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Text            =   "dl utilities"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00008000&
      Caption         =   "no"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00008000&
      Caption         =   "yes"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6000
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "stop auto sign on"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MaskColor       =   &H0000C000&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "start auto sign on"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MaskColor       =   &H0000C000&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "start idle"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      MaskColor       =   &H0000C000&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "stop idle"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      MaskColor       =   &H0000C000&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Text            =   "reason here"
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "go to a specific pr. when signed back on?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "continue download when signed back on?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Line Line12 
      X1              =   5640
      X2              =   5640
      Y1              =   1080
      Y2              =   1800
   End
   Begin VB.Line Line11 
      X1              =   2880
      X2              =   5640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line10 
      X1              =   2880
      X2              =   2880
      Y1              =   1200
      Y2              =   1800
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   120
   End
   Begin VB.Line Line8 
      X1              =   5640
      X2              =   5640
      Y1              =   120
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   2880
      X2              =   120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   1800
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   5640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "koko's download utilities "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   2880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   2880
      X2              =   2880
      Y1              =   240
      Y2              =   1200
   End
   Begin VB.Line Line4 
      X1              =   2880
      X2              =   5520
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
AOL40SendChat GetUser
End Sub

Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
AOL40ContinueDownload
End Sub

Private Sub Command3_Click()
Do
AOL40SendChat ("<font face=""Comic Sans MS"">") + ("·•‹ koko's download utilities idle")
TimeOut (60)
Loop
End Sub

Private Sub Command4_Click()
Do
DoEvents
Loop
End Sub

Private Sub Form_Load()
FormAbove Me
AOL40SendChat ("<font face=""Comic Sans MS"">") + ("·•‹ koko's download utilities loaded")
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub
Private Sub Label2_Click()
AOL40SendChat ("<font face=""Comic Sans MS"">") + ("·•‹ koko's download utilities unloaded")
End
End Sub

Private Sub Label5_Click()
WindowState = vbMinimized
End Sub

Private Sub Timer1_Timer()
If GetUser = "" Then
Call AOL40SignOnWithPW(Form3.Text2.Text)
TimeOut (60)
If Option1.Value = True Then
AOL40ContinueDownload
TimeOut (15)
AOL40Keyword ("aol://2719:2-2-") + Form1.Text2
End If
End If
End Sub


