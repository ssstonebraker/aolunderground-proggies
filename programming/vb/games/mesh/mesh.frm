VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Tanner's 3D Mesh Creator"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10935
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   729
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".msh"
      DialogTitle     =   "Open/Save Mesh File"
      Filter          =   "Mesh Files | *.msh"
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      DrawWidth       =   4
      Height          =   2055
      Left            =   120
      ScaleHeight     =   35.19
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   39.423
      TabIndex        =   15
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "Reset the Matrix"
      Height          =   1575
      Left            =   4200
      TabIndex        =   14
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Frame FrameStuff 
      Caption         =   "Preset Patterns"
      Height          =   1815
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton Option2 
         Caption         =   "Valley"
         Height          =   315
         Left            =   960
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mountain"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   1200
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton CmdRandomize 
         Caption         =   "Randomize "
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox TxtRandomize 
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Text            =   "5"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRndFractal 
         Caption         =   "Randomize w/fractal algorithm"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox TxtRnd 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Text            =   "3"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TxtLoops 
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Text            =   "15"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton CmdMountain 
         Caption         =   "Create:"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Rnd Value:"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Rnd:"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Loops:"
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H000000FF&
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "COLORFORM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label LblHigh 
      Caption         =   "0"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label LblLow 
      Caption         =   "0"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "AERIAL VIEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "TOP VIEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnusave 
         Caption         =   "&Save mesh file"
      End
      Begin VB.Menu mnuload 
         Caption         =   "&Load mesh file"
      End
      Begin VB.Menu mnusepbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
   End
   Begin VB.Menu mnualgorithms 
      Caption         =   "&Algorithms"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3D Mesh example by Tanner Helland

'Here is an awesome example of how a little bit of ingenuity and a
'little bit of math can create incredible graphic effects.  This
'project uses simulated terraforming to create the effect of a
'3D landscape.  The code is mostly a lot of line drawing, and is
'somewhat complex.  It isn't commented, so if you really want to know
'how it works, e-mail me at tannerhelland@hotmail.com.  Enjoy!

'Feel free to use this code for any personal purposes, but please
'give me some credit and let me know how you use it.

'If you're into game programming, my affiliate company (ART Global) is
'currently working on Realms of Time, a state-of-the-art RPG that
'is going to revolutionize the gaming world.  If you have experience
'in C++ or excellent skills in VB graphic or text programming, or
'if you just want to hear the details of this incredible project,
'contact me at the e-mail address listed above.  Beta testers will
'be needed at some point in the future (NOT NOW, though, so don't
'ask for now!), so if you want to be on the list of candidates,
'contact me with a PROFESSIONAL resume detailing your experience.

Dim CurX, CurY As Integer
Dim Pointarray() As Integer
Const XDist = 1
Const YDist = 0
Const SizeOfArray = 11
Dim Randomness As Single
Dim Loops As Integer
Const RandLimit = 10
Const Draw = 10

Private Sub CmdMountain_Click()

Dim Corner1
Dim Corner2
Dim Corner3
Dim Corner4
Dim Average

Randomness = TxtRnd
Loops = TxtLoops

ReDim Pointarray(0 To SizeOfArray, 0 To SizeOfArray)

For X = 0 To SizeOfArray
Pointarray(X, 0) = Int(Rnd * 5)
Next X
For Y = 0 To SizeOfArray
Pointarray(0, Y) = Int(Rnd * 5)
Next Y

For z = 1 To Loops
For X = 1 To SizeOfArray - 1
For Y = 1 To SizeOfArray - 1
Corner1 = Pointarray(X, Y + 1)
Corner2 = Pointarray(X + 1, Y)
Corner3 = Pointarray(X - 1, Y)
Corner4 = Pointarray(X, Y - 1)
Randomize
If Option1.Value = True Then
Average = Int((Corner1 + Corner2 + Corner3 + Corner4) \ 4) + Int(Rnd * Randomness)
Else
Average = Int((Corner1 + Corner2 + Corner3 + Corner4) \ 4) - Int(Rnd * Randomness)
End If
Pointarray(X, Y) = Average
Next Y
Next X
Next z

Call DrawMesh
End Sub

Private Sub CmdRandomize_Click()
For a = 0 To 10 Step 1
For b = 0 To 10 Step 1
Randomize
Pointarray(a, b) = Int(Rnd * TxtRandomize.Text)
Next b
Next a

Call DrawMesh

End Sub

Private Sub CmdReset_Click()
For a = 0 To 11
For b = 0 To 11
Pointarray(a, b) = 0
Next b
Next a
Call DrawMesh

End Sub

Private Sub CmdRndFractal_Click()

Dim Corner1
Dim Corner2
Dim Corner3
Dim Corner4
Dim Average

Randomness = TxtRnd
Loops = TxtLoops

ReDim Pointarray(0 To SizeOfArray, 0 To SizeOfArray)

For X = 0 To SizeOfArray
Pointarray(X, 0) = Int(Rnd * 5)
Next X
For Y = 0 To SizeOfArray
Pointarray(0, Y) = Int(Rnd * 5)
Next Y

For z = 1 To Loops
For X = 1 To SizeOfArray - 1
For Y = 1 To SizeOfArray - 1
Corner1 = Pointarray(X, Y + 1)
Corner2 = Pointarray(X + 1, Y)
Corner3 = Pointarray(X - 1, Y)
Corner4 = Pointarray(X, Y - 1)
Randomize
Average = Int((Corner1 + Corner2 + Corner3 + Corner4) \ 4) + Int(Rnd * Randomness)
Dim temp As Integer
temp = Int(Rnd * 2)
If temp = 1 Then Average = -Average
Pointarray(X, Y) = Average
Next Y
Next X
Next z

Call DrawMesh

End Sub

Private Sub Form_Load()
ReDim Pointarray(0 To 11, 0 To 11) As Integer
Form1.Show
Call DrawMesh
End Sub

Public Sub DrawMesh()
Dim temp As Integer
temp = (SizeOfArray - 1) * Draw
Picture1.Cls
For a = 0 To temp Step Draw
For b = 0 To temp Step Draw
Picture1.Line (a + Pointarray(a / Draw, b / Draw), b + Pointarray(a / Draw, b / Draw))-(a + Draw + Pointarray(a / Draw + 1, b / Draw), b + Pointarray(a / Draw + 1, b / Draw))
Picture1.Line (a + Pointarray(a / Draw, b / Draw), b + Pointarray(a / Draw, b / Draw))-(a + Pointarray(a / Draw, b / Draw + 1), b + Draw + Pointarray(a / Draw, b / Draw + 1))
Next b
Next a

Dim StartPoint As Integer
Picture2.Cls

For X = 0 To 50 Step 5
For Y = 0 To 50 Step 5
StartPoint = 55 - X + Y
'Picture1.PSet ((x + y + (XDist * y + x)), StartPoint), RGB(0, 0, 0)
Picture2.Line ((X + Y + (XDist * Y + X)), StartPoint - Pointarray(X / 5, Y / 5))-((X + Y + (XDist * Y + X + 10)), 5 - X + 50 - 5 + Y - Pointarray((X / 5) + 1, Y / 5))
Picture2.Line ((X + Y + (XDist * Y + X)), StartPoint - Pointarray(X / 5, Y / 5))-((X + Y + (XDist * Y + X + 10)), 5 - X + 50 + 5 + Y - Pointarray(X / 5, (Y / 5) + 1))
Next Y
Next X

Dim TmpHigh As Integer
Dim TmpLow As Integer
TmpHigh = 0
TmpLow = 0

For X = 0 To 10
For Y = 0 To 10
If Pointarray(X, Y) < TmpLow Then TmpLow = Pointarray(X, Y)
Next Y
Next X

For X = 0 To 10
For Y = 0 To 10
Pointarray(X, Y) = Pointarray(X, Y) + Abs(TmpLow)
Next Y
Next X

For X = 0 To 10
For Y = 0 To 10
If Pointarray(X, Y) > TmpHigh Then TmpHigh = Pointarray(X, Y)
Next Y
Next X
LblLow = TmpLow
LblHigh = TmpHigh

'Color

Dim MagicNum As Double
Dim Color As Integer
Dim TempArray(0 To 10, 0 To 10) As Integer
Picture3.Cls
On Error Resume Next
'TmpHigh = 256 - TmpHigh
MagicNum = 256 / TmpHigh
For X = 0 To 10
For Y = 0 To 10
Color = Int(Pointarray(X, Y) * MagicNum)
Picture3.PSet (X, Y), RGB(Color, Color, Color)
Next Y
Next X

End Sub




Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuload_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Open CommonDialog1.FileName For Binary As #1
Get #1, 1, Pointarray
Close #1
Call DrawMesh
End Sub

Private Sub mnusave_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
On Error Resume Next
Kill CommonDialog1.FileName
Open CommonDialog1.FileName For Binary As #1
Put #1, 1, Pointarray
Close #1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = Int(X / Draw)
CurY = Int(Y / Draw)
On Error Resume Next
If Button = 1 Then
Pointarray(CurX, CurY) = Pointarray(CurX, CurY) + 1
End If
If Button = 2 Then
Pointarray(CurX, CurY) = Pointarray(CurX, CurY) - 1
End If

Call DrawMesh
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = Int(X / Draw)
CurY = Int(Y / Draw)
On Error Resume Next
If Button = 1 Then
Pointarray(CurX, CurY) = Pointarray(CurX, CurY) + 1
Call DrawMesh
End If
If Button = 2 Then
Pointarray(CurX, CurY) = Pointarray(CurX, CurY) - 1
Call DrawMesh
End If

End Sub
