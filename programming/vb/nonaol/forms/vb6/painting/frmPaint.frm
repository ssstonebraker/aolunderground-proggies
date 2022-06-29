VERSION 5.00
Begin VB.Form frmpaint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ben's Paint Program"
   ClientHeight    =   5370
   ClientLeft      =   540
   ClientTop       =   780
   ClientWidth     =   6975
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   Icon            =   "FRMPAINT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FRMPAINT.frx":030A
   Palette         =   "FRMPAINT.frx":045C
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5370
   ScaleWidth      =   6975
   Begin VB.ListBox lstTools 
      Height          =   1230
      Left            =   4680
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   5175
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   5115
      ScaleWidth      =   4275
      TabIndex        =   9
      Top             =   120
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   3
      Left            =   4800
      Max             =   25
      Min             =   2
      TabIndex        =   7
      Top             =   720
      Value           =   3
      Width           =   1935
   End
   Begin VB.Timer tmrCursor 
      Interval        =   1
      Left            =   480
      Top             =   5520
   End
   Begin VB.PictureBox pCol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   4680
      MouseIcon       =   "FRMPAINT.frx":2132
      MousePointer    =   99  'Custom
      Picture         =   "FRMPAINT.frx":243C
      ScaleHeight     =   1110
      ScaleWidth      =   2145
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image target 
      Height          =   480
      Left            =   2040
      Picture         =   "FRMPAINT.frx":A15E
      Top             =   5520
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to choose pen/fill color"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image bucket 
      Height          =   480
      Left            =   1320
      Picture         =   "FRMPAINT.frx":A468
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image curpencil 
      Height          =   480
      Left            =   960
      Picture         =   "FRMPAINT.frx":A5BA
      Top             =   5400
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4680
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pen Size"
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblPenSize 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmpaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type pointapi
   x As Double
   y As Double
End Type

Option Explicit

Dim pressed As Boolean     'the variable that i use to tell if the mouse is being held down
Dim colpressed As Boolean  'same as above but for the color chooser
Dim filltool As Boolean    'Variable that tells if the fill tool has been selected
Dim drawtool As Boolean    'variable that tells if the pen tool has been selected
Dim circletool As Boolean  'circle tool variable
Dim whatradius As Variant  'used for circle tool
Dim eyedroptool As Boolean 'var for eye dropper
Dim circgradient As Boolean 'var for circualar gradient
Dim radius As Integer
Dim about
Dim onlynumbers            'to make sure radius is an integer
Dim point1 As pointapi
Dim point2 As pointapi
Dim g1                     'these 3 are for the gradient tool
Dim g2
Dim g3
Dim cgformat
Dim gformat
Dim gdirection
Dim lf1                    'gradient (d)irection,number
Dim lr2
Dim lr3
Dim lr4
Dim ud1
Dim ud2
Dim ud3
Dim ud4
Dim cg1                    'circular gradient variables
Dim cg2
Dim cg3
Dim cg4
Dim index
Dim index2
Dim index3
Dim index4
Dim i As Integer
Dim c As Integer
Private Sub cmdExit_Click()
Unload Me          'unloads the program
End Sub



Private Sub Command1_Click()
picBoard.BackColor = &H80000009  'sets the backcolor of picBoard to white
End Sub


Private Sub Form_Load()
HScroll1.Value = 2
picBoard.DrawWidth = 2
picBoard.MouseIcon = curpencil
pCol.MouseIcon = target
drawtool = True
lstTools.AddItem ("Pen")
lstTools.AddItem ("Circle")
lstTools.AddItem ("Paint Bucket")
lstTools.AddItem ("Eye Dropper")
lstTools.AddItem ("Gradient")
lstTools.AddItem ("Circular Gradient")
lstTools.AddItem ("Clear")
picBoard.ScaleHeight = 255
picBoard.ScaleWidth = 255
End Sub

Private Sub HScroll1_Change()
lblPenSize.Caption = HScroll1.Value          'sets the pen size according
picBoard.DrawWidth = HScroll1.Value          'to the value of the scroll bar
End Sub

Private Sub lstTools_Click()
If lstTools.Text = "Pen" Then
drawtool = True    'tells the computer that the
filltool = False   'pen has been selected
circletool = False
eyedroptool = False
circgradient = False
End If

If lstTools.Text = "Circle" Then

On Error Resume Next
drawtool = False
filltool = False
circletool = True
eyedroptool = False
circgradient = False
GetRadius:
whatradius = InputBox("Enter the radius for the circle in pixels:", "Paint")
If IsNumeric(whatradius) Or radius = "" Then
radius = Val(whatradius)
Else
onlynumbers = MsgBox("You have to enter a number!", vbCritical, "Paint")
GoTo GetRadius
End If

End If
If lstTools.Text = "Paint Bucket" Then
filltool = True                 'tells the computer the
drawtool = False                'fill tool has been selected
circletool = False
eyedroptool = False
circgradient = False
End If
If lstTools.Text = "Clear" Then
picBoard.BackColor = &H80000009
End If
If lstTools.Text = "Eye Dropper" Then
eyedroptool = True
drawtool = False                'tells the computer that the
filltool = False                'eye dropper has been selected
circletool = False
circgradient = False
picBoard.MouseIcon = target
End If
If lstTools.Text = "Gradient" Then
Call gradientmaker              'the function that adds gradients
End If
If lstTools.Text = "Circular Gradient" Then
filltool = False                'tells the computer the
drawtool = False                'circular gradient has been selected
circletool = False
eyedroptool = False
circgradient = True
HScroll1.Value = 11             'makes pen size bigger so gradient is smoother
End If
End Sub

Private Sub mnuAbout_Click()
about = MsgBox("Program by Ben Doherty" & Chr$(10) & Chr$(10) & "jake-d@mindspring.com", vbInformation, "About")
End Sub


Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub pCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
colpressed = True                          'this function tells the computer to
Shape1.FillColor = pCol.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = pCol.Point(x, y)
End Sub


Private Sub pCol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If colpressed Then                         'same as above but this is here
Shape1.FillColor = pCol.Point(x, y)        'so you don't have to keep clicking to
picBoard.ForeColor = pCol.Point(x, y)      'change the color...you can just drag
End If
End Sub

Private Sub pCol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
colpressed = False          'stops the selecting of the color when the user 'unclicks'
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

pressed = True
point1.x = x
point1.y = y
If filltool = True Then                    'if the fill tool is selected it fills picBoard with a custom color
picBoard.BackColor = Shape1.FillColor
End If
If drawtool = True Then                    'draws a point where the user clicks on picBoard
picBoard.Line (x, y)-(x, y)
End If
If circletool = True Then
picBoard.Circle (point1.x, point1.y), radius
End If
If eyedroptool = True Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = picBoard.Point(x, y)
End If


If pressed And circgradient Then
On Error GoTo errhandler
cgformat = InputBox("How do you want to format your gradient? (RGB),  1: ##I, 2: #I#, 3: I##, 4:III (black to white)", "Paint", "1,2,3 or 4")
If cgformat = "1" Then GoTo cg1
If cgformat = "2" Then GoTo cg2
If cgformat = "3" Then GoTo cg3
If cgformat = "4" Then GoTo cg4
cg1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")

 For index = 1 To 400 Step 1

picBoard.Circle (x, y), index, RGB(g1, g2, index)                                           'this is how this works Next i                                                                              'for every "I" the radius of the circle increases
                                                                                    'by 1 and for every "I" the rgb variable also increases by 1
pressed = False                                                                                  'that's what gives it a smooth blended look
                                                                                    'that's how linear gradients work.  except the line distance increases
Exit Sub

cg2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")


For index2 = 1 To 400 Step 1

picBoard.Circle (x, y), index2, RGB(g1, index2, g2)
Next index2

pressed = False
Exit Sub
cg3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ")

For index3 = 1 To 400 Step 2

picBoard.Circle (x, y), index3, RGB(index3, g1, g2)
Next index3
pressed = False
Exit Sub
cg4:
For index4 = 1 To 400 Step 2

picBoard.Circle (x, y), index4, RGB(index4, index4, index4)
Next index4
pressed = False
Exit Sub
errhandler:
    MsgBox ("An error has occured")
    Exit Sub
Next
End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If pressed And drawtool Then
point2 = point1
point1.x = x
point1.y = y
picBoard.Line (point1.x, point1.y)-(point2.x, point2.y)         'if the mouse is dragged...the line is continued
End If
If pressed And eyedroptool Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = picBoard.Point(x, y)
End If
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
pressed = False                    'stops drawing the line
End Sub

Private Sub tmrCursor_Timer()
If drawtool = True Then              'this function sets the cursor for
picBoard.MouseIcon = curpencil       'picBoard according to what tool is
End If                               'selected
If filltool = True Then
picBoard.MouseIcon = bucket
End If
If circletool = True Then
picBoard.MouseIcon = target
End If
End Sub

Private Sub gradientmaker()
 On Error GoTo errhandler
gdirection = InputBox("What direction do you want the gradient to fade?", "Paint", "LEFT-RIGHT or UP-DOWN")
gformat = InputBox("How do you want to format your gradient? (RGB),  1: ##I, 2: #I#, 3: I##, 4:III (black to white)", "Paint", "1,2,3 or 4")
If gdirection = "LEFT-RIGHT" And gformat = "1" Then GoTo lr1
If gdirection = "LEFT-RIGHT" And gformat = "2" Then GoTo lr2
If gdirection = "LEFT-RIGHT" And gformat = "3" Then GoTo lr3
If gdirection = "LEFT-RIGHT" And gformat = "4" Then GoTo lr4

If gdirection = "UP-DOWN" And gformat = "1" Then GoTo ud1
If gdirection = "UP-DOWN" And gformat = "2" Then GoTo ud2
If gdirection = "UP-DOWN" And gformat = "3" Then GoTo ud3
If gdirection = "UP-DOWN" And gformat = "4" Then GoTo ud4
lr1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")

 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(g1, g2, i)
    Next i
Exit Sub
lr2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")


 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(g1, i, g2)
    Next i
Exit Sub
lr3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ")

 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(i, g1, g2)
    Next i
Exit Sub
lr4:
 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(i, i, i)
    Next i
Exit Sub
ud1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")


 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(g1, g2, i)
    Next i
Exit Sub
ud2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")

 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(g1, i, g2)
    Next i
Exit Sub
ud3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")

 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(i, g1, g2)
    Next i
Exit Sub
ud4:
 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(i, i, i)
    Next i
Exit Sub
errhandler:
    MsgBox ("An error has occured")
End Sub
