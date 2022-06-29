VERSION 5.00
Begin VB.Form DemoForm 
   BackColor       =   &H00000000&
   Caption         =   "Screen Blanker Demo"
   ClientHeight    =   3855
   ClientLeft      =   960
   ClientTop       =   2535
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   1
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "BLANKER.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   7470
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   120
   End
   Begin VB.CommandButton cmdStartStop 
      BackColor       =   &H00000000&
      Caption         =   "Start Demo"
      Default         =   -1  'True
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1830
   End
   Begin VB.PictureBox picBall 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1800
      Picture         =   "BLANKER.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   8
      Left            =   6330
      Picture         =   "BLANKER.frx":0614
      Top             =   3765
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line linLineCtl 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   240
      X2              =   4080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   7
      Left            =   5760
      Picture         =   "BLANKER.frx":091E
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   6
      Left            =   5160
      Picture         =   "BLANKER.frx":0C28
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   5
      Left            =   4560
      Picture         =   "BLANKER.frx":0F32
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   4
      Left            =   3960
      Picture         =   "BLANKER.frx":123C
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   3
      Left            =   3360
      Picture         =   "BLANKER.frx":1546
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   2
      Left            =   2760
      Picture         =   "BLANKER.frx":1850
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   1
      Left            =   2160
      Picture         =   "BLANKER.frx":1B5A
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   0
      Left            =   1560
      Picture         =   "BLANKER.frx":1E64
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpClone 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H000000FF&
      Height          =   1215
      Index           =   0
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   960
      Top             =   1080
      Width           =   15
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Begin VB.Menu mnuLineCtlDemo 
         Caption         =   "&Jumpy Line"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCtlMoveDemo 
         Caption         =   "Re&bound"
      End
      Begin VB.Menu mnuImageDemo 
         Caption         =   "&Spinning Moon"
      End
      Begin VB.Menu mnuShapeDemo 
         Caption         =   "&Madhouse"
      End
      Begin VB.Menu mnuPSetDemo 
         Caption         =   "&Confetti"
      End
      Begin VB.Menu mnuLineDemo 
         Caption         =   "C&rossfire"
      End
      Begin VB.Menu mnuCircleDemo 
         Caption         =   "Rainbo&w Rug"
      End
      Begin VB.Menu mnuScaleDemo 
         Caption         =   "Co&lor Bars"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "DemoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Declare a variable to track animation frame.
Dim FrameNum
' Declare the X- and Y-coordinate variables to track position.
Dim XPos
Dim YPos
' Declare a variable flag to stop graphic routines in Do Loops.
Dim DoFlag
' Declare a variable to track moving controls.
Dim Motion
' Declare form variables for color.
Dim R
Dim G
Dim B

Private Sub CircleDemo()
    ' Declare local variables.
    Dim Radius
    ' Create random RGB colors.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' Position center of circles in the center of the form.
    XPos = ScaleWidth / 2
    YPos = ScaleHeight / 2
    ' Generate a radius between 0 and almost half the form's height.
    Radius = ((YPos * 0.9) + 1) * Rnd
    ' Draw a circle on the form.
    Circle (XPos, YPos), Radius, RGB(R, G, B)
End Sub

Private Sub cmdStartStop_Click()
' Declare local variables.
Dim UnClone
Dim MakeClone
Dim X1
Dim Y1
    Select Case DoFlag
        Case True
            cmdStartStop.Caption = "Start Demo"
            DoFlag = False
            mnuOption.Enabled = True
            If mnuCtlMoveDemo.Checked = True Then
                ' Hide bouncing graphic again.
                picBall.Visible = False
            ElseIf mnuLineDemo.Checked = True Then
                ' Remove lines from the form.
                Cls
            ElseIf mnuShapeDemo.Checked = True Then
                ' Remove all dynamically loaded Shape controls.
                For UnClone = 1 To 20
                    Unload shpClone(UnClone)
                Next UnClone
                ' Reset background color of form to black.
                DemoForm.BackColor = QBColor(0)
                ' Refresh form so color change takes effect.
                Refresh
            ElseIf mnuPSetDemo.Checked = True Then
                ' Remove confetti bits from form.
                Cls
            ElseIf mnuLineCtlDemo.Checked = True Then
                ' Hide Line control again.
                linLineCtl.Visible = False
                ' Remove any stray pixels left after hiding line.
                Cls
            ElseIf mnuImageDemo.Checked = True Then
                ' Hide bouncing graphic again.
                imgMoon(0).Visible = False
            ElseIf mnuScaleDemo.Checked = True Then
                ' Clear the form.
                Cls
                ' Return form to the default scale.
                Scale
            ElseIf mnuCircleDemo.Checked = True Then
                ' Remove the circles from the form.
                Cls
            End If
        Case False
            cmdStartStop.Caption = "Stop Demo"
            DoFlag = True
            mnuOption.Enabled = False
            If mnuCtlMoveDemo.Checked = True Then
                ' Make the bouncing graphic (picture box control) visible.
                picBall.Visible = True
                ' Determine initial motion of bouncing graphic at random.
                ' Settings are 1 to 4.  The value of the Motion variable determines
                ' what part of the Do Loop routine runs.
                Motion = Int(4 * Rnd + 1)
            ElseIf mnuLineDemo.Checked = True Then
                ' Initialize the random-number generator.
                Randomize
                ' Set the line width.
                DrawWidth = 2
                ' Set the initial X- and Y-coordinates to a random location on the form.
                X1 = Int(DemoForm.Width * Rnd + 1)
                Y1 = Int(DemoForm.Height * Rnd + 1)
            ElseIf mnuShapeDemo.Checked = True Then
                ' Dynamically load a control array of 20 shape controls on the form.
                For MakeClone = 1 To 20
                    Load shpClone(MakeClone)
                Next MakeClone
            ElseIf mnuPSetDemo.Checked = True Then
                ' Set the thickness of the confetti bits.
                DrawWidth = 5
            ElseIf mnuLineCtlDemo.Checked = True Then
                ' Make the line control visible.
                linLineCtl.Visible = True
                ' Set thickness of the line as it will appear.
                DrawWidth = 7
            ElseIf mnuImageDemo.Checked = True Then
                ' Make the bouncing graphic (image control) visible.
                imgMoon(0).Visible = True
                ' Set initial animation frame.
                FrameNum = 0
                ' Determine the initial motion of the bouncing graphic at random.
                ' Settings are 1 to 4.  The Value of the Motion variable determines
                ' what part of the Do Loop routine runs.
                Motion = Int(4 * Rnd + 1)
            ElseIf mnuScaleDemo.Checked = True Then
                ' Initialize the random-number generator.
                Randomize
                ' Set the width of the box outlines so boxes don't overlap.
                DrawWidth = 1
                ' Set the value of the X-coordinate to the left edge of form.
                ' Set the first box's X-coordinate = 1, second box = 2, and so on.
                ScaleLeft = 1
                ' Set the Y-coordinate of top edge of form to 10.
                ScaleTop = 10
                ' Set the number of units of the form width to a random number between
                ' 3 and 12.  This changes the number of boxes drawn each time the
                ' routine starts.
                ScaleWidth = Int(13 * Rnd + 3)
                ' Set the number of units in the form height to -10.  Then the height of all boxes
                ' varies from 0 to 10, and Y-coordinates start at the bottom of the form.
                ScaleHeight = -10
            ElseIf mnuCircleDemo.Checked = True Then
                ' Define the width of the circle outline.
                DrawWidth = 1
                ' Draw circles as dashed lines.
                DrawStyle = vbDash
                ' Draw lines using the XOR pen, combining colors found in the pen or
                ' in the display, but not in both.
                DrawMode = vbXorPen
            End If
    End Select
End Sub

Private Sub CtlMoveDemo()
    Select Case Motion
    Case 1
        ' Move the graphic left and up by 20 twips using the Move method.
        picBall.Move picBall.Left - 20, picBall.TOP - 20
        ' If the graphic reaches the left edge of the form, move it to the right and up.
        If picBall.Left <= 0 Then
            Motion = 2
        ' If the graphic reaches the top edge of the form, move it to the left and down.
        ElseIf picBall.TOP <= 0 Then
            Motion = 4
        End If
    Case 2
        ' Move the graphic right and up by 20 twips.
        picBall.Move picBall.Left + 20, picBall.TOP - 20
        ' If the graphic reaches the right edge of the form, move it to the left and up.
        ' Routine determines the right edge of the form by subtracting the graphic
        ' width from the form width.
        If picBall.Left >= (DemoForm.Width - picBall.Width) Then
            Motion = 1
        ' If the graphic reaches the top edge of the form, move it to the right and down.
        ElseIf picBall.TOP <= 0 Then
            Motion = 3
        End If
    Case 3
        ' Move the graphic right and down by 20 twips.
        picBall.Move picBall.Left + 20, picBall.TOP + 20
        ' If the graphic reaches the right edge of the form, move it to the left and down.
        If picBall.Left >= (DemoForm.Width - picBall.Width) Then
            Motion = 4
        ' If the graphic reaches the bottom edge of the form, move it to the right and up.
        ' Routine determines the bottom of the form by subtracting
        ' the graphic height from the form height less 680 twips for the height
        ' of title bar and menu bar.
        ElseIf picBall.TOP >= (DemoForm.Height - picBall.Height) - 680 Then
            Motion = 2
        End If
    Case 4
        ' Move the graphic left and down by 20 twips.
        picBall.Move picBall.Left - 20, picBall.TOP + 20
        ' If the graphic reaches the left edge of the form, move it to the right and down.
        If picBall.Left <= 0 Then
            Motion = 3
        ' If the graphic reaches the bottom edge of the form, move it to the left and up.
        ElseIf picBall.TOP >= (DemoForm.Height - picBall.Height) - 680 Then
            Motion = 1
        End If
    End Select
End Sub

Private Sub Delay()
    Dim Start
    Dim Check
    Start = Timer
    Do Until Check >= Start + 0.15
        Check = Timer
    Loop
End Sub

Private Sub Form_Load()
    DoFlag = False
End Sub

Private Sub Form_Resize()
    If mnuScaleDemo.Checked = True And DemoForm.WindowState = 0 Then
        ' Initialize the random-number generator.
        Randomize
        ' Set the width of the box outlines to narrow so the boxes don't overlap.
        DrawWidth = 1
        ' Set the value of the X-coordinate of the left edge of the form to 1.
        ' This makes it easy to set the position for each box.  The first box has
        ' an X-coordinate of 1, the second has an X-coordinate of 2, and so on.
        ScaleLeft = 1
        ' Set the value of the Y-coordinate of the top edge of the form to 10.
        ScaleTop = 10
        ' Set the number of units in the width of the form to a random number between
        ' 3 and 12.  This changes the number of boxes that are drawn each time the user
        ' starts this routine.
        ScaleWidth = Int(13 * Rnd + 3)
        ' Set the number of units in the height of the form to -10.  This has
        ' two effects.  First, all the boxes then have a height that varies from 0 to 10.
        ' Second, the negative value causes the Y-coordinates to begin at the bottom
        ' edge of the form instead of at the top.
        ScaleHeight = -10
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub ImageDemo()
    Select Case Motion
    Case 1
        ' Move the graphic to the left and up by 100 twips using the Move method.
        imgMoon(0).Move imgMoon(0).Left - 100, imgMoon(0).TOP - 100
        ' Increment animation to next frame.
        IncrFrame
        ' If the graphic reaches the left edge of the form, move right and up.
        If imgMoon(0).Left <= 0 Then
            Motion = 2
        ' If the graphic reaches the top edge of the form, move left and down.
        ElseIf imgMoon(0).TOP <= 0 Then
            Motion = 4
        End If
    Case 2
        ' Move the graphic right and up by 100 twips.
        imgMoon(0).Move imgMoon(0).Left + 100, imgMoon(0).TOP - 100
        ' Increment animation to next frame.
        IncrFrame
        ' If the graphic reaches the right edge of the form, move left and up.
        ' Routine determines the right edge of the form by subtracting
        ' the graphic width from the control width.
        If imgMoon(0).Left >= (DemoForm.Width - imgMoon(0).Width) Then
            Motion = 1
        ' If the graphic reaches the top edge of the form, move right and down.
        ElseIf imgMoon(0).TOP <= 0 Then
            Motion = 3
        End If
    Case 3
        ' Move the graphic right and down by 100 twips.
        imgMoon(0).Move imgMoon(0).Left + 100, imgMoon(0).TOP + 100
        ' Increment animation to next frame.
        IncrFrame
        ' If the graphic reaches the right edge of the form, move left and down.
        If imgMoon(0).Left >= (DemoForm.Width - imgMoon(0).Width) Then
            Motion = 4
        ' If the graphic reaches bottom edge of form, move right and up.
        ' Routine determines the bottom edge of the form by subtracting the graphic
        ' height from the form height minus 680 twips for the height of the title
        ' bar and menu bar.
        ElseIf imgMoon(0).TOP >= (DemoForm.Height - imgMoon(0).Height) - 680 Then
            Motion = 2
        End If
    Case 4
        ' Move the graphic left and down by 100 twips.
        imgMoon(0).Move imgMoon(0).Left - 100, imgMoon(0).TOP + 100
        ' Increment animation to next frame.
        IncrFrame
        ' If the graphic reaches the left edge of the form, move right and down.
        If imgMoon(0).Left <= 0 Then
            Motion = 3
        ' If the graphic reaches the bottom edge of the form, move left and up.
        ElseIf imgMoon(0).TOP >= (DemoForm.Height - imgMoon(0).Height) - 680 Then
            Motion = 1
        End If
    End Select
End Sub

Private Sub IncrFrame()
    ' Increment frame number.
    FrameNum = FrameNum + 1
    ' Control array with animation frames has elements 0 to 7. At the eighth
    ' frame, reset the frame number to 0 for an endless animation loop.
    If FrameNum > 8 Then
        FrameNum = 1
    End If
    ' Set the Picture property of the image control to the Picture property of the current frame.
    imgMoon(0).Picture = imgMoon(FrameNum).Picture
    ' Pause display so animation isn't too fast.
    Me.Refresh
    Delay
End Sub

Private Sub LineCtlDemo()
    ' Set X- and Y-coordinates (left/right position) of the line's start position to a
    ' random location on the form.
    linLineCtl.X1 = Int(DemoForm.Width * Rnd)
    linLineCtl.Y1 = Int(DemoForm.Height * Rnd)
    ' Set X- and Y-coordinates (left/right position) of line's end position to
    ' a random location on the form.
    linLineCtl.X2 = Int(DemoForm.Width * Rnd)
    linLineCtl.Y2 = Int(DemoForm.Height * Rnd)
    ' Clear the form to remove any stray pixels.
    Cls
    ' Pause display before moving the line again.
    Delay
End Sub

Private Sub LineDemo()
    ' Declare local variables.
    Dim X2
    Dim Y2
    ' Create random RGB colors.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' Set the end point of the line control to a random location on the form.
    X2 = Int(DemoForm.Width * Rnd + 1)
    Y2 = Int(DemoForm.Height * Rnd + 1)
    ' Using the Line method, draw from current coordinates to current end
    ' point, giving line a random color. Each line starts where the last
    ' line ends.
    Line -(X2, Y2), RGB(R, G, B)
End Sub

Private Sub mnuCircleDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = True
End Sub

Private Sub mnuCtlMoveDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = True
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuImageDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = True
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuLineCtlDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = True
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuLineDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = True
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuPSetDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = True
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuScaleDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = False
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = True
    mnuCircleDemo.Checked = False
End Sub

Private Sub mnuShapeDemo_Click()
    Cls
    mnuCtlMoveDemo.Checked = False
    mnuLineDemo.Checked = False
    mnuShapeDemo.Checked = True
    mnuPSetDemo.Checked = False
    mnuLineCtlDemo.Checked = False
    mnuImageDemo.Checked = False
    mnuScaleDemo.Checked = False
    mnuCircleDemo.Checked = False
End Sub

Private Sub PSetDemo()
    ' Create random RGB colors.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' XPos sets the horizontal position of a confetti bit to a random location on the form.
    XPos = Rnd * ScaleWidth
    ' YPos sets the vertical position of a confetti bit to a random location on the form.
    YPos = Rnd * ScaleHeight
    ' Draw a confetti bit at XPos, YPos. Assign the confetti bit a random color.
    PSet (XPos, YPos), RGB(R, G, B)
End Sub

Private Sub ScaleDemo()
    ' Declare local variables.
    Dim Box
    ' Creates the same number of boxes as units in the width of the form.
    For Box = 1 To ScaleWidth
        ' Create random RGB colors.
        R = 255 * Rnd
        G = 255 * Rnd
        B = 255 * Rnd
        ' Draw boxes using te Line method with the B (box) F (filled) options.
        ' Boxes start at each X-coordinate determined by ScaleWidth and at
        ' a Y-coordinate of 0 (bottom of form). Each box is 1 unit wide and
        ' has a random height between 0 and 10. Fill the box with a random color.
        Line (Box, 0)-Step(1, (Int(11 * Rnd))), RGB(R, G, B), BF
    Next Box
    ' Pause to display all boxes before redraw.
    Delay
End Sub

Private Sub ShapeDemo()
    ' Declare local variables.
    Dim CloneID
    ' Create random RGB colors.
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    ' Set the form's background color to a random value.
    DemoForm.BackColor = RGB(R, G, B)
    ' Select a random shape control in the control array.
    CloneID = Int(20 * Rnd + 1)
    ' XPos and YPos set position of selected shape control to a random
    ' location on the form.
    XPos = Int(DemoForm.Width * Rnd + 1)
    YPos = Int(DemoForm.Height * Rnd + 1)
    ' Set the shape of the selected shape control to a random shape.
    shpClone(CloneID).Shape = Int(6 * Rnd)
    ' Set the height and width of a selected shape control to a random size between
    ' 500 and 2500 twips.
    shpClone(CloneID).Height = Int(2501 * Rnd + 500)
    shpClone(CloneID).Width = Int(2501 * Rnd + 500)
    ' Set the background color and DrawMode property of the shape control to a random color.
    shpClone(CloneID).BackColor = QBColor(Int(15 * Rnd))
    shpClone(CloneID).DrawMode = Int(16 * Rnd + 1)
    ' Move the selected shape control to XPos, YPos.
    shpClone(CloneID).Move XPos, YPos
    ' Make the selected shape control visible.
    shpClone(CloneID).Visible = True
    ' Wait briefly before selecting and changing the next shape control.
    Delay
End Sub

Private Sub Timer1_Timer()
    If mnuCtlMoveDemo.Checked And DoFlag = True Then
        CtlMoveDemo
    ElseIf mnuLineDemo.Checked And DoFlag = True Then
        LineDemo
    ElseIf mnuShapeDemo.Checked And DoFlag = True Then
        ShapeDemo
    ElseIf mnuPSetDemo.Checked And DoFlag = True Then
        PSetDemo
    ElseIf mnuLineCtlDemo.Checked And DoFlag = True Then
        LineCtlDemo
    ElseIf mnuImageDemo.Checked And DoFlag = True Then
        ImageDemo
    ElseIf mnuScaleDemo.Checked And DoFlag = True Then
        ScaleDemo
    ElseIf mnuCircleDemo.Checked And DoFlag = True Then
        CircleDemo
    End If
End Sub

