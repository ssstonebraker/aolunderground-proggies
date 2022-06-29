VERSION 4.00
Begin VB.Form frmCool 
   Appearance      =   0  'Flat
   Caption         =   "Immitation CoolBar Test Stub"
   ClientHeight    =   5775
   ClientLeft      =   2520
   ClientTop       =   2625
   ClientWidth     =   7440
   ForeColor       =   &H80000008&
   Height          =   6465
   Left            =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   7440
   Top             =   1995
   Width           =   7560
   Begin VB.Frame Frame9 
      Caption         =   "Colours"
      Height          =   855
      Left            =   1560
      TabIndex        =   38
      Top             =   2640
      Width           =   1335
      Begin VB.OptionButton Option8 
         Caption         =   "Greys"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option8 
         Caption         =   "System"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Show Handle"
      Height          =   855
      Left            =   6120
      TabIndex        =   35
      Top             =   2160
      Width           =   1215
      Begin VB.OptionButton Option7 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Separator"
      Height          =   855
      Left            =   6120
      TabIndex        =   27
      Top             =   1200
      Width           =   1215
      Begin VB.OptionButton Option6 
         Caption         =   "Short"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Long"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Tool Tips"
      Height          =   855
      Left            =   4800
      TabIndex        =   24
      Top             =   2160
      Width           =   1215
      Begin VB.OptionButton Option5 
         Caption         =   "On"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Off"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Show Text"
      Height          =   855
      Left            =   4800
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
      Begin VB.OptionButton Option4 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Button Up Offsets"
      Height          =   855
      Left            =   3000
      TabIndex        =   18
      Top             =   1200
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "0 pixels"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "-1 pixel"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Button Down Offsets"
      Height          =   855
      Left            =   3000
      TabIndex        =   15
      Top             =   2160
      Width           =   1695
      Begin VB.OptionButton Option3 
         Caption         =   "0 pixels"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "1 pixel"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bar Style"
      Height          =   1335
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
      Begin VB.OptionButton Option1 
         Caption         =   "Flat"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Default"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Raised"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Lowered"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enable Buttons"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
      Begin VB.CheckBox Check1 
         Caption         =   "Print Page"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Print All"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "First"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Prev"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Next"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Last"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Help"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Value           =   1  'Checked
         Width           =   900
      End
   End
   Begin VB.PictureBox picCoolBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   494
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      Begin VB.PictureBox picMasterBitmaps 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   930
         Left            =   4080
         Picture         =   "Fcoolbar.frx":0000
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.PictureBox picSourceBitmaps 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   930
         Left            =   4200
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.PictureBox picScratchPad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   930
         Left            =   4320
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.Timer timerCoolBar 
         Enabled         =   0   'False
         Interval        =   55
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Click the Help Button for more information."
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Width           =   7215
   End
   Begin VB.Label Label2 
      Caption         =   $"Fcoolbar.frx":17F2
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   5280
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   $"Fcoolbar.frx":187D
      Height          =   1455
      Left            =   120
      TabIndex        =   30
      Top             =   3840
      Width           =   7215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmCool"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim cCoolBar As New clsCoolbar

Private Sub Check1_Click(Index As Integer)

    ' enable/disable buttons
    
    Select Case Index
    Case 0:
        Call cCoolBar.EnableButton(1, Check1(Index).Value)
    Case 1:
        Call cCoolBar.EnableButton(2, Check1(Index).Value)
    Case 2:
        Call cCoolBar.EnableButton(3, Check1(Index).Value)
    Case 3:
        Call cCoolBar.EnableButton(4, Check1(Index).Value)
    Case 4:
        Call cCoolBar.EnableButton(5, Check1(Index).Value)
    Case 5:
        Call cCoolBar.EnableButton(6, Check1(Index).Value)
    Case 6:
        Call cCoolBar.EnableButton(7, Check1(Index).Value)
    End Select

End Sub

Private Sub Form_Load()

    ' tell the class about the buttons and create it, most of the other
    ' properties are set the way I want them by default when the class
    ' initialises (including the size of the buttons)
    
    Caption = Caption & " [Version " & cCoolBar.Version & "]"
    
    Call cCoolBar.Create(Me, picCoolBar, picMasterBitmaps, picSourceBitmaps, picScratchPad, timerCoolBar)
    Call cCoolBar.AddButton("Print Page", "Prints the current page")
    Call cCoolBar.AddButton("Print All", "Prints the whole report")
    Call cCoolBar.AddButton("First", "Moves to the first page of the report")
    Call cCoolBar.AddButton("Prev", "Moves to the previous page")
    Call cCoolBar.AddButton("Next", "Moves to the next page")
    Call cCoolBar.AddButton("Last", "Moves to the last page of the report")
    Call cCoolBar.AddButton("Help", "Opens the Help File")
    Call cCoolBar.AddBar(3)
    Call cCoolBar.AddBar(7)
    Call cCoolBar.Refresh

End Sub

Private Sub Form_Paint()

    ' if the system colours have changed then this event will be triggered
    ' the Class will detect that the colours have changed and will redraw
    ' itself in the new colours, this event will also be triggered every
    ' time a portion of it is exposed - such as when a window above it in
    ' the Z-Order is closed, NOTE: the PictureBox_Paint event does not get
    ' triggered because the PictureBox has its AutoRedraw set to True, and
    
    cCoolBar.Refresh
    
End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub Option1_Click(Index As Integer)

    ' show the different styles of bar shading
    
    Select Case Index
    Case 0:
        cCoolBar.BarStyle = 0
    Case 1:
        cCoolBar.BarStyle = 1
    Case 2:
        cCoolBar.BarStyle = 2
    Case 3:
        cCoolBar.BarStyle = 3
    End Select
    
End Sub

Private Sub Option2_Click(Index As Integer)

    ' toggle the bitmap and text offset when the button is in the
    ' raised position
    
    Select Case Index
    Case 0:
        Call cCoolBar.RaisedOffset(0, 0)
    Case 1:
        Call cCoolBar.RaisedOffset(-1, -1)
    End Select
    
End Sub

Private Sub Option3_Click(Index As Integer)

    ' toggle the bitmap and text offset when the button is in the
    ' lowered position
    
    Select Case Index
    Case 0:
        Call cCoolBar.LoweredOffset(0, 0)
    Case 1:
        Call cCoolBar.LoweredOffset(1, 1)
    End Select
    
End Sub

Private Sub Option4_Click(Index As Integer)

    ' switch the button text on/off and change the size of the buttons
    
    Select Case Index
    Case 0:
        cCoolBar.ShowText = True
    Case 1:
        cCoolBar.ShowText = False
    End Select
    
End Sub

Private Sub Option5_Click(Index As Integer)

    ' toggle on/off the tooltips text
    
    Select Case Index
    Case 0:
        cCoolBar.ShowToolTips = True
    Case 1:
        cCoolBar.ShowToolTips = False
    End Select
    
End Sub

Private Sub Option6_Click(Index As Integer)

    Select Case Index
    Case 0:
        cCoolBar.SeparatorStyle = 0
    Case 1:
        cCoolBar.SeparatorStyle = 1
    End Select
    
End Sub

Private Sub Option7_Click(Index As Integer)

    Select Case Index
    Case 0:
        cCoolBar.ShowHandle = True
    Case 1:
        cCoolBar.ShowHandle = False
    End Select
    
End Sub

Private Sub Option8_Click(Index As Integer)

    Select Case Index
    Case 0:
        cCoolBar.UseSystemColours = True
    Case 1:
        cCoolBar.UseSystemColours = False
    End Select
    
End Sub

Private Sub picCoolBar_Click()

    ' this event doesn't get the mouse X and Y co-ords so these will have to be
    ' calculated by the class, the return value of the call to the .ClickEvent
    ' method will be the index number of the button which the mouse is over
    
    Dim iRet As Integer
    
    iRet = cCoolBar.ClickEvent()
    
    Select Case iRet
    Case 1:
        Debug.Print "Print Page Button clicked"
    Case 2:
        Debug.Print "Print All Button clicked"
    Case 3:
        Debug.Print "Rewind Button clicked"
    Case 4:
        Debug.Print "Prev Button clicked"
    Case 5:
        Debug.Print "Next Button clicked"
    Case 6:
        Debug.Print "Fast Forward Button clicked"
    Case 7:
        Debug.Print "Help Button clicked"
        SendKeys "{F1}"
    End Select

End Sub

Private Sub picCoolBar_KeyDown(KeyCode As Integer, Shift As Integer)

    ' it is necessary to catch this event because in order to detect when
    ' the mouse moves off of the CoolBar we call SetCapture() to ensure
    ' that all mouse move events are sent to this PictureBox, keyboard
    ' events also get sent here and not to the control you probably want
    ' them to go to - this will release the capture until the mouse is
    ' moved over the PictureBox again
    
    Call cCoolBar.KeyDownEvent
    
End Sub

Private Sub picCoolBar_KeyUp(KeyCode As Integer, Shift As Integer)
    
    ' see the comments in the _KeyUp() event
    
    Call cCoolBar.KeyUpEvent
    
End Sub

Private Sub picCoolBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' we need to detect when the mouse goes down so that we can draw the
    ' button in the down position
    
    Dim iRet As Integer

    iRet = cCoolBar.MouseDownEvent(Button, Shift, X, Y)

End Sub

Private Sub picCoolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' we need to detect when the mouse moves over the PictureBox control
    ' so that we can draw the correct button in the Up position
    
    Dim iRet As Integer

    iRet = cCoolBar.MouseMoveEvent(Button, Shift, X, Y)

End Sub

Private Sub picCoolBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' again this event is required to detect when to redraw the button in
    ' the Up position
    
    Dim iRet As Integer
    
    iRet = cCoolBar.MouseUpEvent(Button, Shift, X, Y)
    
End Sub

Private Sub picCoolBar_Resize()

    ' when the PictureBox is resized is will probably need to be redrawn
    
    Call cCoolBar.ResizeEvent

End Sub

Private Sub timerCoolBar_Timer()

    ' the timer is used to tell us when to show the tool tip text
    ' some implementations of ToolTips hide the text again after a few
    ' seconds and others dont - this one doesn't but it could be extended
    ' to do so
    
    Call cCoolBar.TimerEvent
    
End Sub

