VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scroll Bar"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox scrollBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      ScaleHeight     =   225
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "percent"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   3765
      TabIndex        =   1
      Top             =   645
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   645
      Width           =   195
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   480
      X2              =   480
      Y1              =   600
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   3720
      X2              =   3720
      Y1              =   600
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   480
      X2              =   3720
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------------------------------------------------
' Author:   Brett George
' Email:    shecky@goteddie.com
' Date:     I forget
' Comments: I made this example because i was tired of the
'           uglyness of the default Windows scrollbar. I
'           might release a verticle scroll bar in the near
'           future.
' ---------------------------------------------------------

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long ' I use this sub for getting the X coordinate of the mouse on the screen
Private Type POINTAPI ' Use this to store the mouse's X and Y coordinates
        X As Long
        Y As Long
End Type
Dim cursor As POINTAPI ' You have to dim a variable as the POINTAPI type. You can't just use POINTAPI.
Dim MouseX As Long ' This is to store the mouse's X coordinate, for this program, i do not need a Y coordinate.
Dim ismouseup As Boolean ' Dims the variable that tells if the left mouse button is up
'-----------------------------------
Dim lbl1onx As Integer ' Dims the variable that stores X coordinate of where the user clicks scrollBar
Dim totalp As Long ' Dims the variable that stores the value of the scroll bar

Private Sub Form_Load()
'Start of the initialization
'---------------------------------
ismouseup = True  ' Says that the left mouse button is not being pressed.
totalp& = Line1.X2 - (Line1.X1 + scrollBar.Width)  ' This is part of the initialization of the total value of the scroll bar.
End Sub

Private Sub scrollBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then  ' Only activates if the Left mouse button is held.
    ismouseup = False  ' This says that the left mouse button is being held down.
    lbl1onx = X + 60  ' This keeps track of where the user clicked on (scrollBar).
    scrollBar.BackColor = &H0& ' This just changes the color of the bar when you click it. Added just for effect.
    Call ActivateScroll  ' Activates the main code that operates the whole thing.
End If
End Sub

Private Sub scrollBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ismouseup = True
    scrollBar.BackColor = &H808080 ' Puts the color of the bar to normal
End If
End Sub

Private Sub ActivateScroll()
Dim lblx As Long  ' This is the variable that controls where on (scrollBar) the person clicked.
Do Until ismouseup = True  ' Start the loop of the actual code that moves the scroll bar.
    '=----------------------------------------------------------------=
    ' This is the code that get's the cursor's position on the screen
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ismouseup = False  '  This shows that the user is still holding down the mouse button.
    NoFreeze% = DoEvents  '  Keeps the program from freezing.
    GetCursorPos cursor  '  Get's the cursor's X and Y positions on screen. For this code, we only use the X coordinate.
    MouseX& = cursor.X  ' Puts the cursor's X position into a variable.
    '=----------------------------------------------------------------=
    ' This detects whether the label should be moved or not. The heart
    ' and soul of this example.
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    If MouseX& <= (Form1.Left + Line1.X1 + lbl1onx) / Screen.TwipsPerPixelX Then
        If Not scrollBar.Left = Line1.X1 Then scrollBar.Left = Line1.X1
    ElseIf MouseX& >= ((Form1.Left + Line1.X2 + lbl1onx) - scrollBar.Width) / Screen.TwipsPerPixelX Then
        If Not scrollBar.Left = Line1.X2 - scrollBar.Width Then scrollBar.Left = Line1.X2 - scrollBar.Width
    Else
        lblx& = ((MouseX& * Screen.TwipsPerPixelX) - (Form1.Left + lbl1onx))
        If Not scrollBar.Left = lblx& Then
            scrollBar.Left = lblx& - 10 ' I subtract 10 from the hot spot since i am using Twips for acuracy. If i didn't subtract 10 Twips, it would move over 1 pixel (or 10 Twips) every time you clicked it
        End If
    End If
    '=----------------------------------------------------------------=
    ' This isn't that important.  It just displays the value of the
    ' scroll bar (1 to 100; you can make that higher if you'd like)
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ddcap& = Int(((scrollBar.Left - Line1.X1) / totalp) * 100)
    If Not Label5.Caption = ddcap& Then Label5.Caption = ddcap&
    '=---------------------------------------------------------------=
    ' Start of code for the change of the color of the bar. This
    ' isn't important either. I just threw it in for fun.
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '    NowCol& = 255 * ((scrollBar.Left - Line1.X1) / totalp)
    '    Line1.BorderColor = RGB(255 - NowCol&, 255 - (NowCol& / 2), NowCol&)
        'Line1.BorderColor = RGB(NowCol&, NowCol&, NowCol&)
    '=---------------------------------------------------------------=
Loop
End Sub
