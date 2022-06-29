VERSION 2.00
Begin Form Mouse 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Project"
   ClientHeight    =   4575
   ClientLeft      =   1245
   ClientTop       =   1680
   ClientWidth     =   6030
   Height          =   5010
   Icon            =   CHP14.FRX:0000
   Left            =   1170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6030
   Top             =   1320
   Width           =   6180
   Begin PictureBox pictColor 
      BackColor       =   &H0000FFFF&
      Height          =   492
      Index           =   4
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   7
      Top             =   3960
      Width           =   492
   End
   Begin PictureBox pictColor 
      BackColor       =   &H00800000&
      Height          =   492
      Index           =   3
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   3960
      Width           =   492
   End
   Begin PictureBox pictColor 
      BackColor       =   &H000000FF&
      Height          =   492
      Index           =   2
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   3960
      Width           =   492
   End
   Begin PictureBox pictColor 
      BackColor       =   &H00000000&
      Height          =   492
      Index           =   1
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   4
      Top             =   3960
      Width           =   492
   End
   Begin PictureBox pictColor 
      Height          =   492
      Index           =   0
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   3960
      Width           =   492
   End
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   492
      Left            =   1200
      TabIndex        =   2
      Top             =   3960
      Width           =   972
   End
   Begin CommandButton cmndClear 
      Caption         =   "&Clear"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   972
   End
   Begin PictureBox pictCanvas 
      Height          =   3732
      Left            =   120
      ScaleHeight     =   3705
      ScaleWidth      =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   5784
   End
   Begin Label lablCurrentColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "·"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Symbol"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   492
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
      Width           =   612
   End
End
Dim DrawOn As Integer              'flags if we're in the middle of drawing something

Dim StartX As Single               'remembers the original start coordinates of
Dim StartY As Single               'a drawing operation

Dim LastX As Single                'remembers the last used drawing coordinates
Dim LastY As Single

Dim SaveColor As Single            'remembers the previous color setting changed in the Click event

Sub cmndClear_Click ()
    pictCanvas.Cls                 'clear the canvas
End Sub

Sub cmndExit_Click ()
    End                            'End the program
End Sub

Sub pictCanvas_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawOn = True                  'flag that we've started drawing
    StartX = X                     'remember our starting coordinates
    StartY = Y
End Sub

Sub pictCanvas_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawOn Then                                     'if we've started drawing,
    If Button And LEFT_BUTTON Then                 'left button pressed,
	pictCanvas.Line (LastX, LastY)-(X, Y)      'just let the user scribble
    ElseIf Button And RIGHT_BUTTON Then            'right button makes cool guy designs...
	If Shift And SHIFT_MASK Then               'shift button draws a
	   pictCanvas.Line (StartX, StartY)-(X, Y), , B   'series of boxes
	End If
	pictCanvas.Line (StartX, StartY)-(X, Y)    'this connects starting point to the current
    End If                                         'point, making a "fan" shape
    End If
    LastX = X                                          'remember where we left off
    LastY = Y
End Sub

Sub pictCanvas_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawOn = False                 'done drawing now!
    StartX = 0                     'reset starting coordinates
    StartY = 0
End Sub

Sub pictColor_Click (Index As Integer)
    SaveColor = pictCanvas.ForeColor                        'remember ForeColor in case of doubleclick
    pictCanvas.ForeColor = pictColor(Index).BackColor       'reset the drawing colors
    lablCurrentColor.ForeColor = pictColor(Index).BackColor 'and give the user visual feedback
End Sub

Sub pictColor_DblClick (Index As Integer)
    pictCanvas.ForeColor = SaveColor                        'reset ForeColor changed during Click event
    lablCurrentColor.ForeColor = SaveColor                  'reset visual feedback...
    pictCanvas.BackColor = pictColor(Index).BackColor       'and now change the background color
    lablCurrentColor.BackColor = pictColor(Index).BackColor 'and the visual feedback
End Sub

