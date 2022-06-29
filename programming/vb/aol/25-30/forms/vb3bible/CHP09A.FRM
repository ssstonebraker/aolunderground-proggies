VERSION 2.00
Begin Form formMain 
   Caption         =   "Coordinates Project"
   ClientHeight    =   3390
   ClientLeft      =   450
   ClientTop       =   1410
   ClientWidth     =   5070
   Height          =   3825
   Icon            =   CHP09A.FRX:0000
   Left            =   375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5070
   Top             =   1050
   Width           =   5220
   Begin PictureBox pictSurface 
      Height          =   2052
      Left            =   240
      ScaleHeight     =   2025
      ScaleWidth      =   4665
      TabIndex        =   10
      Top             =   1200
      Width           =   4692
   End
   Begin ComboBox cmboMeasure 
      Height          =   288
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   1692
   End
   Begin TextBox textMouseY 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3060
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   840
      Width           =   1275
   End
   Begin TextBox textMouseX 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   3060
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   600
      Width           =   1212
   End
   Begin TextBox textCurrentY 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   3060
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   360
      Width           =   1212
   End
   Begin TextBox textCurrentX 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   3060
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Width           =   1212
   End
   Begin CommandButton cmndClear 
      Caption         =   "&Clear"
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
   Begin Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mouse Y"
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mouse X"
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Current Y"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current X"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End

Sub cmboMeasure_Click ()
     If cmboMeasure.Text = "User Defined" Then             'We need to define coordinate system
      formGetUserDefined.Show 1                    'Show modal input box
      pictSurface.ScaleTop = Val(formGetUserDefined!textTop.Text)         'Top
      pictSurface.ScaleLeft = Val(formGetUserDefined!textLeft.Text)       'Left
      pictSurface.ScaleWidth = Val(formGetUserDefined![textWidth].Text)   'Width
      pictSurface.ScaleHeight = Val(formGetUserDefined![textHeight].Text) 'Height
     Else
      pictSurface.ScaleMode = cmboMeasure.ListIndex    'items were added to combo box in
     End If                                                 'order, so index is ScaleMode
     textCurrentX.Text = Format$(pictSurface.CurrentX, "####0.00")     'update display
     textCurrentY.Text = Format$(pictSurface.CurrentY, "####0.00")     'when scaleMode changes
End Sub

Sub cmndClear_Click ()
     pictSurface.Cls                            'Clear the picture
End Sub

Sub Form_Load ()
     formMain.Scale (0, 0)-(100, 100)      'custom coordinate system for form
     cmboMeasure.AddItem "User Defined"    'scaleMode choices into combo box
     cmboMeasure.AddItem "Twips"           'note that we add these in the
     cmboMeasure.AddItem "Points"          'correct order, so that listIndex
     cmboMeasure.AddItem "Pixels"          'indicates correct ScaleMode
     cmboMeasure.AddItem "Characters"
     cmboMeasure.AddItem "Inches"
     cmboMeasure.AddItem "Millimeters"
     cmboMeasure.AddItem "Centimeters"
     cmboMeasure.ListIndex = 1             'set initial ScaleMode to twips
End Sub

Sub Form_Resize ()
     pictSurface.Move 0, 40                             'leave room for controls on top
     pictSurface.Left = formMain.ScaleLeft + 2          'fill rest of form, less a small border
     pictSurface.Width = formMain.ScaleWidth - 4
     pictSurface.Height = formMain.ScaleHeight - 4 - pictSurface.Top
End Sub

Sub pictSurface_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
     Static drawing As Integer, x1!, y1!       'drawing is a flag; x1 and y1 are old positions
     If drawing = True Then
      pictSurface.DrawWidth = 1            'if user's already clicked once,
      pictSurface.Line (x1!, y1!)-(X, Y)   'draw the line
      drawing = False                      'and get a fresh start next click
     Else
      pictSurface.DrawWidth = 3            '....otherwise, start the line draw by:
      pictSurface.PSet (X, Y)              'make a dot where user is starting the line
      x1! = X                              'remember our starting coordinates
      y1! = Y
      drawing = True                       'and flag that we're drawing
     End If
     textCurrentX.Text = Format$(pictSurface.CurrentX, "####0.00")    'update display box
     textCurrentY.Text = Format$(pictSurface.CurrentY, "####0.00")    'update display box
End Sub

Sub pictSurface_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
     textMouseX.Text = Format$(X, "####0.00")               'update MouseX box
     textMouseY.Text = Format$(Y, "####0.00")               'update MouseY box
End Sub

