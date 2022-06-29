VERSION 2.00
Begin Form formScreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2028
   ClientLeft      =   5628
   ClientTop       =   4812
   ClientWidth     =   2988
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   14  'Merge Pen Not
   Height          =   2472
   Left            =   5568
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2028
   ScaleWidth      =   2988
   Top             =   4428
   Width           =   3108
   Begin Timer timeBlanker 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1080
   End
End

Sub Form_Click ()
    timeBlanker.Enabled = False
    Msg$ = "Would you like to save this image?"
    Title$ = "Graphics Project"
    Ans = MsgBox(Msg$, MB_YESNO, Title$)
    If Ans = 6 Then
        MousePointer = HOURGLASS
        'the savepicture line will not work if the \vb directory doesn't exist
        On Error Resume Next
        SavePicture formScreen.Image, "C:\VB\~TEMP.BMP"
        MousePointer = DEFAULT
    End If
    formScreen.Cls
    formScreen.Hide
    formGraphics.Show
End Sub

Sub Form_Load ()
    WindowState = MAXIMIZE
End Sub

Sub Form_Paint ()
    BackColor = RGB(255, 255, 255)' White background
End Sub

Sub timeBlanker_Timer ()
    Static Color As Integer 'stores color from previous timer triggering
    X = Int((ScaleWidth - (ScaleWidth / 20)) * Rnd + (ScaleWidth / 20))
    Y = Int((ScaleHeight - (ScaleHeight / 20)) * Rnd + (ScaleHeight / 20))
    DrawWidth = Int((100 - (20)) * Rnd + (20))
    Select Case BlankerType
        Case 0 'Blank screen
            formScreen.BackColor = QBColor(0) ' Black
        Case 1 'Spots
            formScreen.PSet (X, Y), QBColor(Color)
        Case 2 'Circles
            Radius = Int((ScaleWidth / 2 - (ScaleWidth / 20)) * Rnd + (ScaleWidth / 20))
            Circle (X, Y), Radius, QBColor(Color)
        Case 3 'Squares
            X1 = Int((ScaleWidth - (ScaleWidth / 20)) * Rnd + (ScaleWidth / 20))
            Y1 = Int((ScaleHeight - (ScaleHeight / 20)) * Rnd + (ScaleHeight / 20))
            Line (X, Y)-(X1, Y1), QBColor(Color), BF
        Case 4 'Lines
            X1 = Int((ScaleWidth - (ScaleWidth / 20)) * Rnd + (ScaleWidth / 20))
            Y1 = Int((ScaleHeight - (ScaleHeight / 20)) * Rnd + (ScaleHeight / 20))
            Line (X, Y)-(X1, Y1), QBColor(Color)
    End Select
    If BlankerType <> 0 Then
         If Color <> 15 Then
             Color = Color + 1
         Else
             Color = 0
             If Rnd > .6 Then BackColor = Point(CurrentX, CurrentY)
        End If
    End If
End Sub

