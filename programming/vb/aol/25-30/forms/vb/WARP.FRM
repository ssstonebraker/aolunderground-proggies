VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Warp Factor VB"
   ClientHeight    =   5820
   ClientLeft      =   1890
   ClientTop       =   2235
   ClientWidth     =   7365
   ClipControls    =   0   'False
   Height          =   6225
   Left            =   1830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Top             =   1890
   Width           =   7485
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4305
      Top             =   4020
   End
End
Option Explicit

' First we need to define two arrays, one to hold the X coordinates of the stars, one to hold
' the Y.
Dim nXCoord(50) As Integer
Dim nYCoord(50) As Integer

' Now we need variables to hold the stars X and Y Speeds
Dim nXSpeed(50) As Integer
Dim nYSpeed(50) As Integer

Sub Form_Load ()

    Dim nIndex As Integer

    ' At form load, the initial coordinates of all the stars needs to be set to off screen.
    ' The timer event will recognise this and bring the stars back on screen.

    For nIndex = 0 To 49

        nXCoord(nIndex) = -1
        nYCoord(nIndex) = -1

    Next

    ' Call the randomize method to tell VB to get ready to think of some random numbers
    Randomize

    Timer1.Enabled = True

End Sub

Sub Timer1_Timer ()

    ' The timer event performs three functions here.
    '       1. Stars that are off screen are remade at the centre of the screen
    '       2. Stars previously drawn are erase by redrawing them in black
    '       3. Each star's position is recalculated and the star redrawn.


    Dim nIndex As Integer


    For nIndex = 0 To 49

        'erase the previously drawn star
        PSet (nXCoord(nIndex), nYCoord(nIndex)), &H0&


        ' If the star number nIndex is off screen, then bring it back
        If nXCoord(nIndex) < 0 Or nXCoord(nIndex) > frmMain.ScaleWidth Or nYCoord(nIndex) < 0 Or nYCoord(nIndex) > frmMain.ScaleHeight Then

            nXCoord(nIndex) = frmMain.ScaleWidth \ 2
            nYCoord(nIndex) = frmMain.ScaleHeight \ 2

            ' Decide on some random speeds for the new star
            nXSpeed(nIndex) = Int(Rnd(1) * 200) - 100   ' Gives a speed between -100 and 100
            nYSpeed(nIndex) = Int(Rnd(1) * 200) - 100   ' Gives a speed between -100 and 100
        End If

        ' Now redraw the star so that it appears to move
        nXCoord(nIndex) = nXCoord(nIndex) + nXSpeed(nIndex)
        nYCoord(nIndex) = nYCoord(nIndex) + nYSpeed(nIndex)

        PSet (nXCoord(nIndex), nYCoord(nIndex)), &HFFFFFF

    ' Move on to the next star
    Next

End Sub

